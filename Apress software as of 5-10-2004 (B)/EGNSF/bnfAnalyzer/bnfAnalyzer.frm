VERSION 5.00
Begin VB.Form frmBNFanalyzer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backus-Naur Form Analyzer"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraParseStatus 
      Caption         =   "Parse Status"
      Height          =   1455
      Left            =   3360
      TabIndex        =   17
      Top             =   2040
      Width           =   1575
      Begin VB.OptionButton optParseStatusCompleteReport 
         Caption         =   "Complete report"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Select for a complete report on parsing (time-consuming for large files)"
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optParseStatusSimpleReport 
         Caption         =   "Simple report"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Select for a simple parse report"
         Top             =   600
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optParseStatusNoReport 
         Caption         =   "No report"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Select for no report on parsing"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CheckBox chkProgress 
      Caption         =   "Progress reports"
      Height          =   255
      Left            =   3360
      TabIndex        =   16
      ToolTipText     =   "Click to get extra progress reports (in the Status Reports box) as the input BNF is parsed"
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CommandButton cmdRegistry2Form 
      Caption         =   "Restore Settings"
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      ToolTipText     =   "Click to restore the settings you have selected from your Registry"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdForm2Registry 
      Caption         =   "Save Settings"
      Height          =   375
      Left            =   3360
      TabIndex        =   14
      ToolTipText     =   "Click to save the settings you have selected to your Registry for future use"
      Top             =   720
      Width           =   1575
   End
   Begin VB.ListBox lstTerminals 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   7560
      TabIndex        =   13
      ToolTipText     =   "Lists all the terminals in the parsed BNF"
      Top             =   2040
      Width           =   2535
   End
   Begin VB.ListBox lstNonterminals 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   5040
      TabIndex        =   12
      ToolTipText     =   "Lists all the nonterminal grammar categories in the parsed BNF"
      Top             =   2040
      Width           =   2535
   End
   Begin VB.ListBox lstStatus 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   5040
      TabIndex        =   10
      ToolTipText     =   $"bnfAnalyzer.frx":0000
      Top             =   360
      Width           =   5055
   End
   Begin VB.CommandButton cmdStatusZoom 
      Caption         =   "Zoom"
      Height          =   255
      Left            =   9120
      TabIndex        =   9
      ToolTipText     =   "Click for a zoomed and larger view of status and progress up to now"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      ToolTipText     =   "Will save your form selections and close the application"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdCreateReferenceManual 
      Caption         =   "Create Reference Manual"
      Default         =   -1  'True
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      ToolTipText     =   "Click to scan and parse the BNF as needed, and create the manual"
      Top             =   120
      Width           =   1575
   End
   Begin VB.FileListBox filBNFlocation 
      BackColor       =   &H8000000F&
      Height          =   1845
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Input BNF file: click once to select"
      Top             =   1920
      Width           =   3135
   End
   Begin VB.DirListBox dirBNFlocation 
      BackColor       =   &H8000000F&
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Folder where the input BNF file is located"
      Top             =   720
      Width           =   3135
   End
   Begin VB.DriveListBox drvBNFlocation 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Drive or server where the input BNF file is located"
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label lblStatusProgress 
      BackColor       =   &H00C0FFFF&
      Height          =   135
      Left            =   5040
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status Reports"
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label lblTerminals 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Terminals"
      Height          =   255
      Left            =   7560
      TabIndex        =   7
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label lblNonterminals 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nonterminals"
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label lblBNFlocation 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BNF Location"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileParse 
         Caption         =   "Parse the BNF in %FILETITLE"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsReferenceManual 
         Caption         =   "Create language reference manual"
      End
      Begin VB.Menu mnuToolsReferenceManualOptions 
         Caption         =   "Reference manual options"
      End
      Begin VB.Menu mnuToolsRRdiagram 
         Caption         =   "Create railroad diagram"
      End
      Begin VB.Menu mnuToolsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsMkParseTree 
         Caption         =   "Create the parse tree"
      End
      Begin VB.Menu mnuToolsParseBNF 
         Caption         =   "Parse the BNF"
      End
      Begin VB.Menu mnuToolsDestroyParseTree 
         Caption         =   "Destroy the parse treee"
      End
      Begin VB.Menu mnuToolsDumpParseTree 
         Caption         =   "Dump the parse tree"
      End
      Begin VB.Menu mnuToolsParseTree2XML 
         Caption         =   "Parse tree to XML"
      End
      Begin VB.Menu mnuToolsInspectParseTree 
         Caption         =   "Inspect the parse tree"
      End
      Begin VB.Menu mnuToolsViewBNF 
         Caption         =   "View the source BNF"
      End
      Begin VB.Menu mnuToolsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsListNonterminals 
         Caption         =   "List nonterminal symbols..."
      End
      Begin VB.Menu mnuToolsListTerminals 
         Caption         =   "List terminal symbols..."
      End
      Begin VB.Menu mnuToolsSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsDumpScanTable 
         Caption         =   "Dump scanTable..."
      End
      Begin VB.Menu mnuToolsInspectScanTable 
         Caption         =   "Inspect scanTable..."
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Visible         =   0   'False
      Begin VB.Menu mnuViewAndChange 
         Caption         =   "View and change file %FILEID"
      End
   End
End
Attribute VB_Name = "frmBNFanalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************************
' *                                                                   *
' * Backus-Naur Form Analyzer                                         *
' *                                                                   *
' *                                                                   *
' * This form and application parses files containing BNF and it      *
' * error-checks and analyzes the BNF.  It produces a reference       *
' * manual including a list of nonterminals, a list of terminals, and *
' * an outline format list of rules for forming elements of the       *
' * language, which can be a basis for a reference manual.            *
' *                                                                   *
' * The rest of this comment block describes the following:           *
' *                                                                   *
' *                                                                   *
' *      *  The BNF syntax, of the BNF syntax used in this application*
' *      *  Friendly methods exposed by this form                     *
' *      *  Implementation notes                                      *
' *      *  Compile-time symbols                                      *
' *      *  Change record                                             *
' *      *  Issues                                                    *
' *                                                                   *
' *                                                                   *
' * BNF OF THE BNF -------------------------------------------------- *
' *                                                                   *
' * This section identifies:                                          *
' *                                                                   *
' *                                                                   *
' *      *  The lexical (character by character) syntax of the        *
' *         expected Backus-Naur Form file                            *
' *                                                                   *
' *      *  The BNF of BNF                                            *
' *                                                                   *
' *                                                                   *
' * ----- Lexical syntax notes                                        *
' *                                                                   *
' * The BNF text file should be a stream of ASCII characters. White   *
' * space is ignored except inside quoted strings; note that the      *
' * reserved name WHITESPACE identifies white space as a terminal in  *
' * the language being defined.                                       *
' *                                                                   *
' * The input BNF "language" is completely insensitive to case.  How- *
' * ever, for best results, use the convention of camelCase for       *
' * nonterminals and UPPERCASE for names of terminals that you do not *
' * expand in your BNF.                                               *
' *                                                                   *
' * The following lexical elements are supported.                     *
' *                                                                   *
' *                                                                   *
' *      *  Completely blank lines, and lines that commence with an   *
' *         apostrophe, are treated as comments. Lines may also end   *
' *         in comments; any characters after the leftmost unquoted   *
' *         apostrophe, including the apostrophe, are treated as      *
' *         comments.                                                 *
' *                                                                   *
' *      *  Lines may be continued, simply by making sure the first   *
' *         character of continuation lines is a blank                *
' *                                                                   *
' *         Note in this connection that newlines may follow Windows  *
' *         or Web conventions, where the Windows convention is       *
' *         carriage return and linefeed while the Web convention is  *
' *         that a single linefeed will separate lines. These two     *
' *         conventions can appear, in fact, within the same file.    *
' *                                                                   *
' *      *  Identifiers follow Visual Basic 6 conventions: starting   *
' *         with a letter, they should contain letters, digits and    *
' *         the underscore exclusively. There is no limit to the      *
' *         length of identifiers, except common sense.               *
' *                                                                   *
' *         However and unlike Visual Basic identifiers, identifiers  *
' *         are completely case-sensitive as in the case of C++. The  *
' *         case of the first letter of the identifier shows its type:*
' *                                                                   *
' *         + Identifiers that start with a lower-case letter are     *
' *           assumed to be nonterminals of the grammar. These identi-*
' *           fiers must appear as defined on the left hand side of   *
' *           at least one production.                                *
' *                                                                   *
' *         + Identifiers that start with an upper-case letter are    *
' *           assumed to be "symbolic" grammar terminals (note that   *
' *           strings may also be terminals.) These identifiers may   *
' *           not appear on the left hand side of a production.       *
' *                                                                   *
' *           For best results, symbolic grammar terminals should be  *
' *           exclusively UPPERCASE: but they may also be Proper case,*
' *           with only the first character in upper case.            *
' *                                                                   *
' *      *  The := operator is the production operator which separates*
' *         a defined nonterminal from its definition                 *
' *                                                                   *
' *      *  The stroke operator | separates alternatives              *
' *                                                                   *
' *      *  The left and right square brackets [ and ] group optional *
' *         sequences on the right hand side of the := operator       *
' *                                                                   *
' *      *  The left and right parentheses ( and ) group sequences    *
' *         (for evaluation precedence) on the right hand side of the *
' *         := operator                                               *
' *                                                                   *
' *      *  The asterisk and the plus sign perform a role similar to  *
' *         the role they perform in regular expressions              *
' *                                                                   *
' *      *  Quoted strings follow VB conventions (double quotes are   *
' *         delimiters; internal double quotes must be doubled) and   *
' *         are used to specify exact character sequences.            *
' *                                                                   *
' *                                                                   *
' * ----- BNF of BNF                                                  *
' *                                                                   *
' * Note: this BNF should be available in the test file BNFBNF.TXT    *
' * that is shipped with this software. This means that you can       *
' * submit the grammar implemented by this application, to this       *
' * application. This is either elegant or confusing or both.         *
' *                                                                   *
' *                                                                   *
' *      bnfGrammar := production [ bnfGrammar ]                      *
' *      production := [ nonTerminal ":=" productionRHS ]             *
' *                    (NEWLINE|EOF)                                  *
' *      production := NEWLINE ' Allows for empty lines               *
' *      nonTerminal := IDENTIFIER                                    *
' *      productionRHS := sequenceFactor [ sequenceFactor ]           *
' *      sequenceFactor := mockRegularExpression                      *
' *                        [ alternationFactorRHS ]                   *
' *      mockRegularExpression := mreFactor [ mrePostfix ]            *
' *      mreFactor := nonTerminal |                                   *
' *                   terminal |                                      *
' *                   STRING |                                        *
' *                   "(" productionRHS ")" |                         *
' *                   "[" productionRHS "]"                           *
' *      mrePostfix := "*" | "-"                                      *
' *      alternationFactorRHS := "|" mockRegularExpression            *
' *                              [ alternationFactorRHS ]             *
' *                                                                   *
' *                                                                   *
' * In the above, the following upper-case identifiers are used with  *
' * the following meanings.                                           *
' *                                                                   *
' *                                                                   *
' *      *  EOF: end of file                                          *
' *                                                                   *
' *      *  IDENTIFIER: an identifier with VB-6 syntax                *
' *                                                                   *
' *      *  NEWLINE: environment newline: carriage return and line    *
' *         feed for Windows: line feed for the Web.                  *
' *                                                                   *
' *      *  STRING: a string with VB-6 syntax                         *
' *                                                                   *
' *                                                                   *
' * FRIENDLY METHODS EXPOSED BY THIS FORM --------------------------- *
' *                                                                   *
' * This form exposes the following Friend methods:                   *
' *                                                                   *
' *                                                                   *
' *      *  nonTerminal2NodeIndex(tree, nt) where tree is a parse tree*
' *         in the collection format described below, and nt is a     *
' *         nonterminal name, returns 0 or the nonterminal's index in *
' *         the parse tree.                                           *
' *                                                                   *
' *      *  index2Name(tree, index, type) where tree is a parse tree  *
' *         in the collection format described below, index is the    *
' *         terminal or nonterminal index, and type is True for a     *
' *         nonterminal or False for a terminal, returns the name of  *
' *         the terminal or nonterminal.                              *
' *                                                                   *
' *                                                                   *
' * IMPLEMENTATION NOTES -------------------------------------------- *
' *                                                                   *
' * Without apology this form and application is an object developed  *
' * using the Collection object and Variant data typing. That is be-  *
' * cause this software was developed rapidly in VB-6 during a period *
' * when my .Net laptop was in the pawnshop.                          *
' *                                                                   *
' * This form and application does avoid weak typing, and, in order   *
' * to avoid the problems in collections that represent complex data  *
' * structures, the EGN dump and inspect methodology has been used    *
' * both to display objects and to make sure they contain no errors...*
' * or as few errors as possible.                                     *
' *                                                                   *
' *                                                                   *
' * COMPILE-TIME SYMBOLS -------------------------------------------- *
' *                                                                   *
' * When the compile-time symbol SIXTY_DAY_EVALUATION is True, an     *
' * EXE can be produced which will work for sixty days from first     *
' * use.                                                              *
' *                                                                   *
' * When the compile-time symbol BNFANALYZER_RRDIAGRAM is True,       *
' * pictorial syntax railroad diagrams may be produced from the       *
' * input BNF. Note that this capability isn't described in the Apress*
' * book.                                                             *
' *                                                                   *
' * When the compile-time symbol BNFANALYZER_FILEVIEW is True,        *
' * right-clicking a file in the filelist will bring up the menu      *
' * option to view and change the file in a simple window. Note that  *
' * this capability isn't described in the Apress book.               *
' *                                                                   *
' *                                                                   *
' * C H A N G E   R E C O R D --------------------------------------- *
' *   DATE     PROGRAMMER     DESCRIPTION OF CHANGE                   *
' * --------   ----------     --------------------------------------- *
' * 04 16 03   Nilges         Started version 1                       *
' *                                                                   *
' * 04 26 03   Nilges         Finished version 1 and archived it to   *
' *                           c:\egnsf\bnfAnalyzer\archive\04262003\  *
' *                           version1.                               *
' *                                                                   *
' * 04 27 03   Nilges         Version 1.1                             *
' *                                                                   *
' *                           1.  Don't list undefined info for       *
' *                               terminals                           *
' *                                                                   *
' * 04 27 03   Nilges         Version 1.2                             *
' *                                                                   *
' *                           1.  Added SIXTY_DAY_EVALUATION          *
' *                                                                   *
' * 01 11 04   Nilges         Version 2                               *
' *                                                                   *
' *                           1.  Wrap the where used list            *
' *                           2.  If indentation becomes excessive    *
' *                               do not indent.                      *
' *                                                                   *
' * 01 24 04   Nilges         Version 3                               *
' *                                                                   *
' *                           1.  Indent and format the reference     *
' *                               manual                              *
' *                                                                   *
' *                           2.  Better control of line wrapping.    *
' *                                                                   *
' *                           3.  Added control box.                  *
' *                                                                   *
' *                           4.  Added ability to convert to XML     *
' *                                                                   *
' * 01 25 04   Nilges         Made utilities class a part of the      *
' *                           project, with extensions for this       *
' *                           project. It's unlikely the utilities    *
' *                           class will be used in new VB-6 applica- *
' *                           tions because it's unlikely I'll        *
' *                           develop goddamn new VB-6 applications.  *
' *                                                                   *
' * 01 27 04   Nilges         1.  XML formatting options added        *
' *                           2.  Tooltips added to all forms         *
' *                           3.  Close boxes added to all forms      *
' *                           4.  Don't reparse and rescan            *
' *                               unnecessarily                       *
' *                                                                   *
' * 01 28 04   Nilges         1.  Added railroad diagrams.            *
' *                                                                   *
' * 01 31 04   Nilges         1.  Added support for Web as well as    *
' *                               Windows newlines.                   *
' *                                                                   *
' * 02 20 04   Nilges         1.  Reactivated railroad diagramming and*
' *                               made it conditional on symbol       *
' *                               BNFANALYZER_RRDIAGRAM               *
' *                                                                   *
' *                           2.  Supporting view and change of files *
' *                               conditional on symbol               *
' *                               BNFANALYZER_FILEVIEW                *
' *                                                                   *
' * I S S U E S ----------------------------------------------------- *
' *   DATE       POSTER       DESCRIPTION OF ISSUE                    *
' * --------   ----------     --------------------------------------- *
' * 01 11 04   Nilges         Why does header of qb BNF say it's for  *
' *                           the "bnfAnalyzer language?"             *
' *                                                                   *
' *                           RESOLVED: this is because the name of   *
' *                           the language is entered in the          *
' *                           introductory screen for the reference   *
' *                           manual.                                 *
' *                                                                   *
' * 01 25 04   Nilges         Indentation of reference manual isn't   *
' *                           fully OK. Extra spaces appear before    *
' *                           words like "this" and extra indentation *
' *                           occurs. But not working on this further *
' *                           at this time because the XML extension  *
' *                           will provide Internet Explorer          *
' *                           formatting.                             *
' *                                                                   *
' * 01 25 04   Nilges         Initial version of XML output will not  *
' *                           include where used.                     *
' *                                                                   *
' * 01 26 04   Nilges         Ensure that the documented structure of *
' *                           the parse tree corresponds to its actual*
' *                           structure.                              *
' *                                                                   *
' * 01 26 04   Nilges         In some cases, the indentation of the   *
' *                           BNF is incorrect in that sequences and  *
' *                           alternatives are indented an extra 4    *
' *                           characters.                             *
' *                                                                   *
' * 01 26 04   Nilges         Add more options to control XML format  *
' *                           RESOLVED                                *
' *                                                                   *
' * 01 26 04   Nilges         Add support for XML indentation         *
' *                                                                   *
' * 01 30 04   Nilges         Use formatOutline as converted to VB-6  *
' *                           to format outline                       *
' *                                                                   *
' * 01 31 04   Nilges         Railroad diagram implementation isn't   *
' *                           complete.                               *
' *                                                                   *
' * 01 31 04   Nilges         Note that at times periods appear after *
' *                           Dewey outline numbers in the reference  *
' *                           manual, in an inconsistent fashion.     *
' *                                                                   *
' * 02 01 04   Nilges         The language reference manual for       *
' *                           "repeater" in the parse of the BNF for  *
' *                           regular expressions, may not be correct *
' *                                                                   *
' *********************************************************************

' ***** Scanned BNF ***************************************************
Private Enum ENUscannedType
    alternation             ' stroke
    comma                   ' comma
    comment                 ' comments
    nonterminalIdentifier   ' Mixed-case nonterminal identifier
    mreOperator             ' asterisk or plus
    newline                 ' Newline other than continuation line
    parenthesis             ' tokenValue identifies left or right,
                            ' round parenthesis versus square bracket
    productionAssignment    ' :=
    stringToken             ' tokenValue does include quotes
    terminalIdentifier      ' Upper-case nonterminal identifier
    unknown                 ' Indicates an error
End Enum
Private Type TYPscanned
    enuType As ENUscannedType
    lngStartIndex As Long
    lngLength As Long
    strTokenValue As String
End Type
Private USRscanned() As TYPscanned
' --- Inspection rules
Private Const SCANNED_INSPECTION_ISSOMETHING = _
    "The scan table must be an allocated array"
Private Const SCANNED_INSPECTION_TYPEVALID = _
    "Each entry must contain a valid token type (excluding comment)"
Private Const SCANNED_INSPECTION_STARTINDEXVALID = _
    "Each entry must contain a valid start index"
Private Const SCANNED_INSPECTION_LENGTHVALID = _
    "Each entry must contain a valid length"
Private Const SCANNED_INSPECTION_ORDER = _
    "The entries must be in ascending order"

' ***** Parsed BNF *****************************************************
' *                                                                    *
' * The parsed BNF is a treelike collection containing a list of the   *
' * nonterminals in item(1), a list of the terminals in item(2), and   *
' * parse tree nodes in the remaining entries. Each node that corre-   *
' * sponds to the complete parse of a production has a key equivalent  *
' * to the terminal on the production's left hand side; each node, that*
' * corresponds to a subexpression, has no key.                        *
' *                                                                    *
' * If the collection structure changes you must change the following  *
' * routines:                                                          *
' *                                                                    *
' *                                                                    *
' *      *  parseTree2Rules                                            *
' *      *  parseTree2XML                                              *
' *      *  frmRRdiagram.parseTree2RRDiagram                           *
' *                                                                    *
' *                                                                    *
' * The subcollection has the following structure; note that when the  *
' * the structure's requirements change you should change both this    *
' * comment, and the inspection rules below.                           *
' *                                                                    *
' *                                                                    *
' *      *  Item(1) is the nonterminal name for this entry or the      *
' *         null string for an anonymous entry.                        *
' *                                                                    *
' *         If a terminal appears on the left hand side of more than   *
' *         one production note that extra right hand sides are conso- *
' *         lidated into one entry, using the tree form of an          *
' *         alternation operator.                                      *
' *                                                                    *
' *      *  Item(2) is the self-referential Long index of the entry in *
' *         the parse collection                                       *
' *                                                                    *
' *      *  Item(3) is the root of the tree which represents the       *
' *         right hand side of the nonterminal's definition.  It       *
' *         contains a Variant of the following type.                  *
' *                                                                    *
' *         + A Long variant with a nonzero, positive value means that *
' *           the right hand side is a simple nonterminal, and it      *
' *           points to the definition, in the tree collection, of this*
' *           nonterminal.                                             *
' *                                                                    *
' *         + A Long variant with a value of -1 means that the right   *
' *           hand side is a simple terminal                           *
' *                                                                    *
' *         + A Long variant with a value of zero means the nonterminal*
' *           in Item(1) is undefined.                                 *
' *                                                                    *
' *         + A Collection means that the nonterminal expands into more*
' *           than a single terminal or nonterminal. This sub-col-     *
' *           lection will have the following items.                   *
' *                                                                    *
' *           - Item(1) will be the String major operator of a         *
' *             unary or binary subexpression                          *
' *                                                                    *
' *           - Item(2) will contain the index, in the main parse      *
' *             collection, of an "anonymous" entry.                   *
' *                                                                    *
' *             This entry will have a null string in item(1), its own *
' *             self-referential collection index in item(2), and the  *
' *             value of the first or only operand of the major opera- *
' *             tor in item 3.                                         *
' *                                                                    *
' *           - If the major operator in item(1) is a binary operator, *
' *             item(3), like item(2), will be the Long index of the   *
' *             anonymous expansion of the second operator.            *
' *                                                                    *
' *      *  Item(4) is the start index of the BNF code for the         *
' *         right hand side of the nonterminal definition: note that   *
' *         this is a pointer to the start character.                  *
' *                                                                    *
' *      *  Item(5) is the length in the BNF code of the right hand    *
' *         side of the nonterminal definition: note that this is a    *
' *         character count.                                           *
' *                                                                    *
' *                                                                    *
' * In addition, two unkeyed anonymous entries appear in items 1..2    *
' * of the parse collection:                                           *
' *                                                                    *
' *                                                                    *
' *      *  Item(1) of the parse collection is a subCollection,        *
' *         containing sub-sub-collections in each of its items.       *
' *                                                                    *
' *         + Item(1) of each sub-sub-collection identifies one non-   *
' *           terminal as a String                                     *
' *                                                                    *
' *         + Item(2) of each sub-sub-collection contains the entry    *
' *           index                                                    *
' *                                                                    *
' *         + Item(3) through item(n) are the Long indexes of each     *
' *           production in which the nonterminal is used.             *
' *                                                                    *
' *           Note that all the "start" symbols of the grammar (all the*
' *           symbols which do not appear on the right hand side of a  *
' *           production) can be identified by finding all             *
' *           subCollection members with a Count of 1.                 *
' *                                                                    *
' *         Each one of these entries will have a key equal to its     *
' *         data.                                                      *
' *                                                                    *
' *      *  Item(2) of the parse collection is also a subCollection,   *
' *         also containing sub-sub-collections in each of its items.  *
' *                                                                    *
' *         + Item(1) of each sub-sub-collection identifies one        *
' *           terminal as a String                                     *
' *                                                                    *
' *         + Item(2) of each sub-sub-collection contains the entry    *
' *           index                                                    *
' *                                                                    *
' *         + Item(3) through item(n) are the Long indexes of each     *
' *           production in which the terminal is used.                *
' *                                                                    *
' *         Each one of these entries will have a key equal to its     *
' *         data.                                                      *
' *                                                                    *
' *                                                                    *
' * Since the parseTree collection is a virtual object it is subject   *
' * to the EGN "dump and inspect" regime:                              *
' *                                                                    *
' *                                                                    *
' *      *  The menu item on the form that is labeled "Dump Parse Tree"*
' *         creates a dump, best viewed in a monospaced font, of the   *
' *         parseTree.                                                 *
' *                                                                    *
' *      *  The menu item on the form that is labeled "Inspect Parse   *
' *         Tree" checks the parseTree for valid structure.            *
' *                                                                    *
' *                                                                    *
' **********************************************************************
Private COLparseTree As Collection
' --- Inspection rules: note correspondence to comments above
Private Const PARSETREE_INSPECTION_ISSOMETHING = _
    "The parse collection must be Something and not Nothing"
Private Const PARSETREE_INSPECTION_MINCOUNT = _
    "The parse collection must contain two entries at minimum           "
Private Const PARSETREE_INSPECTION_NONTERMINAL_LIST = _
    "Item(1) of the parse collection is a subCollection, containing     " & vbNewLine & _
    "sub-sub-collections in each of its items.                          " & vbNewLine & vbNewLine & _
    "     *  Item(1) of each sub-sub-collection identifies one non-     " & vbNewLine & _
    "        terminal as a String                                       " & vbNewLine & vbNewLine & _
    "     *  Item(2) of each sub-sub-collection contains the entry index" & vbNewLine & vbNewLine & _
    "     *  Item(3) through item(n) are the Long indexes of each       " & vbNewLine & _
    "        (anonymous or keyed) definition in which the nonterminal   " & vbNewLine & _
    "        is used.                                                   "
Private Const PARSETREE_INSPECTION_TERMINAL_LIST = _
    "Item(2) of the parse collection is a also a subCollection,         " & vbNewLine & _
    "also containing sub-sub-collections in each of its items.          " & vbNewLine & vbNewLine & _
    "will be one of the following:                                      " & vbNewLine & vbNewLine & _
    "     *  Item(1) of each sub-sub-collection identifies one          " & vbNewLine & _
    "        terminal as a String                                       " & vbNewLine & vbNewLine & _
    "     *  Item(2) of each sub-sub-collection contains the entry index" & vbNewLine & vbNewLine & _
    "     *  Item(3) through item(n) are the Long indexes of each       " & vbNewLine & _
    "        production in which the terminal is used                   " & vbNewLine & vbNewLine & _
    "Each one of these entries will have a key equal to its data.       "
Private Const PARSETREE_INSPECTION_NODE_STRUCTURE1 = _
    "Items 3..Count of the parseTree are subCollections of this form:" & vbNewLine & vbNewLine & _
    "     *  Item 1 is the non-null key (terminal name) or a null       " & vbNewLine & _
    "        string for an anonymous expansion                          " & vbNewLine & vbNewLine & _
    "     *  Item(2) is the self-referential Long index of the entry in " & vbNewLine & _
    "        the parse collection                                       " & vbNewLine & vbNewLine & _
    "     *  Item(3) is a the root of the tree which represents the     " & vbNewLine & _
    "        right hand side of the nonterminal's definition.  It       " & vbNewLine & _
    "        contains a Variant of the following type.                  " & vbNewLine & vbNewLine
Private Const PARSETREE_INSPECTION_NODE_STRUCTURE2 = _
    "        + A Long variant with a nonzero, positive value means that " & vbNewLine & _
    "          the right hand side is a simple nonterminal, and it      " & vbNewLine & _
    "          points to the definition, in the tree collection, of this" & vbNewLine & _
    "          nonterminal.                                             " & vbNewLine & vbNewLine & _
    "        + A Long variant with a value of -1 means that the right   " & vbNewLine & _
    "          hand side is a simple terminal                           " & vbNewLine & vbNewLine & _
    "        + A Long variant with a value of zero means the nonterminal" & vbNewLine & _
    "          in Item(1) is undefined.                                 " & vbNewLine & vbNewLine & _
    "        + A Collection means that the nonterminal expands into more" & vbNewLine & _
    "          than a single terminal or nonterminal. This sub-col-     " & vbNewLine & _
    "          lection will have the following items.                   " & vbNewLine & vbNewLine & _
    "          - Item(1) will be the String major operator of a unary or" & vbNewLine & _
    "            binary subexpression                                   " & vbNewLine & vbNewLine & _
    "          - Item(2) will contain the index, in the main parse      " & vbNewLine & _
    "            collection, of an ""anonymous"" entry.                 " & vbNewLine & vbNewLine & _
    "            This entry will have a null string in item(1), its own " & vbNewLine & _
    "            self-referential collection index in item(2), and the  " & vbNewLine & _
    "            value of the first or only operand of the major opera- " & vbNewLine & _
    "            tor in item 3.                                         " & vbNewLine & vbNewLine & _
    "          - If the major operator in item(1) is a binary operator, " & vbNewLine & _
    "            item(3), like item(2), will be the Long index of the   " & vbNewLine & _
    "            anonymous expansion of the second operator.            " & vbNewLine & vbNewLine
Private Const PARSETREE_INSPECTION_NODE_STRUCTURE3 = _
    "     *  Item(4) is the start index of the BNF code for the         " & vbNewLine & _
    "        right hand side of the nonterminal definition: note that   " & vbNewLine & _
    "        this is a pointer to the scantable.                        " & vbNewLine & vbNewLine & _
    "     *  Item(5) is the length in the BNF code of the right hand    " & vbNewLine & _
    "        side of the nonterminal definition: note that this is a    " & vbNewLine & _
    "        count of scantable entries.                                "
Private Const PARSETREE_INSPECTION_NODE_STRUCTURE = _
              PARSETREE_INSPECTION_NODE_STRUCTURE1 & _
              PARSETREE_INSPECTION_NODE_STRUCTURE2 & _
              PARSETREE_INSPECTION_NODE_STRUCTURE3

' ***** Shared utilities ***********************************************
Private OBJutilities As clsUtilities

' ***** Constants ******************************************************
' --- Easter eggs
Private Const EGNADDRESS = _
    "Edward G. Nilges" & vbNewLine & _
    "spinoza1111@yahoo.COM" & vbNewLine & _
    "http://members.screenz.com/EdNilges"
Private Const ABOUTINFO = _
    "This form and application parses files containing BNF and it " & _
    "produces an analysis of the BNF definition, including a list " & _
    "of terminal symbols, a list of nonterminals, and at least the " & _
    "start of a reference manual for the language defined by the BNF." & _
    vbNewLine & vbNewLine & _
    "This form and application was developed starting on 4/17/2003 " & _
    "by:" & _
    vbNewLine & vbNewLine & _
    EGNADDRESS
#If SIXTY_DAY_EVALUATION Then
    Private Const SIXTYDAY_INFO = _
            "This evaluation version %ISWAS used for the first time on " & _
            "%EVALUATION_STARTDATE: it will expire on %EVALUATION_ENDDATE. " & _
            "Contact Edward G. Nilges at the above locations for information on obtaining " & _
            "an object or a source license for this code, together with support."
#End If
' --- Error handler captions
Private Const ERRORHANDLERCAPTION_RAISEERRORS = "Raise errors"
Private Const ERRORHANDLERCAPTION_DISPLAYTHISFORM = "Display this form"
Private Const ERRORHANDLERCAPTION_CANCEL = "Cancel"

#If SIXTY_DAY_EVALUATION Then
    ' ***** Sixty day encoded Registry keys ***************************
    Private STRencryptionKey As String
    Private STRstartDateKeyEncoded As String
#End If

' ***** Parse report options ******************************************
Private Enum ENUreportOptions
    noReport
    simpleReport
    completeReport
End Enum

' ***** Indentation max ***********************************************
Private Const INDENT_MAX = 4
Friend Function nonTerminal2NodeIndex(ByRef colTree As Collection, _
                                      ByVal strNonterminal As String) As Long
    '
    ' Converts name of the nonterminal to index of its definition node in parse tree,
    ' or 0
    '
    '
    Dim colHandle As Collection
    nonTerminal2NodeIndex = 0
    On Error Resume Next
        Set colHandle = colTree.item(symbol2Key(strNonterminal))
        If Err.Number <> 0 Then Exit Function
        nonTerminal2NodeIndex = colHandle.item(2)
    On Error GoTo 0
End Function

Private Sub cmdClose_Click()
    closer
End Sub

Private Sub cmdCreateReferenceManual_Click()
    referenceManualInterface
End Sub

Private Sub cmdForm2Registry_Click()
    form2Registry
End Sub
Private Sub cmdRegistry2Form_Click()
    registry2Form
End Sub
Private Sub cmdStatusZoom_Click()
    Dim intIndent As Integer
    Dim intWidth As Integer
    Dim lngIndex1 As Long
    Dim strIndent As String
    Dim strStatus As String
    strStatus = "STATUS REPORTS" & vbNewLine
    With lstStatus
        For lngIndex1 = 0 To .ListCount - 1
            intIndent = OBJutilities.verify(.List(lngIndex1) & "a", " ") - 1
            If intIndent < 10 Then
                intWidth = 72
            ElseIf intIndent < 20 Then
                intWidth = 50
            Else
                intWidth = 32
            End If
            strIndent = Mid$(.List(lngIndex1), 1, intIndent)
            strStatus = strStatus & _
                        vbNewLine & strIndent & _
                        Replace(OBJutilities.string2HardParagraph(.List(lngIndex1), _
                                                                  intWidth), _
                                vbNewLine, _
                                vbNewLine & strIndent & "   ")
        Next lngIndex1
    End With
    OBJutilities.errorHandler 0, _
                              strStatus, _
                              "cmdStatusZoom_Click", _
                              Me.Name
End Sub
Private Sub dirBNFlocation_Change()
    filBNFlocation.Path = dirBNFlocation
End Sub
Private Sub drvBNFlocation_Change()
    Dim strError As String
    strError = OBJutilities.driveServerChange(drvBNFlocation, _
                                              dirBNFlocation, _
                                              booAdvice:=True)
    If strError <> "" Then
        MsgBox strError: Exit Sub
    End If
End Sub
Private Sub filBNFlocation_Click()
    With filBNFlocation
        If .ListIndex > -1 Then
            toggleMenu .List(.ListIndex)
        End If
    End With
    Set COLparseTree = Nothing
End Sub

Private Sub filBNFlocation_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static strCaption As String
    Dim strFileid As String
    #If BNFANALYZER_FILEVIEW Then
        If Button = 2 Then
            With filBNFlocation
                If .ListIndex < 0 Then Exit Sub
                If strCaption = "" Then strCaption = mnuViewAndChange.Caption
                strFileid = OBJutilities.attachPath(dirBNFlocation.Path, .FileName)
                mnuViewAndChange.Caption = Replace(strCaption, _
                                                   "%FILEID", _
                                                   OBJutilities.enquote(strFileid))
                PopupMenu mnuView, , X + .Left, Y + .Top
                .Tag = strFileid
            End With
        End If
    #End If
End Sub

Private Sub filBNFlocation_PathChange()
    toggleMenu ""
End Sub
Private Sub Form_Load()
    Dim booDisplayAbout As Boolean
    Dim datInstallation As Date
    Dim strInstallation As String
    Dim strKeyFolder As String
    Dim strLockName As String
    #If SIXTY_DAY_EVALUATION Then
        ' Very crude way of enforcing evaluation copy
        strKeyFolder = encodeInterface("key", "Lookit dose jungle trees")
        STRencryptionKey = GetSetting(App.EXEName, Name, strKeyFolder, "")
        If STRencryptionKey = "" Then
            STRencryptionKey = genRandomKey
            If Not trialInstallationAnnunciate Then closer booFast:=True
            SaveSetting App.EXEName, Name, strKeyFolder, STRencryptionKey
            If Not trialInstallation(Name, STRencryptionKey) Then closer booFast:=True
        End If
        strLockName = encodeInterface("Lock", STRencryptionKey)
        If GetSetting(App.EXEName, Name, strLockName) = strLockName Then
            trialExpirationAnnunciate
            closer booFast:=True
        End If
        STRstartDateKeyEncoded = encodeInterface("installationDate", STRencryptionKey)
        On Error Resume Next
            datInstallation = CDate(decodeInterface(GetSetting(App.EXEName, _
                                                               Name, _
                                                               STRstartDateKeyEncoded, _
                                                               ""), _
                                                    STRencryptionKey))
        On Error GoTo 0
        If DateDiff("d", datInstallation, Now) >= 60 Then
            trialExpirationAnnunciate
            SaveSetting App.EXEName, _
                        Name, _
                        strLockName, _
                        strLockName
            closer booFast:=True
        End If
    #End If
    #If BNFANALYZER_RRDIAGRAM Then
        mnuToolsRRdiagram.Visible = True
    #Else
        mnuToolsRRdiagram.Visible = False
    #End If
    On Error GoTo Form_Load_Lbl1_createError
        Set OBJutilities = New clsUtilities
    On Error GoTo 0
    Show
    Refresh
    updateStatus "Loading", 1
    customizeEHF ERRORHANDLERCAPTION_RAISEERRORS, False, _
                 ERRORHANDLERCAPTION_DISPLAYTHISFORM, False, _
                 ERRORHANDLERCAPTION_CANCEL, False
    On Error Resume Next
        booDisplayAbout = GetSetting(App.EXEName, _
                                     Me.Name, _
                                     "displayAbout", _
                                     "True")
    On Error GoTo 0
    If booDisplayAbout Then
        displayAbout
        On Error Resume Next
            SaveSetting App.EXEName, _
                        Me.Name, _
                        "displayAbout", _
                        "False"
        On Error GoTo 0
    End If
    On Error Resume Next
        ChDir OBJutilities.attachPath(App.Path, "test Files")
    On Error GoTo 0
    registry2Form
    Load frmRefManOptions
    updateStatus "", -1: updateStatus "Load complete", 0
    With filBNFlocation
        If .ListIndex > -1 Then
            toggleMenu .List(.ListIndex)
        Else
            toggleMenu ""
        End If
    End With
    Exit Sub
Form_Load_Lbl1_createError:
    ownErrorHandler "Can't create utilities object: " & _
                    Err.Number & " " & Err.Description
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    closer
    End
End Sub

Private Sub mnuFileExit_Click()
    closer
End Sub
Private Sub mnuFileParse_Click()
    parseInterface
End Sub
Private Sub mnuHelpAbout_Click()
    OBJutilities.errorHandlerForm.loadDefaults App.EXEName, Me.Name & "_errorHandler"
    displayAbout
End Sub
Private Sub mnuToolsDestroyParseTree_Click()
    If parseTree_destroy(COLparseTree) Then
        MsgBox "The parse tree has been successfully deallocated"
    Else
        MsgBox "The parse tree has not been successfully deallocated"
    End If
End Sub
Private Sub mnuToolsDumpParseTree_Click()
    OBJutilities.errorHandler 0, _
                              Replace(parseTree_dump(COLparseTree), _
                                      "vbCollection", _
                                      vbNewLine & "vbCollection" & vbNewLine & "     "), _
                              "mnuToolsDumpParseTree_Click", _
                              Me.Name
End Sub
Private Sub mnuToolsDumpScanTable_Click()
    OBJutilities.errorHandler 0, _
                              scanDump(USRscanned), _
                              "mnuToolsDumpScanTable_Click", _
                              Me.Name
End Sub
Private Sub mnuToolsInspectParseTree_Click()
    Dim strReport As String
    If parseTree_inspect(COLparseTree, strReport) Then
        MsgBox "Inspection succeeded"
    Else
        MsgBox "Inspection failed"
    End If
    OBJutilities.errorHandler 0, _
                              strReport, _
                              "cmdInspectParseTree_Click", _
                              Me.Name
End Sub
Private Sub mnuToolsInspectScanTable_Click()
    Dim strReport As String
    If scanInspect(strReport) Then
        MsgBox "Scan table inspection has succeeded"
    Else
        MsgBox "Scan table inspection has failed"
    End If
    OBJutilities.errorHandler 0, _
                              strReport, _
                              "mnuToolsInspectScanTable_Click", _
                              Me.Name
End Sub
Private Sub mnuToolsListNonterminals_Click()
    OBJutilities.errorHandler 0, _
                              listNonterminals, _
                              "mnuToolsListNonterminals_Click", _
                              Me.Name
End Sub
Private Sub mnuToolsListTerminals_Click()
    OBJutilities.errorHandler 0, _
                              listTerminals, _
                              "mnuToolsListTerminals_Click", _
                              Me.Name
End Sub
Private Sub mnuToolsMkParseTree_Click()
    If mkParseTree Then
        MsgBox "Parse tree has been created successfully"
    Else
        MsgBox "Parse tree has not been created successfully"
    End If
End Sub

Private Sub mnuToolsParseBNF_Click()
    parseInterface
End Sub

Private Sub mnuToolsParseTree2XML_Click()
    OBJutilities.errorHandler 0, _
                              parseTree2XML(COLparseTree, vbNewLine, False, False), _
                              "mnuToolsParseTree2XML_Click", _
                              Name
End Sub

Private Sub mnuToolsReferenceManual_Click()
    mkReferenceManual
End Sub
Private Sub mnuToolsReferenceManualOptions_Click()
    With frmRefManOptions
        .Show vbModal
    End With
End Sub
Private Function about() As String
    '
    ' Return information about this form and application
    '
    '
    about = ABOUTINFO
End Function
Private Function camelCase2ProperCase(ByVal strCamelCase As String) As String
    '
    ' camelCase to Proper Case: insert blanks before the upper case letters
    ' and the digits, and capitalize the first letter
    '
    '
    Dim lngIndex1 As Long
    Dim lngIndex2 As Long
    Dim strProperCase As String
    If strCamelCase = "" Then Exit Function
    strProperCase = UCase$(Mid(strCamelCase, 1, 1))
    lngIndex1 = 2
    Do While lngIndex1 <= Len(strCamelCase)
        lngIndex2 = OBJutilities.verify(strCamelCase & "A", _
                                        "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", _
                                        intStartIndex:=lngIndex1, _
                                        booMatch:=True)
        strProperCase = strProperCase & _
                        Mid$(strCamelCase, _
                             lngIndex1, _
                             lngIndex2 - lngIndex1)
        If lngIndex1 > Len(strCamelCase) Then Exit Do
        strProperCase = strProperCase & " " & _
                        Mid(strCamelCase, lngIndex2, 1)
        lngIndex1 = lngIndex2 + 1
    Loop
    camelCase2ProperCase = strProperCase
End Function
Private Sub closer(Optional ByVal booFast As Boolean = False)
    '
    ' Get rid of test object: close the application
    '
    '
    If Not (COLparseTree Is Nothing) Then parseTree_destroy COLparseTree
    If Not booFast Then
        updateStatus "Closing", 1
        form2Registry
        updateStatus "", -1
        updateStatus "Terminating", 0
    End If
    Unload Me
    End
End Sub
Private Sub collectionErrorHandler(ByVal lngErr As Long, _
                                   ByVal strErr As String, _
                                   ByVal strProcedure As String)
    '
    ' Handle generic collection errors
    '
    '
    OBJutilities.errorHandler 0, _
                              "Collection error: " & lngErr & " " & strErr, _
                              strProcedure, _
                              Me.Name
End Sub
Private Sub customizeEHF(ParamArray varCaptions())
    '
    ' Customise the error handler form
    '
    '
    ' This method conceals and shows Captioned controls on the error handler form.
    '
    ' The varCaptions parameter array should have this format:
    '
    '
    '      <caption>, True|False, <caption>, ...
    '
    '
    ' <caption> is the case-insensitive caption property of one control. If <caption>
    ' is followed by True, then the control is made Visible: if <caption> is followed
    ' by False, then the control is hidden.
    '
    '
    Dim booVisible As Boolean
    Dim ctlHandle As Control
    Dim lngIndex1 As Long
    Dim lngIndex2 As Long
    Dim strCaption As String
    Dim strCaptionKey As String
    Dim strCaptionNext As String
    If UBound(varCaptions) Mod 2 = 0 Then
        ownErrorHandler "Internal programming error in calling customizeEHF: " & _
                        "an even number of ParamArray entries is required"
        End
    End If
    For lngIndex1 = LBound(varCaptions) To UBound(varCaptions) Step 2
        On Error GoTo customizeEHF_Lbl1_varTypeErrorHandler
            strCaption = Trim$(UCase$(varCaptions(lngIndex1)))
            booVisible = varCaptions(lngIndex1 + 1)
        On Error GoTo 0
        Set ctlHandle = errorHandlerCaption2Ctl(strCaption)
        ctlHandle.Visible = booVisible
    Next lngIndex1
    Exit Sub
customizeEHF_Lbl1_varTypeErrorHandler:
    ownErrorHandler ("Programming error in calling customizeEHF: vartype incorrect")
    End
End Sub
Private Function decodeInterface(ByVal strInstring As String, _
                                 ByVal strKey As String) As String
    '
    ' Decode interface
    '
    '
    decodeInterface = OBJutilities.decode(OBJutilities.hex2String(strInstring), strKey)
End Function
Private Sub displayAbout()
    '
    ' Interface to the about display
    '
    '
    Dim strAbout As String
    strAbout = ABOUTINFO
    #If SIXTY_DAY_EVALUATION Then
        strAbout = strAbout & _
                   vbNewLine & vbNewLine & _
                   Replace(Replace(Replace(SIXTYDAY_INFO, "%ISWAS", "was"), _
                                   "%EVALUATION_ENDDATE", _
                                   getEvaluationEndDate), _
                           "%EVALUATION_STARTDATE", _
                           getEvaluationStartDate)
    #End If
    OBJutilities.errorHandler 0, _
                              strAbout, _
                              "displayAbout", _
                              Me.Name
End Sub
Private Function encodeInterface(ByVal strInstring As String, _
                                 ByVal strKey As String) As String
    '
    ' Encode interface: use hex digits to make sure that string is displayable
    '
    '
    encodeInterface = OBJutilities.string2Hex(OBJutilities.encode(strInstring, strKey))
End Function
Private Function errorHandlerCaption2Ctl(ByVal strCaption As String) As Control
    '
    ' Given the Caption of a control on the errhandler form: return the control
    '
    '
    Static colIndex As Collection           ' Key: caption: value: index
    Dim lngIndex1 As Long
    Dim strCaptionNext As String
    Dim strCaptionWork As String
    Dim strKey As String
    Set errorHandlerCaption2Ctl = Nothing
    strCaptionWork = Trim$(UCase$(strCaption))
    If (colIndex Is Nothing) Then
        ' --- First-time logic to create the controlIndex
        On Error GoTo errorHandlerCaption2Ctl_Lbl1_indexErrorHandler
            Set colIndex = New Collection
        On Error GoTo 0
    End If
    lngIndex1 = -1: strKey = symbol2Key(strCaptionWork)
    On Error Resume Next
        lngIndex1 = colIndex.item(strKey)
    On Error GoTo 0
    If lngIndex1 = -1 Then
        For lngIndex1 = 0 To OBJutilities.errorHandlerForm.Controls.Count - 1
            With OBJutilities.errorHandlerForm.Controls(lngIndex1)
                strCaptionNext = ""
                On Error Resume Next
                    strCaptionNext = Trim$(UCase$(.Caption))
                On Error GoTo 0
                If strCaptionNext <> "" And strCaptionNext = strCaptionWork Then
                    On Error GoTo errorHandlerCaption2Ctl_Lbl1_indexErrorHandler
                        colIndex.Add lngIndex1, strKey
                        Exit For
                    On Error GoTo 0
                End If
            End With
        Next lngIndex1
    End If
    If lngIndex1 >= OBJutilities.errorHandlerForm.Controls.Count Then
        ownErrorHandler "Internal programming error: " & _
                        "invalid caption " & _
                        OBJutilities.enquote(strCaption)
        End
    End If
    Set errorHandlerCaption2Ctl = OBJutilities.errorHandlerForm.Controls(lngIndex1)
    Exit Function
errorHandlerCaption2Ctl_Lbl1_indexErrorHandler:
    ownErrorHandler ("Can't create error handler control index: " & _
                     Err.Number & " " & Err.Description)
    End
End Function
Private Sub form2Registry()
    '
    ' Save form settings
    '
    '
    updateStatus "Saving form settings"
    On Error GoTo form2Registry_Lbl1_errorHandler
        SaveSetting App.EXEName, Me.Name, chkProgress.Name, chkProgress
        SaveSetting App.EXEName, Me.Name, optParseStatusNoReport.Name, optParseStatusNoReport
        SaveSetting App.EXEName, Me.Name, optParseStatusSimpleReport.Name, optParseStatusSimpleReport
        SaveSetting App.EXEName, Me.Name, optParseStatusCompleteReport.Name, optParseStatusCompleteReport
        SaveSetting App.EXEName, Me.Name, drvBNFlocation.Name, drvBNFlocation
        SaveSetting App.EXEName, Me.Name, dirBNFlocation.Name, dirBNFlocation
        With filBNFlocation
            SaveSetting App.EXEName, Me.Name, filBNFlocation.Name, .List(.ListIndex)
        End With
    On Error GoTo 0
    Exit Sub
form2Registry_Lbl1_errorHandler:
    OBJutilities.errorHandler 0, _
                              "Could not save form settings: " & _
                              Err.Number & " " & Err.Description, _
                              "form2Registry", _
                              Me.Name
End Sub
Private Function genRandomKey() As String
    '
    ' Generate a random key of 16 characters
    '
    '
    Dim bytIndex1 As Integer
    Dim strKey As String
    For bytIndex1 = 1 To 16
        strKey = strKey & ChrW$(CInt(Rnd * 255))
    Next bytIndex1
    genRandomKey = strKey
End Function
Private Function getReportOption() As ENUreportOptions
    '
    ' Get parse report option from the screen...quietly repair screen if needed
    '
    '
    If optParseStatusNoReport Then
        getReportOption = noReport
    ElseIf optParseStatusSimpleReport Then
        getReportOption = simpleReport
    ElseIf optParseStatusCompleteReport Then
        getReportOption = completeReport
    Else
        optParseStatusSimpleReport = True
        getReportOption = getReportOption
    End If
End Function
Private Function getStartSymbols(ByRef colParse As Collection, _
                                 ByRef colStartSymbols As Collection) As Boolean
    '
    ' Get nonterminal start symbols
    '
    '
    ' The start symbols are returned as a string collection
    '
    '
    Dim colHandle(1 To 2) As Collection
    Dim lngIndex1 As Long
    getStartSymbols = False
    On Error GoTo getStartSymbols_Lbl1_errorHandler
        If Not (colStartSymbols Is Nothing) Then Set colStartSymbols = Nothing
        Set colStartSymbols = New Collection
    On Error GoTo 0
    Set colHandle(1) = colParse.item(1)
    With colHandle(1)
        For lngIndex1 = 1 To .Count
            Set colHandle(2) = .item(lngIndex1)
            With colHandle(2)
                If .Count < 3 Then
                    On Error GoTo getStartSymbols_Lbl1_errorHandler
                        colStartSymbols.Add .item(1)
                    On Error GoTo 0
                End If
            End With
        Next lngIndex1
    End With
    getStartSymbols = True
    Exit Function
getStartSymbols_Lbl1_errorHandler:
    collectionErrorHandler Err.Number, Err.Description, "getStartSymbols"
End Function
Private Function inCollection(ByRef colCollection As Collection, _
                              ByVal strKey As String) As Boolean
    '
    ' Tell caller whether key is defined in collection
    '
    '
    Dim intErr As Integer
    Dim varValue As Variant
    inCollection = False
    On Error Resume Next
        varValue = colCollection.item(strKey): intErr = Err.Number
    On Error GoTo 0
    inCollection = (intErr = 0)
End Function
Private Function indentation(ByVal strInstring As String) As Integer
    '
    ' Return the indentation of a line that starts with a Dewey number
    '
    '
    Dim intIndex1 As Integer
    For intIndex1 = 1 To Len(strInstring)
        If OBJutilities.verify(UCase(Mid(strInstring, intIndex1, 1)), _
                                "ABCDEFGHIJKLMNOPQRSTUVWXYZ") _
           = _
           0 Then
            Exit For
        End If
    Next intIndex1
    If intIndex1 <= Len(strInstring) Then indentation = intIndex1 - 1
    Exit Function
indentation_Lbl1_splitErrorHandler:
    OBJutilities.errorHandler Err.Number, Err.Description, "indentation", Name
End Function
Private Function indentContinuations(ByVal strInstring As String) As String
    '
    ' Indent continuation lines in the reference manual
    '
    '
    Dim intIndex1 As Integer
    Dim strIndent As String
    Dim strOutstring As String
    strOutstring = OBJutilities.item(strInstring, 1, vbNewLine, booSetDelimiter:=False)
    strIndent = OBJutilities.copies(" ", indentation(strOutstring))
    For intIndex1 = 2 To OBJutilities.items(strInstring, vbNewLine, booSetDelimiter:=False)
        strOutstring = strOutstring & _
                       vbNewLine & _
                       strIndent & _
                       OBJutilities.item(strInstring, intIndex1, vbNewLine, booSetDelimiter:=False)
    Next intIndex1
    indentContinuations = strOutstring
End Function
Private Function inspectAppend(ByRef strReport As String, _
                                          ByVal strRule As String, _
                                          ByVal booInspectResult As Boolean, _
                                          ByRef booInspection As Boolean, _
                                          Optional ByVal strComment As String = "", _
                                          Optional ByVal booBox As Boolean = True) As Boolean
    '
    ' Appends to an inspection report
    '
    '
    Dim strCommentWork As String
    Dim strRuleDisplay As String
    inspectAppend = False
    If Trim$(strComment) <> "" Then
        strCommentWork = vbNewLine & _
                         OBJutilities.string2HardParagraph(strComment, 50)
    End If
    strRuleDisplay = strRule & _
                     vbNewLine & _
                     IIf(InStr(strRule, vbNewLine) <> 0, vbNewLine, "") & _
                     "Rule application has " & _
                     IIf(booInspectResult, "succeeded", "FAILED!") & _
                     strCommentWork & _
                     vbNewLine
    If booBox Then
        strRuleDisplay = OBJutilities.string2Box(strRuleDisplay)
    End If
    strReport = strReport & vbNewLine & vbNewLine & strRuleDisplay
    booInspection = booInspectResult
    inspectAppend = booInspection
End Function
Private Function isIdentifier(ByVal strInstring As String) As Boolean
    '
    ' Return True if strInstring is of VB identifier format
    '
    '
    Dim strInstringWork As String
    strInstringWork = UCase$(strInstring)
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", Mid$(strInstringWork, 1, 1)) = 0 Then Exit Function
    isIdentifier = OBJutilities.verify(strInstringWork, _
                                     "ABCDEFGHIJKLMNOPQRSTUVWXYZ_0123456789") _
                   = _
                   0
End Function
Private Function isNonTerminal(ByVal strSymbol As String) As Boolean
    '
    ' Tell caller whether symbol is a nonterminal symbol
    '
    '
    isNonTerminal = scanTokenType2String(string2ScanTokenType(strSymbol)) _
                    = _
                    scanTokenType2String(nonterminalIdentifier)
End Function
Private Function isTerminal(ByVal strSymbol As String) As Boolean
    '
    ' Tell caller whether symbol is a terminal symbol: a quoted
    ' string or upper case symbol
    '
    '
    Dim enuType As ENUscannedType
    isTerminal = False
    enuType = string2ScanTokenType(strSymbol)
    If enuType = stringToken Or enuType = terminalIdentifier Then
        isTerminal = True
    End If
End Function
Private Function languageReference(ByRef colTree As Collection, _
                                   ByRef usrScan() As TYPscanned, _
                                   ByVal booBoxComments As Boolean, _
                                   ByVal booIncludeBNF As Boolean, _
                                   ByVal booIncludeNonterminals As Boolean, _
                                   ByVal booIncludeSyntax As Boolean, _
                                   ByVal booIncludeTerminals As Boolean, _
                                   ByVal booIncludeWhere As Boolean, _
                                   ByVal strName As String, _
                                   ByVal booIndentReferenceManual As Boolean, _
                                   ByVal booXML As Boolean, _
                                   ByVal booXMLnewlines As Boolean, _
                                   ByVal booInlineBNF As Boolean, _
                                   ByVal booEndTagComment As Integer) As String
    '
    ' Return the language reference manual
    '
    '
    ' Note that this method has suffered parameter bloat in consequence of an early
    ' decision (or lack of a conscious decision) on my part to produce a proprietary
    ' and geeky, mono-spaced text language reference manual, rather than XML ab initio
    ' and from the start. XML was added later with the result that this code is rather
    ' schizophrenic as it must choose its "modes".
    '
    ' But in my defense I discover few cheap or free tools for converting XML to
    ' documents.
    '
    '
    Dim strBackMatter As String
    Dim strFrontMatter As String
    Dim strOutstring(1 To 3) As String
    Dim strXMLnewline As String
    If (colTree Is Nothing) Then
        strOutstring(1) = "Cannot produce a reference manual: no language exists"
        If booXML Then strOutstring(1) = OBJutilities.mkXMLComment(strOutstring(1))
        languageReference = strOutstring(1)
        Exit Function
    End If
    If booXML Then
        If booXMLnewlines Then strXMLnewline = vbNewLine
    End If
    If booIncludeNonterminals Then
        If booXML Then
            strOutstring(1) = symbols2XML(colTree, 1, strXMLnewline)
        Else
            strOutstring(1) = vbNewLine & _
                              "The following are the nonterminal " & _
                              "symbols of the language" & _
                              vbNewLine & vbNewLine & _
                              listNonterminals(booHeader:=False) & _
                              vbNewLine
            If booBoxComments Then
                strOutstring(1) = OBJutilities.string2Box(strOutstring(1), _
                                                          strBoxLabel:="NONTERMINAL SYMBOLS")
            Else
                strOutstring(1) = "NONTERMINAL SYMBOLS" & vbNewLine & vbNewLine & _
                                  strOutstring(1)
            End If
        End If
    End If
    If booIncludeTerminals Then
        If booXML Then
            strOutstring(2) = symbols2XML(colTree, 2, strXMLnewline)
        Else
            strOutstring(2) = vbNewLine & _
                              "The following are the terminal " & _
                              "elements of the language" & _
                              vbNewLine & vbNewLine & _
                              listTerminals(booHeader:=False) & _
                              vbNewLine
            If booBoxComments Then
                strOutstring(2) = OBJutilities.string2Box(strOutstring(2), _
                                                          strBoxLabel:="TERMINAL SYMBOLS")
            Else
                strOutstring(2) = "TERMINAL SYMBOLS" & vbNewLine & vbNewLine & _
                                  strOutstring(2)
            End If
        End If
    End If
    If booIncludeSyntax Then
        If booXML Then
            strOutstring(3) = "<bnfProductions>" & _
                              strXMLnewline & _
                              Replace(parseTree2XML(COLparseTree, _
                                                    strXMLnewline, _
                                                    booInlineBNF, _
                                                    booEndTagComment), _
                                      vbNewLine, _
                                      strXMLnewline) & _
                              strXMLnewline & _
                              "</bnfProductions>"
        Else
            strOutstring(3) = vbNewLine & _
                              "The following are the rules " & _
                              "of the language" & _
                              vbNewLine & vbNewLine & _
                              parseTree2Rules(colTree, _
                                              usrScan, _
                                              booIncludeBNF, _
                                              booIncludeWhere, _
                                              booIndentReferenceManual) & _
                              vbNewLine
            If booBoxComments Then
                strOutstring(3) = OBJutilities.string2Box(strOutstring(3), _
                                                          strBoxLabel:="LANGUAGE SYNTAX")
            Else
                strOutstring(3) = "LANGUAGE SYNTAX" & vbNewLine & vbNewLine & _
                                  strOutstring(3)
            End If
        End If
    End If
    If booXML Then
        languageReference = xmlHeader & _
                            strXMLnewline & _
                            Replace(strOutstring(1), vbNewLine, strXMLnewline) & _
                            strXMLnewline & _
                            Replace(strOutstring(2), vbNewLine, strXMLnewline) & _
                            strXMLnewline & _
                            Replace(strOutstring(3), vbNewLine, strXMLnewline) & _
                            IIf(booXMLnewlines, vbNewLine, "") & _
                            xmlTrailer
        Exit Function
    End If
    languageReference = strOutstring(1) & _
                        vbNewLine & vbNewLine & vbNewLine & _
                        strOutstring(2) & _
                        vbNewLine & vbNewLine & vbNewLine & _
                        strOutstring(3)
End Function
Private Function listNonterminals(Optional ByVal booHeader As Boolean) As String
    '
    ' List nonterminal symbols
    '
    '
    If (COLparseTree Is Nothing) Then
        listNonterminals = "No parse tree exists: cannot list nonterminals"
        Exit Function
    End If
    listNonterminals = listSymbols(COLparseTree.item(1), _
                                   True, _
                                   COLparseTree, _
                                   booHeader)
End Function
Private Function listSymbols(ByRef colSymbols As Collection, _
                             ByVal booNonterminals As Boolean, _
                             ByRef colParse As Collection, _
                             ByVal booHeader As Boolean) As String
    '
    ' List nonterminal or terminal symbols
    '
    '
    Dim booDefined As Boolean
    Dim bytSequence As Byte
    Dim colHandle(1 To 2) As Collection
    Dim colRemarks As Collection
    Dim colSortedIndex As Collection
    Dim colWhereUsed As Collection
    Dim intMaxSymbol As Integer
    Dim intMaxRemarks As Integer
    Dim intMaxWhereUsed As Integer
    Dim lngIndex1 As Long
    Dim lngIndex2 As Long
    Dim lngUndefined As Long
    Dim strContinuation As String
    Dim strIndent As String
    Dim strOutstring As String
    Dim strPrefix As String
    Dim strRemark As String
    Dim strStartSymbolInfo As String
    Dim strWhereUsed As String
    listSymbols = ""
    On Error GoTo listSymbols_Lbl1_errorHandler
        Set colRemarks = New Collection
        Set colWhereUsed = New Collection
    On Error GoTo 0
    Set colHandle(1) = COLparseTree.item(IIf(booNonterminals, 1, 2))
    ' --- Sort the symbols
    If Not mkCollectionSortIndex(colHandle(1), colSortedIndex, Nothing) Then
        Exit Function
    End If
    ' --- Pass one: determine column widths and create column data
    ' Rest of list
    With colSortedIndex
        booDefined = True
        For lngIndex1 = 1 To .Count
            Set colHandle(2) = colHandle(1).item(colSortedIndex.item(lngIndex1))
            booDefined = True
            strWhereUsed = ""
            strWhereUsed = OBJutilities.string2HardParagraph _
                           (parseTree_whereList(COLparseTree, _
                                                colHandle(2).item(1), _
                                                False), _
                            intWidth:=50)
            If booNonterminals Then
                booDefined = nonTerminal2NodeIndex(COLparseTree, _
                                                             colHandle(2).item(1)) _
                             <> _
                             0
            End If
            strRemark = ""
            If Not booDefined Then
                strRemark = "Undefined"
            End If
            If booNonterminals _
               And _
               strWhereUsed = "" _
               And _
               Not parseTree_isExtensionName(colHandle(2).item(1), _
                                             strPrefix, _
                                             bytSequence) Then
                OBJutilities.append strRemark, _
                                    ": ", _
                                    "Start symbol", _
                                    booAppendInPlace:=True
            End If
            listSymbols_update_pass1Info colHandle(2).item(1), _
                                              colWhereUsed, _
                                              strWhereUsed, _
                                              colRemarks, _
                                              strRemark, _
                                              intMaxSymbol, _
                                              intMaxWhereUsed, _
                                              intMaxRemarks
            If Not booDefined Then lngUndefined = lngUndefined + 1
        Next lngIndex1
    End With
    ' --- Pass 2: create the list
    With OBJutilities
        intMaxSymbol = .max(intMaxSymbol, IIf(booNonterminals, 11, 8))
        intMaxWhereUsed = .max(intMaxWhereUsed, 10)
        intMaxRemarks = .max(intMaxRemarks, 7)
        strOutstring = IIf(booHeader, _
                           IIf(booNonterminals, "NONTERMINAL", "TERMINAL") & " " & _
                           "LIST AS OF " & Now & _
                           vbNewLine & vbNewLine & vbNewLine, _
                           "") & _
                       IIf(booNonterminals, _
                           strStartSymbolInfo & vbNewLine & _
                           "Number of undefined symbols: " & lngUndefined & " " & _
                           IIf(lngUndefined <> 0, "*** Error ***", "") & _
                           vbNewLine & vbNewLine, _
                           "") & _
                       .align(IIf(booNonterminals, "Nonterminal", "Terminal"), _
                              intMaxSymbol, _
                              enuAlignment:=alignCenter) & " " & _
                       .align("Where Used", _
                              intMaxWhereUsed, _
                              enuAlignment:=alignCenter) & " " & _
                       .align("Remarks", _
                              intMaxRemarks, _
                              enuAlignment:=alignCenter) & _
                       vbNewLine
        strOutstring = strOutstring & _
                       String$(intMaxSymbol, "-") & " " & _
                       String$(intMaxWhereUsed, "-") & " " & _
                       String$(intMaxRemarks, "-")
    End With
    With colHandle(1)
        strIndent = OBJutilities.copies(" ", intMaxSymbol + 1)
        For lngIndex1 = 1 To .Count
            Set colHandle(2) = colHandle(1).item(colSortedIndex.item(lngIndex1))
            strContinuation = ""
            lngIndex2 = InStr(colWhereUsed.item(lngIndex1), vbNewLine)
            If lngIndex2 <> 0 Then
                strContinuation = Replace(Mid(colWhereUsed.item(lngIndex1), lngIndex2), _
                                          vbNewLine, _
                                          vbNewLine & strIndent)
            End If
            strOutstring = strOutstring & vbNewLine & _
                           OBJutilities.align(colHandle(2).item(1), intMaxSymbol) & " " & _
                           OBJutilities.align(OBJutilities.item(colWhereUsed.item(lngIndex1), _
                                                                1, _
                                                                vbNewLine, _
                                                                False), _
                                              intMaxWhereUsed) & " " & _
                           strContinuation & " " & _
                           OBJutilities.align(colRemarks.item(lngIndex1), intMaxRemarks)
        Next lngIndex1
    End With
    listSymbols = strOutstring
    Exit Function
listSymbols_Lbl1_errorHandler:
    collectionErrorHandler Err.Number, Err.Description, "listSymbols"
End Function
Private Sub listSymbols_update_pass1Info(ByVal strNonterminal As String, _
                                         ByRef colWhereUsed As Collection, _
                                         ByVal strWhereUsed As String, _
                                         ByRef colRemarks As Collection, _
                                         ByVal strRemark As String, _
                                         ByRef intMaxNonterminal As Integer, _
                                         ByRef intMaxWhereUsed As Integer, _
                                         ByRef intMaxRemarks As Integer)
    '
    ' Update where-used and remarks collections, maximum lengths
    '
    '
    On Error GoTo listNonterminals_update_pass1Info_Lbl1_errorHandler
        colWhereUsed.Add strWhereUsed
        colRemarks.Add strRemark
    On Error GoTo 0
    intMaxNonterminal = OBJutilities.max(intMaxNonterminal, Len(strNonterminal))
    intMaxWhereUsed = OBJutilities.max(intMaxWhereUsed, maxLineLength(strWhereUsed))
    intMaxRemarks = OBJutilities.max(intMaxRemarks, Len(strRemark))
    Exit Sub
listNonterminals_update_pass1Info_Lbl1_errorHandler:
    collectionErrorHandler Err.Number, Err.Description, _
                           "listNonterminals_update_pass1Info_Lbl1_errorHandler"
End Sub
Private Function listTerminals(Optional ByVal booHeader As Boolean) As String
    '
    ' List terminal symbols
    '
    '
    If (COLparseTree Is Nothing) Then
        listTerminals = "No parse tree exists: cannot list terminals"
        Exit Function
    End If
    listTerminals = listSymbols(COLparseTree.item(2), _
                                False, _
                                COLparseTree, _
                                booHeader)
End Function
Private Function maxLineLength(ByVal strInstring As String) As String
    '
    ' Determine maximum line length
    '
    '
    Dim intIndex1 As Integer
    Dim intMaxLength As Integer
    Dim strSplit() As String
    On Error GoTo maxLineLength_Lbl1_splitErrorHandler
        strSplit = Split(strInstring, vbNewLine)
    On Error GoTo 0
    For intIndex1 = LBound(strSplit) To UBound(strSplit)
        intMaxLength = OBJutilities.max(intMaxLength, Len(Trim(strSplit(intIndex1))))
    Next intIndex1
    maxLineLength = intMaxLength
    Exit Function
maxLineLength_Lbl1_splitErrorHandler:
    OBJutilities.errorHandler 0, _
                              "Cannot split string: " & Err.Number & " " & Err.Description, _
                              "maxLineLength", _
                              Name
End Function
Private Function mkCollectionSortIndex(ByRef colCollection As Collection, _
                                       ByRef colIndex As Collection, _
                                       ByRef colExcluded As Collection) As Boolean
    '
    ' Rather inefficient, but makes a list of collection indexes in ascending sequence
    '
    '
    Dim booSkipExcluded As Boolean
    Dim colHandle(1 To 2) As Collection
    Dim lngIndex1 As Long
    Dim lngIndexMin As Long
    On Error GoTo mkCollectionSortIndex_Lbl1_errorHandler
        Set colIndex = New Collection
    On Error GoTo 0
    With colCollection
        Do
            lngIndexMin = 0
            For lngIndex1 = 1 To .Count
                Set colHandle(1) = .item(lngIndex1)
                booSkipExcluded = False
                If Not (colExcluded Is Nothing) Then
                    booSkipExcluded = inCollection(colExcluded, colHandle(1).item(1))
                End If
                If Not booSkipExcluded Then
                    If Not inCollection(colIndex, symbol2Key(lngIndex1)) Then
                        If lngIndexMin = 0 Then
                            lngIndexMin = lngIndex1
                        Else
                            Set colHandle(2) = .item(lngIndexMin)
                            If colHandle(1).item(1) < colHandle(2).item(1) Then
                                lngIndexMin = lngIndex1
                            End If
                        End If
                    End If
                End If
            Next lngIndex1
            If lngIndexMin = 0 Then Exit Do
            On Error GoTo mkCollectionSortIndex_Lbl1_errorHandler
                colIndex.Add lngIndexMin, symbol2Key(lngIndexMin)
            On Error GoTo 0
        Loop
    End With
    mkCollectionSortIndex = True
    Exit Function
mkCollectionSortIndex_Lbl1_errorHandler:
    collectionErrorHandler Err.Number, Err.Description, "mkCollectionSortIndex"
End Function
Private Function mkParseTree() As Boolean
    '
    ' Create our parse tree
    '
    '
    mkParseTree = False
    If Not (COLparseTree Is Nothing) Then
        If MsgBox("Note that a parse tree exists already.  " & _
                  "Do you want to delete this object?", _
                  vbYesNo) _
           <> _
           vbYes Then
            Exit Function
        End If
        If Not parseTree_destroy(COLparseTree) Then Exit Function
    End If
    mkParseTree = parseTree_create(COLparseTree)
End Function
Private Function mkWideString(ByVal strInstring As String) As String
    '
    ' Make string into a   w i d e   string
    '
    '
    Dim lngIndex1 As Long
    Dim strOutstring As String
    mkWideString = "": strOutstring = ""
    For lngIndex1 = 1 To Len(strInstring)
        OBJutilities.append strOutstring, _
                            " ", _
                            Mid(strInstring, lngIndex1, 1), _
                            booAppendInPlace:=True
    Next lngIndex1
    mkWideString = strOutstring
End Function
Private Sub mkReferenceManual()
    '
    ' Make reference manual
    '
    '
    Dim lngLength As Long
    Dim strComment As String
    Dim strFileid As String
    Dim strManual As String
    Dim strQuoted As String
    With frmRefManOptions
        .LanguageName = "anonymousLanguage"
        If filBNFlocation <> "" Then
            .LanguageName = Mid$(filBNFlocation, _
                                   1, _
                                   OBJutilities.max(0, _
                                                    InStrRev(filBNFlocation, ".") - 1))
            .LanguageName = Mid$(.LanguageName, _
                                 1, _
                                 InStr(.LanguageName & " ", " ") - 1)
        End If
        If .ShowForm Then
            .Show vbModal
            If .Cancel Then Exit Sub
        End If
        If Not .XML Then
            strComment = OBJutilities.string2Box _
                                  (OBJutilities.string2HardParagraph _
                                   (UCase$(mkWideString("Reference manual for the " & _
                                                        .LanguageName & " " & _
                                                        "language")), _
                                    60, _
                                    strSpace:="   ", _
                                    enuAlignment:=ENUalign.alignCenter)) & _
                                  vbNewLine & vbNewLine & vbNewLine
        End If
        strManual = strComment & _
                    languageReference(COLparseTree, _
                                      USRscanned, _
                                      .BoxComments, _
                                      .IncludeBNF, _
                                      .IncludeNonterminals, _
                                      .IncludeSyntax, _
                                      .IncludeTerminals, _
                                      .IncludeWhere, _
                                      .LanguageName, _
                                      .IndentReferenceManual, _
                                      .XML, _
                                      .XMLnewline, _
                                      .XMLinlineBNF, _
                                      .XMLendtagComments)
        lngLength = Len(strManual)
        If lngLength > 32767 Then
            strFileid = OBJutilities.attachPath(App.Path, "bnf.TXT")
            strQuoted = OBJutilities.enquote(strFileid)
            strFileid = InputBox$("The reference manual may not fit into the display: " & _
                                  "because it is " & _
                                  lngLength & " " & _
                                  "characters, it may be truncated at the end. " & _
                                  vbNewLine & vbNewLine & _
                                  "Click OK to save the reference manual in the file " & _
                                  strQuoted & ", " & _
                                  "change the file id to a preferred value and click OK " & _
                                  "or just click Cancel to proceed." & _
                                  vbNewLine & vbNewLine & _
                                  "If you click OK, you can then view the file " & _
                                  "in most contemporary versions of Notepad.", , _
                                  strFileid)
            If strFileid <> "" Then
                On Error GoTo mkReferenceManual_LBL1_string2FileErrorHandler
                    OBJutilities.string2File strManual, strFileid
                    MsgBox "The complete BNF is now available in the file " & strQuoted
                On Error GoTo 0
            End If
        End If
        OBJutilities.errorHandler 0, _
                                  strManual, _
                                  "mkReferenceManual", _
                                  Me.Name
    End With
    Exit Sub
mkReferenceManual_LBL1_string2FileErrorHandler:
    OBJutilities.errorHandler 0, _
                              "The complete BNF was not written to a file: " & _
                              Err.Number & " " & Err.Description, _
                              "mkReferenceManual", _
                              Name
End Sub
Private Function parseTree2Rules(ByRef colTree As Collection, _
                                 ByRef usrScan() As TYPscanned, _
                                 ByVal booIncludeBNF As Boolean, _
                                 ByVal booIncludeWhere As Boolean, _
                                 ByVal booIndentReferenceManual As Boolean) As String
    '
    ' Create the language rule outline from the parse tree
    '
    '
    ' Unfortunately this code rather parallels parseTree2XML as well as
    ' parseTree2RRDiagram in the railroad diagram form. This is a misfortune since
    ' if the structure of the parse tree changes all of these routines must change.
    '
    '
    Dim bytSequence As Byte
    Dim colHandle(1 To 2) As Collection
    Dim lngIndex1 As Long
    Dim lngIndex2 As Long
    Dim strPrefix As String
    Dim strReport As String
    Dim strRules As String
    parseTree2Rules = ""
    If (colTree Is Nothing) Then Exit Function
    updateStatus "Creating the reference manual for this language", 1
    With colTree
        Set colHandle(1) = .item(1)
        With colHandle(1)
            lngIndex2 = 1
            For lngIndex1 = 1 To .Count
                updateProgress "Creating the reference manual for this language", _
                               "nonterminal", _
                               lngIndex1, _
                               .Count
                Set colHandle(2) = .item(lngIndex1)
                With colHandle(2)
                    If nonTerminal2NodeIndex(colTree, .item(1)) <> 0 Then
                        OBJutilities.append _
                            strRules, _
                            vbNewLine, _
                            parseTree2Rules_node2Rules(colTree, _
                                                       usrScan, _
                                                       .item(1), _
                                                       nonTerminal2NodeIndex _
                                                       (colTree, .item(1)), _
                                                       booIncludeBNF, _
                                                       booIncludeWhere, _
                                                       lngIndex2, _
                                                       parseTree_isExtensionName _
                                                       (colHandle(2).item(1), _
                                                        strPrefix, _
                                                        bytSequence), _
                                                       booIndentReferenceManual), _
                            booAppendInPlace:=True
                        lngIndex2 = lngIndex2 + 1
                    End If
                End With
            Next lngIndex1
            updateStatus "", -1
            updateStatus "Reference manual complete"
            lblStatusProgress.Visible = False
        End With
    End With
    parseTree2Rules = removeBlankLines(strRules)
End Function
Private Function parseTree2Rules_node2Rules(ByRef colTree As Collection, _
                                             ByRef usrScan() As TYPscanned, _
                                             ByVal strNonterminal As String, _
                                             ByVal lngNodeIndex As Long, _
                                             ByVal booIncludeBNF As Boolean, _
                                             ByVal booIncludeWhere As Boolean, _
                                             ByVal strOutlineNumber As String, _
                                             ByVal booExtension As Boolean, _
                                             ByVal booIndentReferenceManual) _
        As String
    '
    ' Expand one node to its rules
    '
    '
    Dim colNodeHandle As Collection
    Dim colSubnodeHandle As Collection
    Dim intLast As Integer
    Dim intOutlineLevel As Integer
    Dim lngIndex1 As Long
    Dim strConsistsOf As String
    Dim strName As String
    Dim strNounPhrase As String
    Dim strRule As String
    Dim strSplit() As String
    Dim strSubnodeType As String
    Dim strTerminal As String
    Dim strWhere As String
    With colTree
        If lngNodeIndex = 0 Then
            parseTree2Rules_node2Rules = "": Exit Function
        End If
        Set colNodeHandle = .item(lngNodeIndex)
        With colNodeHandle
            If (TypeOf .item(3) Is Collection) Then
                Set colSubnodeHandle = .item(3)
                If .item(1) <> "" Then
                    strConsistsOf = symbol2Text(.item(1))
                    strNounPhrase = strConsistsOf
                    If booIncludeBNF Then
                        strConsistsOf = strConsistsOf & " " & _
                                    "(" & _
                                    "with the BNF " & _
                                    OBJutilities.enquote(scannedPhrase(usrScan, _
                                                                       .item(4), _
                                                                       .item(5))) & _
                                    ")"
                    End If
                    strSubnodeType = "can " & _
                                     IIf(booExtension, "also ", "") & _
                                     "consist of the following "
                Else
                    strConsistsOf = "": strSubnodeType = "This "
                End If
                Select Case colSubnodeHandle.item(1)
                    Case OPERATOR_ALTERNATION: _
                         strSubnodeType = strSubnodeType & "set of alternatives"
                    Case OPERATOR_MRE_ONETRIP: _
                         strSubnodeType = strSubnodeType & _
                                          "(repeated one or more times)"
                    Case OPERATOR_OPTIONALSEQUENCE: _
                         strSubnodeType = strSubnodeType & _
                                          "optional sequence"
                    Case OPERATOR_PRODUCTION: _
                    Case OPERATOR_MRE_ZEROTRIP: _
                         strSubnodeType = strSubnodeType & _
                                          "(repeated zero, one or more times)"
                    Case OPERATOR_SEQUENCE: strSubnodeType = strSubnodeType & "sequence"
                    Case Else:
                        OBJutilities.errorHandler 0, _
                                                  "Internal programming error: " & _
                                                  "unsupported operator", _
                                                  "parseTree2Rules_node2Rules", _
                                                  Me.Name
                End Select
                If booIncludeWhere _
                   And _
                   Not booExtension _
                   And _
                   colNodeHandle.item(1) <> "" Then
                    strWhere = parseTree_whereList(colTree, colNodeHandle.item(1), True)
                    If strWhere = "" Then
                        strWhere = strNounPhrase & " is a start symbol"
                    Else
                        strWhere = strNounPhrase & " can appear in " & _
                                   strWhere
                    End If
                    strWhere = vbNewLine & OBJutilities.string2HardParagraph(strWhere)
                End If
                strConsistsOf = strConsistsOf & " " & _
                                Trim$(strSubnodeType) & ": " & _
                                vbNewLine & _
                                parseTree2Rules_subnode2String(colTree, _
                                                               usrScan, _
                                                               colSubnodeHandle, _
                                                               booIncludeBNF, _
                                                               booIncludeWhere, _
                                                               strOutlineNumber, _
                                                               booIndentReferenceManual) & _
                                IIf(booIncludeWhere, _
                                    strWhere, _
                                    "")
            Else
                lngIndex1 = .item(3)
                If lngIndex1 > 0 Then
                    strConsistsOf = _
                        symbol2Text(index2Name(colTree, lngIndex1, True))
                ElseIf lngIndex1 < 0 Then
                    lngIndex1 = Abs(lngIndex1)
                    strTerminal = index2Name(colTree, lngIndex1, False)
                    If OBJutilities.isQuoted(strTerminal) Then
                        strConsistsOf = "The string " & strTerminal
                    Else
                        strConsistsOf = string2NounPhrase(strTerminal)
                    End If
                Else
                    strConsistsOf = "undefined syntax"
                End If
            End If
            On Error GoTo parseTree2Rules_node2Rules_Lbl1_splitErrorHandler
                strSplit = Split(strOutlineNumber, ".")
            On Error GoTo 0
            If booIndentReferenceManual And UBound(strSplit) < INDENT_MAX Then
                intLast = UBound(strSplit) - 2
                For lngIndex1 = 0 To intLast
                    intOutlineLevel = intOutlineLevel _
                                      + _
                                      Len(strSplit(lngIndex1))
                    If lngIndex1 < intLast Then intOutlineLevel = intOutlineLevel + 1
                Next lngIndex1
                strRule = OBJutilities.copies(" ", intOutlineLevel)
            End If
            strRule = strRule & _
                      strOutlineNumber & _
                      IIf(intOutlineLevel = 0, ".  ", " ") & _
                      removeBlankLines(strConsistsOf)
            If booIndentReferenceManual And UBound(strSplit) < 5 Then
                strRule = indentContinuations(strRule)
            End If
        End With
    End With
    parseTree2Rules_node2Rules = strRule
    Exit Function
parseTree2Rules_node2Rules_Lbl1_splitErrorHandler:
    OBJutilities.errorHandler 0, _
                              "Error in split: " & _
                              Err.Number & " " & Err.Description, _
                              "parseTree2Rules_node2Rules", _
                              Me.Name
End Function
Private Function parseTree2Rules_subnode2String(ByRef colTree As Collection, _
                                                ByRef usrScan() As TYPscanned, _
                                                ByRef colSubnode As Collection, _
                                                ByVal booIncludeBNF As Boolean, _
                                                ByVal booIncludeWhere As Boolean, _
                                                ByVal strOutlineNumber As String, _
                                                ByVal booIndentReferenceManual As Boolean) _
        As String
    '
    ' Subnode (operator, operand1, operand2) to string
    '
    '
    Dim lngIndex1 As Long
    Dim strOutstring As String
    On Error Resume Next
        lngIndex1 = colSubnode.item(2)
    On Error GoTo 0
    If lngIndex1 = 0 Then Exit Function
    strOutstring = parseTree2Rules_node2Rules(colTree, _
                                               usrScan, _
                                               "", _
                                               lngIndex1, _
                                               booIncludeBNF, _
                                               booIncludeWhere, _
                                               strOutlineNumber & ".1", _
                                               False, _
                                               booIndentReferenceManual)
    lngIndex1 = 0
    On Error Resume Next
        lngIndex1 = colSubnode.item(3)
    On Error GoTo 0
    If lngIndex1 <> 0 Then
        strOutstring = strOutstring & _
                       vbNewLine & _
                       parseTree2Rules_node2Rules(colTree, _
                                                   usrScan, _
                                                   "", _
                                                   lngIndex1, _
                                                   booIncludeBNF, _
                                                   booIncludeWhere, _
                                                   strOutlineNumber & ".2", _
                                                   False, _
                                                   booIndentReferenceManual)
    End If
    parseTree2Rules_subnode2String = strOutstring
End Function
Private Sub nonTerminals2Screen()
    '
    ' Place the nonterminals on the display
    '
    '
    Dim colHandle As Collection
    Set colHandle = COLparseTree.item(1)
    symbols2Screen colHandle, lstNonterminals
End Sub
Private Sub ownErrorHandler(ByVal strMessage As String)
    '
    ' Handle errors that can't be handled by the utilities error handler
    '
    '
    MsgBox strMessage
End Sub
Private Function parseBNF() As Boolean
    '
    ' Parse the Backus-Naur Form
    '
    '
    parseBNF = False
    If filBNFlocation.ListIndex < 0 Then
        MsgBox "No input file has been selected"
        Exit Function
    End If
    If parseTreeExists Then
        If Not parseTree_destroy(COLparseTree) Then Exit Function
    End If
    If Not parseTree_create(COLparseTree) Then Exit Function
    parseBNF = BNFcompile(getBNF)
End Function
Private Sub parseInterface()
    '
    ' Parse the BNF: screen interface
    '
    '
    If parseBNF Then
        MsgBox "BNF parse has succeeded"
    Else
        MsgBox "BNF parse has failed"
    End If
End Sub
Private Function parseTreeExists() As Boolean
    '
    ' Tell caller if the parse tree exists
    '
    '
    parseTreeExists = Not (COLparseTree Is Nothing)
End Function
Private Sub referenceManualInterface()
    '
    ' Parse language syntax and create reference manual
    '
    '
    If (COLparseTree Is Nothing) Then
        If Not parseBNF Then
            MsgBox "Can't produce reference manual: there are errors in the " & _
                   "language definition"
            Exit Sub
        End If
    End If
    mkReferenceManual
End Sub
Private Sub registry2Form()
    '
    ' Retrieve form settings
    '
    '
    Dim lngIndex1 As Long
    Dim strFileTitle As String
    updateStatus "Retrieving form settings"
    On Error GoTo registry2Form_Lbl1_errorHandler
        chkProgress = GetSetting(App.EXEName, _
                                 Me.Name, _
                                 chkProgress.Name, _
                                 chkProgress)
        optParseStatusNoReport = GetSetting(App.EXEName, _
                                            Me.Name, _
                                            optParseStatusNoReport.Name, _
                                            optParseStatusNoReport)
        optParseStatusSimpleReport = GetSetting(App.EXEName, _
                                                Me.Name, _
                                                optParseStatusSimpleReport.Name, _
                                                optParseStatusSimpleReport)
        optParseStatusCompleteReport = GetSetting(App.EXEName, _
                                                  Me.Name, _
                                                  optParseStatusCompleteReport.Name, _
                                                  optParseStatusCompleteReport)
        drvBNFlocation = GetSetting(App.EXEName, _
                                    Me.Name, _
                                    drvBNFlocation.Name, _
                                    drvBNFlocation)
        dirBNFlocation = GetSetting(App.EXEName, _
                                    Me.Name, _
                                    dirBNFlocation.Name, _
                                    dirBNFlocation)
        With filBNFlocation
            strFileTitle = UCase$(GetSetting(App.EXEName, _
                                             Me.Name, _
                                             .Name, _
                                             ""))
            If strFileTitle <> "" Then
                For lngIndex1 = 0 To .ListCount - 1
                    If UCase$(.List(lngIndex1)) = strFileTitle Then
                        Exit For
                    End If
                Next lngIndex1
                If lngIndex1 < .ListCount Then .ListIndex = lngIndex1
            End If
        End With
    On Error GoTo 0
    Exit Sub
registry2Form_Lbl1_errorHandler:
    OBJutilities.errorHandler 0, _
                              "Could not save form settings: " & _
                              Err.Number & " " & Err.Description, _
                              "registry2Form", _
                              Me.Name
End Sub
Private Function removeBlankLines(ByVal strInstring As String) As String
    '
    ' Remove blank lines from string
    '
    '
    Dim lngIndex1 As Long
    Dim strOutstring As String
    Dim strSplit() As String
    On Error GoTo removeBlankLines_Lbl1_splitErrorHandler
        strSplit = Split(strInstring, vbNewLine)
    On Error GoTo 0
    For lngIndex1 = LBound(strSplit) To UBound(strSplit)
        If Trim$(strSplit(lngIndex1)) <> "" Then
            OBJutilities.append strOutstring, _
                                vbNewLine, _
                                strSplit(lngIndex1), _
                                booAppendInPlace:=True
        End If
    Next lngIndex1
    removeBlankLines = strOutstring
    Exit Function
removeBlankLines_Lbl1_splitErrorHandler:
    OBJutilities.errorHandler 0, _
                              "Error in split: " & Err.Number & Err.Description, _
                              "removeBlankLines", _
                              Me.Name
End Function
Private Function scannedPhrase(ByRef usrScan() As TYPscanned, _
                               ByVal lngStartIndex As Long, _
                               ByVal lngLength As Long)
    '
    ' Return the string-normalized phrase of tokens
    '
    '
    Dim lngEndIndex As Long
    Dim lngIndex1 As Long
    Dim strPhrase As String
    scannedPhrase = ""
    If lngStartIndex > UBound(usrScan) Then Exit Function
    lngEndIndex = lngStartIndex + lngLength - 1
    For lngIndex1 = lngStartIndex To lngEndIndex
        strPhrase = strPhrase & usrScan(lngIndex1).strTokenValue
        If lngIndex1 < lngEndIndex Then strPhrase = strPhrase & " "
    Next lngIndex1
    scannedPhrase = strPhrase
End Function
Private Function string2NounPhrase(ByVal strInstring As String) As String
    '
    ' Convert string to noun phrase, that begins with "a" or with "an"
    '
    '
    If InStr("AEIOU", UCase$(Mid(strInstring, 1, 1))) <> 0 Then
        string2NounPhrase = "An " & strInstring: Exit Function
    End If
    string2NounPhrase = "A " & strInstring
End Function
Private Function string2ScanTokenType(ByVal strInstring As String) _
        As ENUscannedType
    '
    ' Return the type of the scan token
    '
    '
    Dim usrScan() As TYPscanned
    string2ScanTokenType = ENUscannedType.unknown
    If Not scannedAllocate(usrScan) Then Exit Function
    If Not BNFcompile_scanner(strInstring, usrScan) Then Exit Function
    If UBound(usrScan) <> 1 Then Exit Function
    string2ScanTokenType = usrScan(1).enuType
End Function
Private Function symbol2Deco(ByVal strSymbol As String) As String
    '
    ' Translate symbol to terminal(symbol), nonTerminal(symbol),
    ' or var2Deco's format
    '
    '
    If isTerminal(strSymbol) Then
        symbol2Deco = "terminal(" & strSymbol & ")"
    ElseIf isNonTerminal(strSymbol) Then
        symbol2Deco = "nonTerminal(" & strSymbol & ")"
    Else
        symbol2Deco = OBJutilities.var2Deco(strSymbol)
    End If
End Function
Private Function symbol2Key(ByVal strSymbol As String) As String
    '
    ' Convert the symbol to a case-sensitive collection key (avoids number keys)
    '
    '
    ' The key is prefixed by an asterisk. Each lower case letter in the key is
    ' prefixed by an underscore. For example, Key is translated to *K_e_y.
    '
    '
    Dim intAsc As Integer
    Dim intAsc_a As Integer
    Dim intAsc_z As Integer
    Dim intIndex1 As Integer
    Dim strNext As String
    Dim strOutstring As String
    intAsc_a = AscW("a"): intAsc_z = AscW("z")
    For intIndex1 = 1 To Len(strSymbol)
        strNext = Mid(strSymbol, intIndex1, 1)
        intAsc = AscW(strNext)
        If intAsc >= intAsc_a And intAsc <= intAsc_z Then strOutstring = strOutstring & "_"
        strOutstring = strOutstring & strNext
    Next intIndex1
    symbol2Key = "*" & strOutstring
End Function
Private Function symbol2Text(ByVal strSymbol As String) As String
    '
    ' Converts symbol to camel case noun phrase
    '
    '
    symbol2Text = string2NounPhrase(camelCase2ProperCase(strSymbol))
End Function
Private Function scannedAllocate(ByRef usrScan() As TYPscanned) As Boolean
    '
    ' Allocate the scantable
    '
    '
    scannedAllocate = False
    On Error GoTo scannedAllocate_Lbl1_errorHandler
        ReDim usrScan(0 To 0)
    On Error GoTo 0
    scannedAllocate = True
    Exit Function
scannedAllocate_Lbl1_errorHandler:
    OBJutilities.errorHandler 0, _
                              "Cannot allocate the scan table: " & _
                              Err.Number & " " & Err.Description, _
                              "", _
                              Me.Name
End Function
Private Function scannedExists() As Boolean
    '
    ' Tell caller if the scantable is allocated
    '
    '
    Dim lngUBound As Long
    scannedExists = False
    lngUBound = -1
    On Error Resume Next
        lngUBound = UBound(USRscanned)
    On Error GoTo 0
    If lngUBound > -1 Then scannedExists = True
End Function
Private Function selectiveDeco(ByVal varValue As Variant) As String
    '
    ' Selective var2Deco
    '
    '
    ' If varValue converts without error to a string that contains newlines, blanks, letters,
    ' numbers and selected special characters exclusively, varValue's string value is
    ' returned in quotes. Newlines are changed to <newline>.
    '
    ' Otherwise, the utility method var2Deco is used to convert varValue to a string.
    '
    '
    Dim intErr As Integer
    Dim strValue As String
    selectiveDeco = ""
    On Error Resume Next
        strValue = varValue: intErr = Err.Number
    On Error GoTo 0
    If intErr = 0 Then
        strValue = Replace(strValue, vbNewLine, "<newline>")
        If OBJutilities.verify(strValue, _
                               " " & _
                               "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
                               "abcdefghijklmnopqrstuvwxyz" & _
                               "0123456789" & _
                               "~`!@#$%^&*()_+-={}[]:;'<>,./|\") _
           = _
           0 Then
            selectiveDeco = OBJutilities.enquote(strValue)
            Exit Function
        End If
    End If
    strValue = OBJutilities.var2Deco(varValue)
    If UCase$(strValue) = "VBERROR(ERROR 448)" Then
        strValue = "unspecified value"
    End If
    selectiveDeco = strValue
End Function
Private Sub symbols2Screen(ByRef colCollection As Collection, ByRef lstBox As ListBox)
    '
    ' Place the symbols on the display
    '
    '
    Dim colHandle As Collection
    Dim lngIndex1 As Long
    lstBox.Clear
    With colCollection
        For lngIndex1 = 1 To .Count
            Set colHandle = .item(lngIndex1)
            updateListbox lstBox, colHandle(1)
        Next lngIndex1
    End With
    With lstBox
        .ListIndex = IIf(.ListCount = 0, -1, 0): .Refresh
    End With
End Sub
Private Sub terminals2Screen()
    '
    ' Place the terminals on the display
    '
    '
    Dim colHandle As Collection
    Set colHandle = COLparseTree.item(2)
    symbols2Screen colHandle, lstTerminals
End Sub
Private Sub toggleMenu(ByVal strFileTitle As String)
    '
    ' Toggle File and its associated separator
    '
    '
    Static strCaption As String
    If strCaption = "" Then strCaption = mnuFileParse.Caption
    With mnuFileParse
        If strFileTitle <> "" Then
            .Caption = Replace(strCaption, _
                               "%FILETITLE", _
                               strFileTitle)
            .Visible = True: mnuFileSep1.Visible = True
        Else
            .Visible = False: mnuFileSep1.Visible = False
        End If
    End With
End Sub
Private Sub updateListbox(ByRef lstBox As ListBox, ByVal strItem As String)
    '
    ' Add one item to list box
    '
    '
    With lstBox
        .AddItem strItem
        .ListIndex = .ListCount - 1
        On Error Resume Next
            .SetFocus
        On Error GoTo 0
        .Refresh
    End With
End Sub
Private Sub updateProgress(ByVal strActivity As String, _
                           ByVal strEntity As String, _
                           ByVal lngEntityNumber As Long, _
                           ByVal lngEntityCount As Long)
    '
    ' Update a progress/status report
    '
    '
    If chkProgress <> 1 Then Exit Sub
    With lblStatusProgress
        .Width = OBJutilities.histogram(lngEntityNumber, _
                                        dblRangeMax:=lstStatus.Width, _
                                        dblValueMax:=lngEntityCount)
        .Visible = True
        .Refresh
    End With
    updateStatus strActivity & " " & _
                 "at " & _
                 strEntity & " " & _
                 lngEntityNumber & " " & _
                 "of " & _
                 lngEntityCount, _
                 0
End Sub
Private Sub updateStatus(ByVal strMessage As String, _
                         Optional ByVal intLvlChange As Integer = 0)
    '
    ' Updates the status report
    '
    '
    Static intLvl As Integer
    If strMessage <> "" Then
        updateListbox lstStatus, String$(intLvl * 3, " ") & Now & " " & strMessage
    End If
    intLvl = OBJutilities.min(intLvl + intLvlChange, 5)
    If intLvl < 0 Then intLvl = 0
End Sub
Private Function BNFcompile(ByVal strBNF As String) As Boolean
    '
    ' Scan and parse the input BNF
    '
    '
    BNFcompile = False
    lstNonterminals.Clear: lstTerminals.Clear
    If Not BNFcompile_scanner(strBNF, USRscanned) Then Exit Function
    If Not BNFcompile_parser(USRscanned, COLparseTree) Then Exit Function
    nonTerminals2Screen
    terminals2Screen
    BNFcompile = True
End Function
Private Function BNFcompile_parser(ByRef usrScanTable() As TYPscanned, _
                                   ByRef colParse As Collection) As Boolean
    Dim lngIndex1 As Long
    BNFcompile_parser = False
    updateStatus "Parsing the scanned BNF", 1
    lngIndex1 = LBound(usrScanTable) + 1
    BNFcompile_parser = BNFcompile_parser_bnfGrammar(usrScanTable, _
                                                     colParse, _
                                                     lngIndex1)
    lblStatusProgress.Visible = False
    updateStatus "", -1
    updateStatus "Parse complete", 0
    If lngIndex1 <= UBound(usrScanTable) Then
        updateStatus "Unrecognizable material at end of BNF grammar " & _
                     "commences at character " & _
                     usrScanTable(lngIndex1).lngStartIndex & ": " & _
                     selectiveDeco _
                     (OBJutilities.ellipsis _
                      (scannedPhrase(usrScanTable, _
                                     lngIndex1, _
                                     UBound(usrScanTable) - lngIndex1 + 1), _
                       32)), _
                     0
        BNFcompile_parser = False
        Exit Function
    End If
    If parseTree_isEmpty(colParse) Then
        updateStatus "No productions in BNF file", 0
        BNFcompile_parser = False
        Exit Function
    End If
    BNFcompile_parser = parseTree_inspection(colParse)
End Function
Private Function BNFcompile_parser_alternationFactorRHS(ByRef usrScanTable() As TYPscanned, _
                                                        ByRef colParse As Collection, _
                                                        ByRef lngIndex As Long, _
                                                        ByRef varNode As Variant, _
                                                        ByVal varNodeLHS As Variant, _
                                                        ByVal lngEndIndex As Long, _
                                                        ByVal strLHS As String) As Boolean
    '
    ' alternationFactorRHS := "|" mockRegularExpression [ alternationFactorRHS ]
    '
    '
    Dim lngIndex1 As Long
    Dim lngIndex2 As Long
    Dim varNodeRHS As Variant
    BNFcompile_parser_alternationFactorRHS = False
    BNFcompile_parser_updateStatus_start "alternationFactor", _
                                         "alternationFactorRHS := ""|"" mockRegularExpression [ alternationFactorRHS ]", _
                                         lngIndex, _
                                         usrScanTable, _
                                         strNodeInfo:="Left hand side node: " & _
                                                      OBJutilities.var2Deco(varNodeLHS)
    lngIndex1 = lngIndex: lngIndex2 = lngIndex
    If Not BNFcompile_parser_checkToken(usrScanTable, _
                                        lngIndex, _
                                        ENUscannedType.alternation, _
                                        lngEndIndex, _
                                        booErrorReport:=True) Then
        BNFcompile_parser_updateStatus_end "alternationFactorRHS", _
                                           False, _
                                           lngIndex, lngIndex, _
                                           usrScanTable, _
                                           strComment:="Expected only alternation op (|) " & _
                                                       "at this location"
        Exit Function
    End If
    If Not BNFcompile_parser_mockRegularExpression(usrScanTable, _
                                                   colParse, _
                                                   lngIndex, _
                                                   varNodeRHS, _
                                                   lngEndIndex, _
                                                   strLHS) Then
        lngIndex1 = lngIndex
        BNFcompile_parser_updateStatus_end "alternationFactorRHS", _
                                           False, _
                                           lngIndex, lngIndex, _
                                           usrScanTable, _
                                           strComment:="Alternation operator is " & _
                                                       "followed by invalid expression"
        Exit Function
    End If
    varNode = parseTree_addNode(colParse, _
                                "", _
                                lngIndex1, lngIndex - lngIndex1, _
                                OPERATOR_ALTERNATION, _
                                varNodeLHS, _
                                varNodeRHS)
    BNFcompile_parser_alternationFactorRHS usrScanTable, _
                                           colParse, _
                                           lngIndex, _
                                           varNode, _
                                           varNode, _
                                           lngEndIndex, _
                                           strLHS
    BNFcompile_parser_updateStatus_end "alternationFactorRHS", _
                                       True, _
                                       lngIndex1, lngIndex, _
                                       usrScanTable, _
                                       strNodeInfo:="Alternation node: " & _
                                                    OBJutilities.var2Deco(varNode)
    BNFcompile_parser_alternationFactorRHS = True
End Function
Private Function BNFcompile_parser_bnfGrammar(ByRef usrScanTable() As TYPscanned, _
                                              ByRef colParse As Collection, _
                                              ByRef lngIndex As Long) As Boolean
    '
    ' bnfGrammar := production [ bnfGrammar ]
    '
    '
    Dim lngIndex1 As Long
    BNFcompile_parser_bnfGrammar = False
    lngIndex1 = lngIndex
    BNFcompile_parser_updateStatus_start "bnfGrammar", _
                                         "bnfGrammar := production [ , production ]", _
                                         lngIndex, _
                                         usrScanTable
    If Not BNFcompile_parser_production(usrScanTable, colParse, lngIndex) Then
        BNFcompile_parser_updateStatus_end "bnfGrammar", _
                                           False, _
                                           lngIndex1, lngIndex, _
                                           usrScanTable
        Exit Function
    End If
    If lngIndex < UBound(usrScanTable) Then
        BNFcompile_parser_bnfGrammar usrScanTable, colParse, lngIndex
    End If
    BNFcompile_parser_bnfGrammar = True
    BNFcompile_parser_updateStatus_end "bnfGrammar", _
                                       True, _
                                       lngIndex1, lngIndex, _
                                       usrScanTable
End Function
Private Function BNFcompile_parser_checkToken(ByRef usrScanTable() As TYPscanned, _
                                              ByRef lngIndex As Long, _
                                              ByVal enuWhatFor As ENUscannedType, _
                                              ByVal lngEndIndex As Long, _
                                              Optional ByVal varWhatFor As Variant, _
                                              Optional ByVal booErrorReport As Boolean = False) As Boolean
    '
    ' Check token, advancing index on success
    '
    '
    ' Checks the next available token for the enuWhatFor type and (unless varWhatFor
    ' IsMissing) the string value in varWhatFor.
    '
    '
    Dim strComment As String
    Dim strWhatFor As String
    BNFcompile_parser_checkToken = False
    If lngIndex > lngEndIndex Then Exit Function
    updateStatus "", 1
    updateProgress "Compiling the BNF: " & _
                   "checking for the scantype " & _
                   scanTokenType2String(enuWhatFor) & " " & _
                   "and " & _
                   selectiveDeco(varWhatFor), _
                   "token", _
                   lngIndex, _
                   UBound(usrScanTable)
    With usrScanTable(lngIndex)
        If .enuType = enuWhatFor Then
            If IsMissing(varWhatFor) Then
                If getReportOption <> noReport Then
                    updateStatus "Found token specified by type!", 0
                End If
                BNFcompile_parser_checkToken = True
                lngIndex = lngIndex + 1
                updateStatus "", -1
                Exit Function
            Else
                On Error GoTo BNFcompile_parser_checkToken_Lbl1_errorHandler
                    strWhatFor = varWhatFor
                On Error GoTo 0
                If .strTokenValue = strWhatFor Then
                    If getReportOption <> noReport Then
                        updateStatus "Found token specified by type and value!", 0
                    End If
                    BNFcompile_parser_checkToken = True
                    lngIndex = lngIndex + 1
                    updateStatus "", -1
                    Exit Function
                End If
            End If
        End If
        If booErrorReport And getReportOption <> noReport Then
            If enuWhatFor = nonterminalIdentifier _
               And _
               IsMissing(varWhatFor) _
               And _
               isIdentifier(usrScanTable(lngIndex).strTokenValue) Then
                strComment = ".  " & _
                             "Expected a nonterminal identifier: found " & _
                             OBJutilities.enquote(usrScanTable(lngIndex).strTokenValue) & ". " & _
                             "Note that non-terminal identifiers MUST be camel case, " & _
                             "and start with lower case letters only."
            End If
            updateStatus "Expected " & _
                         scanTokenType2String(enuWhatFor) & " " & _
                         "at position " & _
                         .lngStartIndex & " " & _
                         IIf(Not IsMissing(varWhatFor), _
                             "with the value " & _
                             selectiveDeco(.strTokenValue), _
                             "") & ": " & _
                         "found " & _
                         selectiveDeco(.strTokenValue) & _
                         strComment, _
                         0
        End If
    End With
    updateStatus "", -1
    Exit Function
BNFcompile_parser_checkToken_Lbl1_errorHandler:
    OBJutilities.errorHandler 0, _
                              "Cannot convert varWhatFor to strWhatFor: " & _
                              Err.Number & " " & Err.Description, _
                              "", _
                              Me.Name
End Function
Private Function BNFcompile_parser_findOrAddNonterminal(ByRef colTree As Collection, _
                                                        ByVal strNonterminal As String) _
                 As Integer
    '
    ' Find or add nonterminal, return its index in colTree(1)
    '
    '
    BNFcompile_parser_findOrAddNonterminal = BNFcompile_parser_findOrAddSymbol(colTree, 1, _
                                                                               strNonterminal)
End Function
Private Function BNFcompile_parser_findOrAddSymbol(ByRef colTree As Collection, _
                                                   ByVal bytIndex As Byte, _
                                                   ByVal strSymbol As String) _
                 As Integer
    '
    ' Find or add terminal or nonterminal, return its index
    '
    '
    Dim colHandle As Collection
    Dim lngIndex1 As Long
    With colTree.item(bytIndex)
        On Error Resume Next
            Set colHandle = .item(symbol2Key(strSymbol))
            If Not (colHandle Is Nothing) Then lngIndex1 = colHandle.item(2)
        On Error GoTo 0
        If lngIndex1 = 0 Then lngIndex1 = parseTree_addSymbol(colTree, bytIndex, strSymbol)
        BNFcompile_parser_findOrAddSymbol = lngIndex1
    End With
End Function
Private Function BNFcompile_parser_findOrAddTerminal(ByRef colTree As Collection, _
                                                     ByVal strTerminal As String) _
                 As Integer
    '
    ' Find or add nonterminal, return its index in colTree(1)
    '
    '
    BNFcompile_parser_findOrAddTerminal = BNFcompile_parser_findOrAddSymbol(colTree, 2, _
                                                                            strTerminal)
End Function
Private Function BNFcompile_parser_findRightSide(ByRef usrScanTable() As TYPscanned, _
                                                 ByVal lngIndex As Long, _
                                                 ByVal lngEndIndex As Long, _
                                                 ByVal strLeftSide As String, _
                                                 ByVal strRightSide As String) As Long
    '
    ' Finds the balancing right parenthesis
    '
    '
    Dim intLevel As Integer
    Dim lngIndex1 As Long
    BNFcompile_parser_findRightSide = lngEndIndex
    intLevel = 1
    For lngIndex1 = lngIndex To lngEndIndex
        Select Case usrScanTable(lngIndex1).strTokenValue
            Case strLeftSide: intLevel = intLevel + 1
            Case strRightSide:
                intLevel = intLevel - 1
                If intLevel = 0 Then Exit For
        End Select
    Next lngIndex1
    If lngIndex1 > lngEndIndex Then
        updateStatus "No balancing right parenthesis was found for left parenthesis at " & _
                     lngIndex
    End If
    BNFcompile_parser_findRightSide = lngIndex1
End Function
Private Function BNFcompile_parser_mockRegularExpression(ByRef usrScanTable() As TYPscanned, _
                                                         ByRef colParse As Collection, _
                                                         ByRef lngIndex As Long, _
                                                         ByRef varNode As Variant, _
                                                         ByVal lngEndIndex As Long, _
                                                         ByVal strLHS As String) As Boolean
    '
    ' mockRegularExpression := mreFactor [ mrePostfix ]
    '
    '
    Dim lngIndex1 As Long
    Dim varNodeFactor As Variant
    BNFcompile_parser_mockRegularExpression = False
    BNFcompile_parser_updateStatus_start "BNFcompile_parser_mockRegularExpression", _
                                         "mockRegularExpression := mreFactor [ mrePostfix ]", _
                                         lngIndex, _
                                         usrScanTable
    lngIndex1 = lngIndex
    If Not BNFcompile_parser_mreFactor(usrScanTable, _
                                       colParse, _
                                       lngIndex, _
                                       varNodeFactor, _
                                       lngEndIndex, _
                                       strLHS) Then
        BNFcompile_parser_updateStatus_end "mockRegularExpression", _
                                           False, _
                                           lngIndex, lngIndex, _
                                           usrScanTable, _
                                           strComment:="Expected nonterminal, terminal, " & _
                                                       "or parenthesized expression: " & _
                                                       "found " & _
                                                       selectiveDeco(usrScanTable(lngIndex1).strTokenValue)
        Exit Function
    End If
    varNode = varNodeFactor
    If BNFcompile_parser_checkToken(usrScanTable, _
                                    lngIndex, _
                                    ENUscannedType.mreOperator, _
                                    lngEndIndex) Then
        With usrScanTable(lngIndex - 1)
            varNode = parseTree_addNode(colParse, _
                                        "", _
                                        lngIndex1, lngIndex - lngIndex1, _
                                        .strTokenValue, _
                                        varOperand1:=varNodeFactor)
        End With
    End If
    BNFcompile_parser_updateStatus_end "mockRegularExpression", _
                                       True, _
                                       lngIndex1, lngIndex, _
                                       usrScanTable, _
                                       strNodeInfo:="Node info = " & _
                                                    OBJutilities.var2Deco(varNode)
    BNFcompile_parser_mockRegularExpression = True
End Function
Private Function BNFcompile_parser_mreFactor(ByRef usrScanTable() As TYPscanned, _
                                             ByRef colParse As Collection, _
                                             ByRef lngIndex As Long, _
                                             ByRef varNode As Variant, _
                                             ByVal lngEndIndex As Long, _
                                             ByVal strLHS As String) As Boolean
    '
    ' mreFactor := nonTerminal | _
    '              terminal | STRING |
    '              "(" productionRHS ")" |
    '              "[" productionRHS "]"
    '
    '
    Dim booOK As Boolean
    Dim lngIndex1 As Long
    Dim strNonterminal As String
    Dim strLeftSide As String
    Dim strRightSide As String
    BNFcompile_parser_mreFactor = False
    BNFcompile_parser_updateStatus_start "BNFcompile_parser_mreFactor", _
                                         "mreFactor := nonTerminal | STRING | NONCANONICALSPECIAL | ( productionRHS )", _
                                         lngIndex, _
                                         usrScanTable
    lngIndex1 = lngIndex
    If BNFcompile_parser_nonTerminal(usrScanTable, _
                                     colParse, _
                                     lngIndex, _
                                     strNonterminal, _
                                     lngEndIndex) Then
        varNode = BNFcompile_parser_findOrAddNonterminal(colParse, strNonterminal)
        If varNode <> 0 Then
            parseTree_addWhereUsed colParse, True, strLHS, varNode
            varNode = parseTree_addNode(colParse, _
                                        "", _
                                        lngIndex1, _
                                        lngIndex - lngIndex1, _
                                        varNode)
        End If
    Else
        booOK = BNFcompile_parser_checkToken(usrScanTable, _
                                             lngIndex, _
                                             ENUscannedType.terminalIdentifier, _
                                             lngEndIndex)
        If Not booOK Then booOK = BNFcompile_parser_checkToken(usrScanTable, _
                                                               lngIndex, _
                                                               ENUscannedType.stringToken, _
                                                               lngEndIndex)
        If booOK Then
            varNode = -BNFcompile_parser_findOrAddTerminal(colParse, _
                                                           usrScanTable(lngIndex - 1).strTokenValue)
            If varNode <> 0 Then
                parseTree_addWhereUsed colParse, False, strLHS, -varNode
                varNode = parseTree_addNode(colParse, _
                                            "", _
                                            lngIndex1, _
                                            lngIndex - lngIndex1, _
                                            varNode)
            End If
        Else
            If BNFcompile_parser_checkToken(usrScanTable, _
                                            lngIndex, _
                                            ENUscannedType.parenthesis, _
                                            lngEndIndex, _
                                            varWhatFor:="(") Then
                strLeftSide = "("
            End If
            If strLeftSide = "(" Then
                strRightSide = ")"
            Else
                If BNFcompile_parser_checkToken(usrScanTable, _
                                                lngIndex, _
                                                ENUscannedType.parenthesis, _
                                                lngEndIndex, _
                                                varWhatFor:="[") Then
                    strLeftSide = OPERATOR_OPTIONALSEQUENCE
                    strRightSide = "]"
                End If
            End If
            If strLeftSide <> "" Then
                lngIndex1 = lngIndex
                If Not BNFcompile_parser_productionRHS(usrScanTable, _
                                                       colParse, _
                                                       lngIndex, _
                                                       varNode, _
                                                       BNFcompile_parser_findRightSide _
                                                       (usrScanTable, _
                                                        lngIndex, _
                                                        lngEndIndex, _
                                                        strLeftSide, _
                                                        strRightSide), _
                                                       strLHS) Then
                    BNFcompile_parser_updateStatus_end "mreFactor", _
                                                       False, _
                                                       lngIndex1, lngIndex, _
                                                       usrScanTable, _
                                                       strComment:="Parenthesized production RHS is not valid"
                    Exit Function
                End If
                If strLeftSide = OPERATOR_OPTIONALSEQUENCE Then
                    varNode = parseTree_addNode(colParse, _
                                                "", _
                                                lngIndex1, _
                                                lngIndex - lngIndex1, _
                                                strLeftSide, _
                                                varOperand1:=varNode)
                End If
                lngIndex = lngIndex + 1
            Else
                BNFcompile_parser_updateStatus_end "mreFactor", _
                                                   False, _
                                                   lngIndex, lngIndex, _
                                                   usrScanTable, _
                                                   strComment:="mreFactor starts with " & _
                                                               "invalid characters"
                Exit Function
            End If
        End If
    End If
    BNFcompile_parser_updateStatus_end "mreFactor", _
                                       True, _
                                       lngIndex1, lngIndex, _
                                       usrScanTable, _
                                       strNodeInfo:="Node info = " & _
                                                    OBJutilities.var2Deco(varNode)
    BNFcompile_parser_mreFactor = True
End Function
Private Function BNFcompile_parser_nonTerminal(ByRef usrScanTable() As TYPscanned, _
                                               ByRef colParse As Collection, _
                                               ByRef lngIndex As Long, _
                                               ByRef strNonterminal As String, _
                                               ByVal lngEndIndex As Long) As Boolean
    '
    ' nonTerminal := IDENTIFIER
    '
    '
    BNFcompile_parser_nonTerminal = False
    BNFcompile_parser_updateStatus_start "nonTerminal", _
                                         "nonTerminal := IDENTIFIER", _
                                         lngIndex, _
                                         usrScanTable
    If Not BNFcompile_parser_checkToken(usrScanTable, _
                                        lngIndex, _
                                        ENUscannedType.nonterminalIdentifier, _
                                        lngEndIndex) Then
        BNFcompile_parser_updateStatus_end "nonTerminal", _
                                           False, _
                                           lngIndex, lngIndex, _
                                           usrScanTable
        Exit Function
    End If
    strNonterminal = usrScanTable(lngIndex - 1).strTokenValue
    BNFcompile_parser_findOrAddNonterminal colParse, strNonterminal
    BNFcompile_parser_nonTerminal = True
    BNFcompile_parser_updateStatus_end "nonTerminal", _
                                       True, _
                                       lngIndex - 1, lngIndex, _
                                       usrScanTable
End Function
Private Function BNFcompile_parser_production(ByRef usrScanTable() As TYPscanned, _
                                              ByRef colParse As Collection, _
                                              ByRef lngIndex As Long) As Boolean
    '
    ' production := nonTerminal := productionRHS ( NEWLINE | EOF )
    ' production := NEWLINE ' Allows for stripped comments
    '
    '
    Dim booEOF As Boolean
    Dim booNewline As Boolean
    Dim booNull As Boolean
    Dim lngIndex1 As Long
    Dim strNonterminal As String
    Dim varNode As Variant
    BNFcompile_parser_production = False
    lngIndex1 = lngIndex
    BNFcompile_parser_updateStatus_start "production", _
                                         "production := [ nonTerminal := productionRHS ] NEWLINE", _
                                         lngIndex, _
                                         usrScanTable
    If BNFcompile_parser_nonTerminal(usrScanTable, _
                                     colParse, _
                                     lngIndex, _
                                     strNonterminal, _
                                     UBound(usrScanTable)) Then
        If Not BNFcompile_parser_checkToken(usrScanTable, _
                                            lngIndex, _
                                            ENUscannedType.productionAssignment, _
                                            UBound(usrScanTable), _
                                            booErrorReport:=True) Then
            lngIndex = lngIndex1
            BNFcompile_parser_updateStatus_end "production", _
                                               False, _
                                               lngIndex1, lngIndex, _
                                               usrScanTable, _
                                               strComment:="Production assignment operator is " & _
                                                           "missing"
            Exit Function
        End If
        If Not BNFcompile_parser_productionRHS(usrScanTable, _
                                               colParse, _
                                               lngIndex, _
                                               varNode, _
                                               UBound(usrScanTable), _
                                               strNonterminal) Then
            lngIndex = lngIndex1
            BNFcompile_parser_updateStatus_end "production", _
                                               False, _
                                               lngIndex1, lngIndex, _
                                               usrScanTable, _
                                               strComment:="Cannot parse right hand side of production"
            Exit Function
        End If
        parseTree_addNode colParse, _
                          strNonterminal, _
                          lngIndex1, lngIndex - lngIndex1, _
                          ":=", _
                          varOperand1:=varNode
    Else
        booNull = True
    End If
    booEOF = (lngIndex > UBound(usrScanTable))
    booNewline = BNFcompile_parser_checkToken(usrScanTable, _
                                              lngIndex, _
                                              ENUscannedType.newline, _
                                              UBound(usrScanTable), _
                                              booErrorReport:=True)
    If booNull And Not booNewline Or Not booEOF And Not booNewline Then
        BNFcompile_parser_updateStatus_end "production", _
                                           False, _
                                           lngIndex1, lngIndex, _
                                           usrScanTable, _
                                           strComment:= _
                                           "Nonnull production does not end with " & _
                                           "newline or eof, " & _
                                           "or null production does not end with newline"
        lngIndex = lngIndex1
        Exit Function
    End If
    BNFcompile_parser_production = True
    BNFcompile_parser_updateStatus_end "production", _
                                       True, _
                                       lngIndex1, lngIndex, _
                                       usrScanTable
End Function
Private Function BNFcompile_parser_productionRHS(ByRef usrScanTable() As TYPscanned, _
                                                 ByRef colParse As Collection, _
                                                 ByRef lngIndex As Long, _
                                                 ByRef varNode As Variant, _
                                                 ByVal lngEndIndex As Long, _
                                                 ByVal strLHS As String) As Boolean
    '
    ' productionRHS := sequenceFactor [ productionRHS ]
    '
    '
    Dim lngIndex1 As Long
    Dim varNode2 As Variant
    BNFcompile_parser_productionRHS = False
    lngIndex1 = lngIndex
    BNFcompile_parser_updateStatus_start "productionRHS", _
                                         "productionRHS := alternationFactor [ alternationFactorRHS ]", _
                                         lngIndex, _
                                         usrScanTable
    If Not BNFcompile_parser_sequenceFactor(usrScanTable, _
                                            colParse, _
                                            lngIndex, _
                                            varNode, _
                                            lngEndIndex, _
                                            strLHS) Then
        BNFcompile_parser_updateStatus_end "productionRHS", _
                                           False, _
                                           lngIndex, lngIndex, _
                                           usrScanTable, _
                                           strComment:="Right hand side does not start " & _
                                                       "with a sequence factor"
        BNFcompile_parser_productionRHS = False
        Exit Function
    End If
    If lngIndex <= UBound(usrScanTable) Then
        If BNFcompile_parser_productionRHS(usrScanTable, _
                                            colParse, _
                                            lngIndex, _
                                            varNode2, _
                                            lngEndIndex, _
                                            strLHS) Then
            varNode = parseTree_addNode(colParse, _
                                        "", _
                                        lngIndex1, lngIndex - lngIndex1, _
                                        OPERATOR_SEQUENCE, _
                                        varOperand1:=varNode, _
                                        varOperand2:=varNode2)
        End If
    End If
    BNFcompile_parser_updateStatus_end "productionRHS", _
                                       True, _
                                       lngIndex1, lngIndex, _
                                       usrScanTable
    BNFcompile_parser_productionRHS = True
    Exit Function
BNFcompile_parser_productionRHS_Lbl1_errorHandler:
    lngIndex = lngIndex1
    BNFcompile_parser_updateStatus_end "productionRHS", _
                                       False, _
                                       lngIndex, lngIndex, _
                                       usrScanTable, _
                                       strComment:="Collection error: " & _
                                                   Err.Number & " " & Err.Description
End Function
Private Function BNFcompile_parser_sequenceFactor(ByRef usrScanTable() As TYPscanned, _
                                                  ByRef colParse As Collection, _
                                                  ByRef lngIndex As Long, _
                                                  ByRef varNode As Variant, _
                                                  ByVal lngEndIndex As Long, _
                                                  ByVal strLHS As String) As Boolean
    '
    ' sequenceFactor := mockRegularExpression [ alternationFactorRHS ]
    '
    '
    Dim lngIndex1 As Long
    BNFcompile_parser_sequenceFactor = False
    BNFcompile_parser_updateStatus_start "sequenceFactor", _
                                         "sequenceFactor := alternationFactor [ alternationFactorRHS ]", _
                                         lngIndex, _
                                         usrScanTable
    lngIndex1 = lngIndex
    If Not BNFcompile_parser_mockRegularExpression(usrScanTable, _
                                                    colParse, _
                                                    lngIndex, _
                                                    varNode, _
                                                    lngEndIndex, _
                                                    strLHS) Then
        BNFcompile_parser_updateStatus_end "sequenceFactor", _
                                           False, _
                                           lngIndex, lngIndex, _
                                           usrScanTable, _
                                           strComment:="Sequence factor does not start " & _
                                                       "with alternation factor"
        Exit Function
    End If
    BNFcompile_parser_alternationFactorRHS usrScanTable, _
                                           colParse, _
                                           lngIndex, _
                                           varNode, _
                                           varNode, _
                                           lngEndIndex, _
                                           strLHS
    BNFcompile_parser_updateStatus_end "sequenceFactor", _
                                       True, _
                                       lngIndex1, lngIndex, _
                                       usrScanTable, _
                                       strNodeInfo:="Node info = " & _
                                                    OBJutilities.var2Deco(varNode)
    BNFcompile_parser_sequenceFactor = True
End Function
Private Sub BNFcompile_parser_updateStatus_end(ByVal strGoalNonterminal As String, _
                                               ByVal booSuccess As Boolean, _
                                               ByVal lngStartIndex As Long, _
                                               ByVal lngEndIndex As Long, _
                                               ByRef usrScanTable() As TYPscanned, _
                                               Optional ByVal strComment As String = "", _
                                               Optional ByVal strNodeInfo As String = "")
    '
    ' End the current level of status reporting
    '
    '
    Dim strNonterminal As String
    If getReportOption = noReport Then Exit Sub
    With OBJutilities
        If booSuccess Then
            strNonterminal = selectiveDeco(.ellipsis(scannedPhrase(usrScanTable, _
                                                                   lngStartIndex, _
                                                                   lngEndIndex _
                                                                   - _
                                                                   lngStartIndex), _
                                                      64))
        End If
        updateStatus "The check for the nonterminal " & _
                     OBJutilities.enquote(strGoalNonterminal) & " " & _
                     "has " & _
                     IIf(booSuccess, _
                         "succeeded: " & _
                         "the nonterminal's value is " & _
                         strNonterminal, _
                         "failed") & _
                     IIf(strNodeInfo <> "", ": " & strNodeInfo, "") & _
                     IIf(strComment <> "", ": " & strComment, ""), _
                     -1
    End With
End Sub
Private Sub BNFcompile_parser_updateStatus_start(ByVal strGoalNonterminal As String, _
                                                 ByVal strUsingProduction As String, _
                                                 ByVal lngIndex As Long, _
                                                 ByRef usrScanTable() As TYPscanned, _
                                                 Optional ByVal strNodeInfo As String = "")
    '
    ' Start a new level of parse status reporting
    '
    '
    Dim enuReportOption As ENUreportOptions
    Dim strHandle As String
    Dim strReport As String
    enuReportOption = getReportOption
    If enuReportOption = noReport Then Exit Sub
    With OBJutilities
        If enuReportOption = completeReport Then
            If lngIndex <= UBound(usrScanTable) Then
                strHandle = ": the handle is " & _
                            selectiveDeco(usrScanTable(lngIndex).strTokenValue)
            End If
        End If
        updateStatus "", 1
        Select Case enuReportOption
            Case completeReport:
                strReport = "Checking for the nonterminal " & _
                               .enquote(strGoalNonterminal) & " " & _
                               "using the production " & _
                               .enquote(strUsingProduction) & _
                               strHandle & _
                               IIf(strNodeInfo = "", "", ": " & strNodeInfo)
            Case simpleReport:
                strReport = "Checking for " & strGoalNonterminal
            Case Else:
                OBJutilities.errorHandler 0, _
                                          "Programming error: " & _
                                          "unexpected enumerator", _
                                          "BNFcompile_parser_updateStatus_start", _
                                          Me.Name
        End Select
        updateProgress strReport, _
                       "token", _
                       lngIndex, _
                       usrScanTable(UBound(usrScanTable)).lngStartIndex _
                       + _
                       usrScanTable(UBound(usrScanTable)).lngLength
    End With
End Sub
Private Function BNFcompile_scanner(ByVal strBNF As String, _
                                    ByRef usrScan() As TYPscanned) As Boolean
    '
    ' Scanner
    '
    '
    Dim bytIndex1 As Byte
    Dim bytIndex2 As Byte
    Dim lngFirst As Long
    Dim lngIndex1 As Long
    Dim lngIndex2 As Long
    Dim lngPreceding As Long
    Dim strGap As String
    Dim usrNext(1 To 10) As TYPscanned
    BNFcompile_scanner = False
    If Not scannedAllocate(usrScan) Then Exit Function
    lngIndex1 = 1: lngPreceding = 1
    updateStatus "Scanning BNF", 1
    Do While lngIndex1 <= Len(strBNF)
        updateProgress "Scanning BNF", _
                       "character", _
                       lngIndex1, _
                       Len(strBNF)
        BNFcompile_scanner_findAlternation strBNF, usrNext(1), lngIndex1
        BNFcompile_scanner_findComment strBNF, usrNext(2), lngIndex1
        BNFcompile_scanner_findComma strBNF, usrNext(3), lngIndex1
        BNFcompile_scanner_findMREoperator strBNF, usrNext(4), lngIndex1
        BNFcompile_scanner_findNewline strBNF, usrNext(5), lngIndex1
        BNFcompile_scanner_findNonterminalIdentifier strBNF, usrNext(6), lngIndex1
        BNFcompile_scanner_findParenthesis strBNF, usrNext(7), lngIndex1
        BNFcompile_scanner_findProductionAssignment strBNF, usrNext(8), lngIndex1
        BNFcompile_scanner_findStringToken strBNF, usrNext(9), lngIndex1
        BNFcompile_scanner_findTerminalIdentifier strBNF, usrNext(10), lngIndex1
        lngFirst = Len(strBNF) + 1: bytIndex2 = 0
        For bytIndex1 = LBound(usrNext) To UBound(usrNext)
            With usrNext(bytIndex1)
                If lngFirst > .lngStartIndex Then
                    lngFirst = .lngStartIndex:
                    bytIndex2 = bytIndex1
                End If
            End With
        Next bytIndex1
        If bytIndex2 = 0 Then Exit Do
        strGap = Mid$(strBNF, _
                      lngPreceding, _
                      usrNext(bytIndex2).lngStartIndex _
                      - _
                      lngPreceding)
        If OBJutilities.verify(strGap, OBJutilities.range2String(CInt(0), CInt(32))) _
           <> _
           0 Then
            OBJutilities.errorHandler 0, _
                                      "Invalid characters " & _
                                      selectiveDeco(strGap) & " " & _
                                      "appear at " & _
                                      lngPreceding & ".." & _
                                      usrNext(bytIndex2).lngStartIndex - 1 & " " & _
                                      "in BNF", _
                                      "BNFcompile_scanner", _
                                      Me.Name
            updateStatus "Scan failed", -1: lblStatusProgress.Visible = False
            Exit Function
        End If
        With usrNext(bytIndex2)
            lngPreceding = .lngStartIndex + .lngLength
            If .enuType <> comment Then
                On Error GoTo BNFcompile_scanner_redimErrorHandler
                    ReDim Preserve usrScan(LBound(usrScan) To UBound(usrScan) + 1)
                On Error GoTo 0
                With usrScan(UBound(usrScan))
                    .enuType = usrNext(bytIndex2).enuType
                    .lngLength = usrNext(bytIndex2).lngLength
                    .lngStartIndex = usrNext(bytIndex2).lngStartIndex
                    .strTokenValue = usrNext(bytIndex2).strTokenValue
                    If .enuType = productionAssignment _
                       And _
                       .strTokenValue = "=" Then
                        updateStatus "Use of the equals sign by itself at position " & _
                                     .lngStartIndex & " " & _
                                     "is accepted as the production operator: " & _
                                     "note that := is the standard operator"
                    End If
                End With
            End If
        End With
        lngIndex1 = lngPreceding
    Loop
    lblStatusProgress.Visible = False
    updateStatus "", -1
    updateStatus "Scan complete", 0
    BNFcompile_scanner = True
    If UBound(usrScan) <= 100 Then BNFcompile_scanner = scannerInspection
    Exit Function
BNFcompile_scanner_redimErrorHandler:
    OBJutilities.errorHandler 0, _
                              "Cannot expand scan array: " & _
                              Err.Number & " " & Err.Description, _
                              "BNFcompile_scanner", _
                              Me.Name
End Function
Private Sub BNFcompile_scanner_findAlternation(ByVal strBNF As String, _
                                               ByRef usrNext As TYPscanned, _
                                               ByVal lngIndex As Long)
    '
    ' Find the next stroke op
    '
    '
    With usrNext
        .lngStartIndex = InStr(lngIndex, _
                               strBNF & OPERATOR_ALTERNATION, _
                               OPERATOR_ALTERNATION)
        .lngLength = 1
        .enuType = alternation
        .strTokenValue = Mid$(strBNF, .lngStartIndex, .lngLength)
    End With
End Sub
Private Sub BNFcompile_scanner_findComma(ByVal strBNF As String, _
                                         ByRef usrNext As TYPscanned, _
                                         ByVal lngIndex As Long)
    '
    ' Find the next comma
    '
    '
    With usrNext
        .lngStartIndex = InStr(lngIndex, strBNF & ",", ",")
        .lngLength = 1
        .enuType = comma
        .strTokenValue = Mid$(strBNF, .lngStartIndex, .lngLength)
    End With
End Sub
Private Sub BNFcompile_scanner_findComment(ByVal strBNF As String, _
                                           ByRef usrNext As TYPscanned, _
                                           ByVal lngIndex As Long)
    '
    ' Find the next comment
    '
    '
    With usrNext
        .lngStartIndex = InStr(lngIndex, strBNF & "'", "'")
        .lngLength = InStr(.lngStartIndex, strBNF & vbNewLine, vbNewLine) _
                     - _
                     .lngStartIndex
        .enuType = comment
        .strTokenValue = Mid$(strBNF, .lngStartIndex, .lngLength)
    End With
End Sub
Private Sub BNFcompile_scanner_findMREoperator(ByVal strBNF As String, _
                                               ByRef usrNext As TYPscanned, _
                                               ByVal lngIndex As Long)
    '
    ' Find the next mock regular expression op (asterisk or plus)
    '
    '
    With usrNext
        .lngStartIndex = OBJutilities.verify(strBNF & "*", _
                                             OPERATOR_MRE_ZEROTRIP & _
                                             OPERATOR_MRE_ONETRIP, _
                                             intStartIndex:=lngIndex, _
                                             booMatch:=True)
        .lngLength = 1
        .enuType = mreOperator
        .strTokenValue = Mid$(strBNF, .lngStartIndex, .lngLength)
    End With
End Sub
Private Sub BNFcompile_scanner_findNewline(ByVal strBNF As String, _
                                           ByRef usrNext As TYPscanned, _
                                           ByVal lngIndex As Long)
    '
    ' Find the next non-continued new line
    '
    '
    ' Note that this routine allows both standards: Windows newlines consisting
    ' of a carriage return followed by a linefeed and Web newlines consisting of
    ' a linefeed. In fact, the two standards may be mixed in one file.
    '
    '
    Dim intNewlineLen As Integer
    With usrNext
        .lngStartIndex = lngIndex
        Do
            .lngStartIndex = OBJutilities.min(InStr(.lngStartIndex, strBNF & vbNewLine, vbNewLine), _
                                              InStr(.lngStartIndex, strBNF & vbLf, vbLf))
            If Mid(strBNF, .lngStartIndex, 1) = vbLf Then
                intNewlineLen = Len(vbLf)
            Else
                intNewlineLen = Len(vbNewLine)
            End If
            If Mid$(strBNF, .lngStartIndex + intNewlineLen, 1) <> " " Then Exit Do
            .lngStartIndex = .lngStartIndex + intNewlineLen
        Loop
        .lngLength = intNewlineLen
        .enuType = newline
        .strTokenValue = Mid$(strBNF, .lngStartIndex, .lngLength)
    End With
End Sub
Private Sub BNFcompile_scanner_findNonterminalIdentifier(ByVal strBNF As String, _
                                                         ByRef usrNext As TYPscanned, _
                                                         ByVal lngIndex As Long)
    '
    ' Find the next nonterminal identifier
    '
    '
    With usrNext
        .lngStartIndex = OBJutilities.verify(strBNF & "a", _
                                             "abcdefghijklmnopqrstuvwxyz", _
                                             intStartIndex:=lngIndex, _
                                             booMatch:=True)
        .lngLength = OBJutilities.verify(strBNF & " ", _
                                         "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
                                         "abcdefghijklmnopqrstuvwxyz" & _
                                         "0123456789_", _
                                         intStartIndex:=.lngStartIndex, _
                                         booMatch:=False)
        .lngLength = .lngLength - .lngStartIndex
        .enuType = nonterminalIdentifier
        .strTokenValue = Mid$(strBNF, .lngStartIndex, .lngLength)
    End With
End Sub
Private Sub BNFcompile_scanner_findParenthesis(ByVal strBNF As String, _
                                             ByRef usrNext As TYPscanned, _
                                             ByVal lngIndex As Long)
    '
    ' Find the next right or left round parenthesis or square bracket
    '
    '
    With usrNext
        .lngStartIndex = OBJutilities.verify(strBNF & "(", _
                                             "()[]", _
                                             intStartIndex:=lngIndex, _
                                             booMatch:=True)
        .lngLength = 1
        .enuType = parenthesis
        .strTokenValue = Mid$(strBNF, .lngStartIndex, .lngLength)
    End With
End Sub
Private Sub BNFcompile_scanner_findProductionAssignment(ByVal strBNF As String, _
                                                        ByRef usrNext As TYPscanned, _
                                                        ByVal lngIndex As Long)
    '
    ' Find the next production op (colon equals)
    '
    '
    With usrNext
        .lngStartIndex = lngIndex
        .lngLength = 1
        Do While .lngStartIndex <= Len(strBNF)
            .lngStartIndex = OBJutilities.verify(strBNF & ":", ":=", _
                                                 intStartIndex:=.lngStartIndex, _
                                                 booMatch:=True)
            If Mid$(strBNF, .lngStartIndex, 1) = "=" Then
                .lngLength = 1
                Exit Do
            ElseIf Mid$(strBNF, .lngStartIndex + 1, 1) = "=" Then
                .lngLength = 2
                Exit Do
            End If
            .lngStartIndex = .lngStartIndex + 1
        Loop
        .enuType = productionAssignment
        .strTokenValue = Mid$(strBNF, .lngStartIndex, .lngLength)
    End With
End Sub
Private Sub BNFcompile_scanner_findStringToken(ByVal strBNF As String, _
                                               ByRef usrNext As TYPscanned, _
                                               ByVal lngIndex As Long)
    '
    ' Find the next string
    '
    '
    With usrNext
        .lngStartIndex = InStr(lngIndex, strBNF & """", """")
        .lngLength = .lngStartIndex + 1
        Do While .lngLength <= Len(strBNF)
            .lngLength = InStr(.lngLength, strBNF & """", """")
            If Mid$(strBNF, .lngLength + 1, 1) <> """" Then Exit Do
            .lngLength = .lngLength + 2
        Loop
        .lngLength = .lngLength - .lngStartIndex + 1
        .enuType = stringToken
        .strTokenValue = Mid$(strBNF, .lngStartIndex, .lngLength)
    End With
End Sub
Private Sub BNFcompile_scanner_findTerminalIdentifier(ByVal strBNF As String, _
                                                      ByRef usrNext As TYPscanned, _
                                                      ByVal lngIndex As Long)
    '
    ' Find the next terminal identifier
    '
    '
    With usrNext
        .lngStartIndex = OBJutilities.verify(strBNF & "A", _
                                             "ABCDEFGHIJKLMNOPQRSTUVWXYZ", _
                                             intStartIndex:=lngIndex, _
                                             booMatch:=True)
        .lngLength = OBJutilities.verify(strBNF & " ", _
                                         "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
                                         "abcdefghijklmnopqrstuvwxyz" & _
                                         "0123456789_", _
                                         intStartIndex:=.lngStartIndex, _
                                         booMatch:=False)
        .lngLength = .lngLength - .lngStartIndex
        .enuType = terminalIdentifier
        .strTokenValue = Mid$(strBNF, .lngStartIndex, .lngLength)
    End With
End Sub
Private Function scanDump(ByRef usrScanTable() As TYPscanned) As String
    '
    ' Dump the scantable
    '
    '
    Dim intIndex1 As Integer
    Dim intMaxLength As Integer
    Dim intMaxStartIndex As Integer
    Dim intMaxToken As Integer
    Dim intMaxType As Integer
    Dim strOutstring As String
    Dim strTokenValue As String
    Dim strType As String
    scanDump = "Could not create the scan dump"
    If Not scannedExists Then
        scanDump = "The scanner table does not exist": Exit Function
    End If
    updateStatus "Creating the scanDump", 1
    ' --- Pass 1: determine maximum columns
    updateStatus "Pass 1: determine maximum columns", 1
    For intIndex1 = LBound(usrScanTable) + 1 To UBound(usrScanTable)
        updateProgress "Pass 1: determine maximum columns", _
                       "token", _
                       intIndex1, _
                       UBound(usrScanTable)
        With usrScanTable(intIndex1)
            strType = scanTokenType2String(.enuType)
            strTokenValue = selectiveDeco(.strTokenValue)
            If intIndex1 = LBound(usrScanTable) + 1 Then
                intMaxLength = Len(CStr(.lngLength))
                intMaxStartIndex = Len(CStr(.lngStartIndex))
                intMaxToken = Len(strTokenValue)
                intMaxType = Len(CStr(strType))
            Else
                intMaxLength = OBJutilities.max(intMaxLength, Len(CStr(.lngLength)))
                intMaxStartIndex = OBJutilities.max(intMaxStartIndex, Len(CStr(.lngStartIndex)))
                intMaxToken = OBJutilities.max(intMaxToken, Len(strTokenValue))
                intMaxType = OBJutilities.max(intMaxType, Len(strType))
            End If
        End With
    Next intIndex1
    lblStatusProgress.Visible = False
    updateStatus "", -1
    updateStatus "Pass 1 complete", 0
    ' --- Pass 2: create the dump
    updateStatus "Pass 2: create the dump", 1
    With OBJutilities
        intMaxLength = .max(6, intMaxLength)
        intMaxStartIndex = .max(5, intMaxStartIndex)
        intMaxToken = .max(11, intMaxToken)
        intMaxType = .max(10, intMaxType)
        strOutstring = "SCAN DUMP AS OF " & Now & vbNewLine & vbNewLine & _
                       .align("Token Type", _
                              intMaxType, _
                              enuAlignment:=alignCenter) & "  " & _
                       .align("Start", _
                              intMaxStartIndex, _
                              enuAlignment:=alignCenter) & "  " & _
                       .align("Length", _
                              intMaxLength, _
                              enuAlignment:=alignCenter) & "   " & _
                       .align("Token Value", _
                              intMaxToken, _
                              enuAlignment:=alignCenter) & vbNewLine
        strOutstring = strOutstring & _
                       String$(intMaxType, "-") & "  " & _
                       String$(intMaxStartIndex, "-") & "  " & _
                       String$(intMaxLength, "-") & "   " & _
                       String$(intMaxToken, "-") & vbNewLine
    End With
    For intIndex1 = LBound(usrScanTable) + 1 To UBound(usrScanTable)
        updateProgress "Pass 2: create the dump", _
                       "token", _
                       intIndex1, _
                       UBound(usrScanTable)
        With OBJutilities
            strOutstring = strOutstring & _
                           .align(scanTokenType2String(usrScanTable(intIndex1).enuType), _
                                  intMaxType) & "  " & _
                           .align(usrScanTable(intIndex1).lngStartIndex, _
                                  intMaxStartIndex, _
                                  enuAlignment:=alignRight) & "  " & _
                           .align(usrScanTable(intIndex1).lngLength, _
                                  intMaxLength, _
                                  enuAlignment:=alignRight) & "   " & _
                           .align(selectiveDeco(usrScanTable(intIndex1).strTokenValue), _
                                  intMaxToken)
        End With
        If intIndex1 < UBound(usrScanTable) Then strOutstring = strOutstring & vbNewLine
    Next intIndex1
    lblStatusProgress.Visible = False
    updateStatus "", -1
    updateStatus "Pass 2 complete", 0
    updateStatus "", -1
    updateStatus "Dump created", 0
    scanDump = strOutstring
End Function
Private Function scanInspect(ByRef strReport As String) As Boolean
    '
    ' Inspects the scan table
    '
    '
    Dim booInspection As Boolean
    Dim lngIndex1 As Long
    Dim strType As String
    scanInspect = False
    updateStatus "Inspecting scan table", 1
    booInspection = True
    strReport = "INSPECTION OF THE SCAN TABLE AS OF " & Now & vbNewLine & vbNewLine
    If inspectAppend(strReport, _
                     SCANNED_INSPECTION_ISSOMETHING, _
                     scannedExists, _
                     booInspection, _
                     booBox:=False) Then
        For lngIndex1 = LBound(USRscanned) + 1 To UBound(USRscanned)
            updateProgress "Inspecting scan table", _
                           "token", _
                           lngIndex1, _
                           UBound(USRscanned)
            With USRscanned(lngIndex1)
                strType = UCase$(scanTokenType2String(.enuType))
                If Not inspectAppend(strReport, _
                                     SCANNED_INSPECTION_TYPEVALID, _
                                     strType <> "Invalid" And strType <> "COMMENT", _
                                     booInspection, _
                                     "Type at " & lngIndex1 & " is " & strType, _
                                     booBox:=False) Then
                    Exit For
                End If
                If Not inspectAppend(strReport, _
                                     SCANNED_INSPECTION_STARTINDEXVALID, _
                                     .lngStartIndex > 0, _
                                     booInspection, _
                                     "Start index at " & lngIndex1 & " is " & .lngStartIndex, _
                                     booBox:=False) Then
                    Exit For
                End If
                If Not inspectAppend(strReport, _
                                     SCANNED_INSPECTION_LENGTHVALID, _
                                     .lngLength > 0, _
                                     booInspection, _
                                     "Length at " & lngIndex1 & " is " & .lngLength, _
                                     booBox:=False) Then
                    Exit For
                End If
                If lngIndex1 > LBound(USRscanned) + 1 Then
                    With USRscanned(lngIndex1 - 1)
                        If Not inspectAppend(strReport, _
                                             SCANNED_INSPECTION_ORDER, _
                                             USRscanned(lngIndex1).lngStartIndex >= _
                                             .lngStartIndex + .lngLength, _
                                             booInspection, _
                                             booBox:=False) Then
                            Exit For
                        End If
                    End With
                End If
            End With
        Next lngIndex1
        lblStatusProgress.Visible = False
    End If
    updateStatus "", -1
    updateStatus "Inspection complete", 0
    scanInspect = booInspection
End Function
Private Function scannerInspection() As Boolean
    '
    ' Internal inspector
    '
    '
    Dim strReport As String
    scannerInspection = False
    If Not scanInspect(strReport) Then
        OBJutilities.errorHandler 0, _
                                  "Inspection of scanned BNF data structure has failed" & _
                                  vbNewLine & vbNewLine & _
                                  strReport, _
                                  "scannerInspection", _
                                  Me.Name
        Exit Function
    End If
    scannerInspection = True
End Function
Private Function scanTokenType2String(ByVal enuType As ENUscannedType) As String
    '
    ' Converts the scantoken's type to its name
    '
    '
    scanTokenType2String = "invalid"
    Select Case enuType
        Case ENUscannedType.alternation: scanTokenType2String = "alternation"
        Case ENUscannedType.comma: scanTokenType2String = "comma"
        Case ENUscannedType.comment: scanTokenType2String = "comment"
        Case ENUscannedType.mreOperator: scanTokenType2String = "mreOperator"
        Case ENUscannedType.newline: scanTokenType2String = "newline"
        Case ENUscannedType.nonterminalIdentifier: scanTokenType2String = "nonterminalIdentifier"
        Case ENUscannedType.parenthesis: scanTokenType2String = "parenthesis"
        Case ENUscannedType.productionAssignment: scanTokenType2String = "productionAssignment"
        Case ENUscannedType.stringToken: scanTokenType2String = "stringToken"
        Case ENUscannedType.terminalIdentifier: scanTokenType2String = "terminalIdentifier"
    End Select
End Function
Private Function parseTree_addNode(ByRef colTree As Collection, _
                                   ByVal strName As String, _
                                   ByVal lngBNFstartIndex As Long, _
                                   ByVal lngBNFlength As Long, _
                                   ByVal varOperation As Variant, _
                                   Optional ByVal varOperand1 As Variant, _
                                   Optional ByVal varOperand2 As Variant) _
        As Integer
    '
    ' Add a named or anonymous node to the specified parse tree
    '
    '
    ' strName should be a null string (for an anonymous node) or the nonterminal
    ' name.
    '
    ' The start and the length of the BNF code in the scan table should be
    ' in lngBNFstartIndex and lngBNFlength.
    '
    ' varOperation must be the node operator: a Long integer, which will be zero,
    ' a positive nonzero nonterminal index, a negative nonzero terminal index,
    ' or a string operator.
    '
    ' When varOperation is a string operator, varOperand1 must be the Long index
    ' to the first operand node in the parse tree; varOperand2 may be the Long
    ' index to the second operand node, in the parse tree.
    '
    ' The index of the new node is returned: 0 indicates an error.
    '
    '
    Dim bytExtension As Byte
    Dim colNew(1 To 2) As Collection
    Dim lngIndex1 As Long
    Dim strExtension As String
    Dim strNameWork As String
    parseTree_addNode = 0
    On Error GoTo parseTree_addNode_Lbl1_errorHandler
        ' --- Make new node
        Set colNew(1) = New Collection
        With colNew(1)
            .Add strName
            .Add colTree.Count + 1
            If Not IsMissing(varOperand1) Then
                Set colNew(2) = New Collection
                With colNew(2)
                    .Add varOperation
                    .Add varOperand1
                    If Not IsMissing(varOperand2) Then .Add varOperand2
                    colNew(1).Add colNew(2)
                End With
            Else
                .Add varOperation
            End If
            .Add lngBNFstartIndex
            .Add lngBNFlength
        End With
        ' --- Add new node
        If strName = "" Then
            ' Add anonymous entry for expression on production's RHS
            colTree.Add colNew(1)
        Else
            ' Add keyed entry for production
            strNameWork = strName
            lngIndex1 = nonTerminal2NodeIndex(colTree, strName)
            If lngIndex1 <> 0 Then
                ' There are more than one productions for one nonterminal
                bytExtension = 1
                Do
                    bytExtension = bytExtension + 1
                    strNameWork = parseTree_mkExtensionName(strName, bytExtension)
                Loop While nonTerminal2NodeIndex(colTree, strNameWork) <> 0
                parseTree_addSymbol colTree, 1, strNameWork
            End If
            colTree.Add colNew(1), symbol2Key(strNameWork)
        End If
    On Error GoTo 0
    parseTree_addNode = colTree.Count
    Exit Function
parseTree_addNode_Lbl1_errorHandler:
    collectionErrorHandler Err.Number, Err.Description, "parseTree_addNode"
End Function
Private Function parseTree_addSymbol(ByRef colTree As Collection, _
                                     ByVal bytIndex As Byte, _
                                     ByVal strSymbol As String) As Long
    '
    ' Add terminal or nonterminal to parse tree
    '
    '
    Dim colHandle(1 To 2) As Collection
    Set colHandle(1) = colTree.item(bytIndex)
    On Error GoTo parseTree_addSymbol_Lbl1_errorHandler
        Set colHandle(2) = New Collection
        With colHandle(2)
            .Add strSymbol: .Add colHandle(1).Count + 1
        End With
        colHandle(1).Add colHandle(2), symbol2Key(strSymbol)
    On Error GoTo 0
    parseTree_addSymbol = colHandle(1).Count
    Exit Function
parseTree_addSymbol_Lbl1_errorHandler:
    collectionErrorHandler Err.Number, Err.Description, "parseTree_addSymbol"
End Function
Private Function parseTree_addWhereUsed(ByRef colTree As Collection, _
                                        ByVal booNonterminal As Boolean, _
                                        ByVal strLHS As String, _
                                        ByVal lngNode As Long) As Boolean
    '
    ' Update the where-used list for the left hand side grammar symbol for the
    ' terminal or nonterminal
    '
    '
    Dim colHandle(1 To 2) As Collection
    Dim lngIndex1 As Long
    parseTree_addWhereUsed = False
    Set colHandle(1) = colTree.item(IIf(booNonterminal, 1, 2))
    With colHandle(1)
        lngIndex1 = parseTree_nonTerminal2ListIndex(colTree, strLHS)
        If lngIndex1 = 0 Then
            OBJutilities.errorHandler 0, _
                                      "Programming error: the left hand side symbol " & _
                                      OBJutilities.enquote(strLHS) & " " & _
                                      "cannot be found", _
                                      "parseTree_addWhereUsed", _
                                      Me.Name
            Exit Function
        End If
        Set colHandle(2) = .item(lngNode)
        With colHandle(2)
            On Error Resume Next
                .Add lngIndex1, symbol2Key(lngIndex1)
            On Error GoTo 0
        End With
        parseTree_addWhereUsed = True
    End With
End Function
Private Function parseTree_copyNode(ByRef colTree As Collection, _
                                    ByRef lngIndex As Long) As Long
    '
    ' Add a copy of the indexed node: return its index
    '
    '
    ' Note: the copy is unkeyed and anonymous.
    '
    '
    Dim bytIndex3 As Byte
    Dim colHandle(1 To 2) As Collection
    Dim lngIndex1 As Long
    Dim lngOperand1 As Long
    Dim lngOperand2 As Long
    Dim varOperation As Variant
    parseTree_copyNode = 0
    Set colHandle(1) = colTree.item(lngIndex)
    With colHandle(1)
        If (TypeOf .item(3) Is Collection) Then
            Set colHandle(2) = .item(3)
            With colHandle(2)
                varOperation = .item(1)
                lngOperand1 = .item(2)
                If .Count > 2 Then lngOperand1 = .item(3)
            End With
        Else
            varOperation = .item(3)
        End If
        If lngOperand2 = 0 Then
            parseTree_copyNode = parseTree_addNode(colTree, _
                                                    "", _
                                                    .item(4), .item(5), _
                                                    varOperation, _
                                                    varOperand1:=lngOperand1)
        Else
            parseTree_copyNode = parseTree_addNode(colTree, _
                                                    "", _
                                                    .item(4), .item(5), _
                                                    varOperation, _
                                                    varOperand1:=lngOperand1, _
                                                    varOperand2:=lngOperand2)
        End If
    End With
    Exit Function
parseTree_copyNode_Lbl1_errorHandler:
    collectionErrorHandler Err.Number, Err.Description, "parseTree_copyNode"
End Function
Private Function parseTree_create(ByRef colTree As Collection) As Boolean
    '
    ' Create an empty parse tree
    '
    '
    Dim booInspection As Boolean
    Dim colSubCollection As Collection
    parseTree_create = False
    updateStatus "Creating the parse tree", 1
    On Error GoTo parseTree_create_lbl1_createErrorHandler
        ' Create the collection tree
        Set colTree = New Collection
        ' Create the nonterminal index
        Set colSubCollection = New Collection
        colTree.Add colSubCollection
        ' Create the terminal index
        Set colSubCollection = New Collection
        colTree.Add colSubCollection
    On Error GoTo 0
    booInspection = parseTree_inspection(colTree)
    updateStatus "", -1
    updateStatus "Parse tree creation has " & _
                 IIf(booInspection, "succeeded", "FAILED")
    parseTree_create = booInspection
    Exit Function
parseTree_create_lbl1_createErrorHandler:
    OBJutilities.errorHandler 0, _
                              "Can't create the parse tree collection: " & _
                              Err.Number & " " & Err.Description, _
                              "parseTree_create", _
                              Me.Name
End Function
Private Function parseTree_destroy(ByRef colTree As Collection) As Boolean
    '
    ' Deallocate the parse tree and all of its objects
    '
    '
    Dim booOK As Boolean
    parseTree_destroy = False
    updateStatus "Deallocating the parse tree", 0
    booOK = OBJutilities.collectionDestroy(colTree)
    updateStatus "Deallocation " & IIf(booOK, "successful", "failed"), 0
    parseTree_destroy = booOK
End Function
Private Function parseTree_dump(ByRef colTree As Collection) As String
    '
    ' Dumps the parse tree as a decorated variant
    '
    '
    parseTree_dump = "The parse tree does not exist"
    If (colTree Is Nothing) Then Exit Function
    parseTree_dump = OBJutilities.var2Deco(colTree)
End Function
Private Function parseTree_extendNode(ByRef colTree As Collection, _
                                      ByVal lngIndex As Long, _
                                      ByVal lngBNFstartIndex As Long, _
                                      ByVal lngBNFlength As Long, _
                                      ByVal varOperation As Variant, _
                                      Optional ByVal varOperand1 As Variant, _
                                      Optional ByVal varOperand2 As Variant) _
        As Boolean
    '
    ' Extend the node with info: operator and operands, BNF source start and length
    '
    '
    Dim colHandle As Collection
    Dim colNew As Collection
    parseTree_extendNode = False
    Set colHandle = colTree.item(lngIndex)
    On Error GoTo parseTree_extendNode_Lbl1_errorHandler
        With colHandle
            If VarType(varOperation) = vbLong Then
                .Add varOperation
            ElseIf Not IsMissing(varOperand1) Then
                Set colNew = New Collection
                With colNew
                    .Add varOperation
                    .Add varOperand1
                    If Not IsMissing(varOperand2) Then
                        .Add varOperand2
                    End If
                End With
                .Add colNew
            Else
                OBJutilities.errorHandler 0, _
                                          "Unexpected parameters", _
                                          "parseTree_extendNode", _
                                          Me.Name
            End If
            .Add lngBNFstartIndex: .Add lngBNFlength
        End With
    On Error GoTo 0
    Exit Function
parseTree_extendNode_Lbl1_errorHandler:
    collectionErrorHandler Err.Number, Err.Description, "parseTree_extendNode"
End Function
Friend Function index2Name(colTree, _
                           ByVal lngIndex As Long, _
                           ByVal booNonterminal As Boolean) As String
    '
    ' Convert symbol's index to its name
    '
    '
    Dim colHandle As Collection
    Set colHandle = colTree.item(IIf(booNonterminal, 1, 2))
    Set colHandle = colHandle.item(lngIndex)
    index2Name = colHandle.item(1)
End Function
Private Function parseTree_inspect(ByRef colTree As Collection, _
                                   ByRef strReport As String) As Boolean
    '
    ' Inspect the parse tree
    '
    '
    Dim booInspection As Boolean
    Dim booOK As Boolean
    Dim strComment As String
    updateStatus "Inspecting the parse tree", 1
    parseTree_inspect = False
    booInspection = True
    strReport = "INSPECTION OF PARSE TREE AT " & Now & _
                vbNewLine & vbNewLine
    If Not inspectAppend(strReport, _
                                    PARSETREE_INSPECTION_ISSOMETHING, _
                                    Not (colTree Is Nothing), _
                                    booInspection) Then
        updateStatus "", -1
        updateStatus "Inspection failed because the parse tree does not exist", 0
        Exit Function
    End If
    With colTree
        If inspectAppend(strReport, _
                                    PARSETREE_INSPECTION_MINCOUNT, _
                                    .Count >= 2, _
                                    booInspection, _
                                    strComment:="Count of collection is " & .Count) Then
            booOK = parseTree_inspect_symbolList(colTree, True, strComment)
            inspectAppend strReport, _
                                     PARSETREE_INSPECTION_NONTERMINAL_LIST, _
                                     booOK, _
                                     booInspection, _
                                     strComment:=strComment
            booOK = parseTree_inspect_symbolList(colTree, False, strComment)
            inspectAppend strReport, _
                                     PARSETREE_INSPECTION_TERMINAL_LIST, _
                                     booOK, _
                                     booInspection, _
                                     strComment:=strComment
            booOK = parseTree_inspect_nodeStructure(colTree, strComment)
            inspectAppend strReport, _
                                     PARSETREE_INSPECTION_NODE_STRUCTURE, _
                                     booOK, _
                                     booInspection, _
                                     strComment:=strComment
        End If
        parseTree_inspect = booInspection
        updateStatus "", -1
        updateStatus "Inspection " & IIf(booInspection, "succeeded", "failed"), 0
    End With
End Function
Private Function parseTree_inspect_nodeStructure(ByRef colTree As Collection, _
                                                 ByRef strComment As String) As Boolean
    '
    ' Inspect the main parse tree, for correct nodes
    '
    '
    Dim booOK As Boolean
    Dim colHandle(1 To 2) As Collection
    Dim lngIndex1 As Long
    Dim varWork As Variant
    Dim varWork2 As Variant
    parseTree_inspect_nodeStructure = False
    strComment = ""
    updateStatus "Inspecting the nodes of the parse tree", 1
    With colTree
        If .Count >= 3 And Not scannedExists Then
            strComment = "Nodes exist but the scan table is nonexistent or empty"
            updateStatus "Node inspection has failed: " & strComment, -1
            lblStatusProgress.Visible = False
            Exit Function
        End If
        For lngIndex1 = 3 To .Count
            updateProgress "Inspecting the nodes of the parse tree", _
                           "node", _
                           lngIndex1 - 2, _
                           .Count - 2
            If Not (TypeOf .item(lngIndex1) Is Collection) Then
                strComment = "Node " & _
                             lngIndex1 & _
                             "is not a Collection"
                updateStatus "Node inspection has failed: " & strComment, -1
                lblStatusProgress.Visible = False
                Exit Function
            End If
            Set colHandle(1) = .item(lngIndex1)
            With colHandle(1)
                If .Count <> 5 Then
                    strComment = "Node " & _
                                 lngIndex1 & _
                                 "has an unexpected Count = " & _
                                 .Count
                    updateStatus "Node inspection has failed: " & strComment, -1
                    lblStatusProgress.Visible = False
                    Exit Function
                End If
                varWork = Null
                On Error Resume Next
                    varWork = CStr(.item(1))
                On Error GoTo 0
                If IsNull(varWork) Then
                    strComment = "Node " & _
                                 lngIndex1 & _
                                 "does not contain a string in item(1): " & _
                                 "it contains " & _
                                 OBJutilities.var2Deco(varWork)
                    updateStatus "Node inspection has failed: " & strComment, -1
                    lblStatusProgress.Visible = False
                    Exit Function
                End If
                varWork = Null
                On Error Resume Next
                    varWork = CLng(.item(2))
                On Error GoTo 0
                booOK = False
                If Not IsNull(varWork) Then
                    booOK = (varWork = lngIndex1)
                End If
                If Not booOK Then
                    strComment = "Node " & _
                                 lngIndex1 & _
                                 "contains an invalid nonterminal index, " & _
                                 "which is not the key of this node"
                    updateStatus "Node inspection has failed: " & strComment, -1
                    lblStatusProgress.Visible = False
                    Exit Function
                End If
                If (TypeOf .item(3) Is Collection) Then
                    Set colHandle(2) = .item(3)
                    With colHandle(2)
                        ' Check the expansion count
                        If .Count < 1 Or .Count > 3 Then
                            strComment = "Node " & _
                                         lngIndex1 & " " & _
                                         "contains an invalid expansion with a Count of " & _
                                         .Count
                            updateStatus "Node inspection has failed: " & strComment, -1
                            lblStatusProgress.Visible = False
                            Exit Function
                        End If
                        ' Check the operand pointers
                        If Not parseTree_inspect_nodeStructure_operand(.item(2), _
                                                                       colTree.Count, _
                                                                       lngIndex1, _
                                                                       strComment) Then
                            Exit Function
                        End If
                        If .Count > 2 Then
                            If Not parseTree_inspect_nodeStructure_operand(.item(3), _
                                                                           colTree.Count, _
                                                                           lngIndex1, _
                                                                           strComment) Then
                                Exit Function
                            End If
                        End If
                    End With
                Else
                    varWork = Null
                    On Error Resume Next
                        varWork = CLng(.item(3))
                    On Error GoTo 0
                    If IsNull(varWork) Then
                        strComment = "Node " & _
                                     lngIndex1 & " " & _
                                     "contains the invalid entry " & _
                                     OBJutilities.var2Deco(varWork) & " " & _
                                     "in item 3"
                        updateStatus "Node inspection has failed: " & strComment, -1
                        lblStatusProgress.Visible = False
                        Exit Function
                    End If
                End If
                varWork = Null
                On Error Resume Next
                    varWork = CLng(.item(4))
                On Error GoTo 0
                booOK = False
                If Not IsNull(varWork) Then
                    booOK = (varWork > LBound(USRscanned) And varWork <= UBound(USRscanned))
                End If
                If Not booOK Then
                    strComment = "Node " & _
                                 lngIndex1 & " " & _
                                 "contains an invalid scantable start index in " & _
                                 "item 4"
                    updateStatus "Node inspection has failed: " & strComment, -1
                    lblStatusProgress.Visible = False
                    Exit Function
                End If
                varWork2 = Null
                On Error Resume Next
                    varWork2 = CLng(.item(5))
                On Error GoTo 0
                booOK = False
                If Not IsNull(varWork2) Then
                    booOK = (varWork2 > 0 And varWork2 <= UBound(USRscanned) - varWork + 1)
                End If
                If Not booOK Then
                    strComment = "Node " & _
                                 lngIndex1 & " " & _
                                 "contains an invalid scantable length in " & _
                                 "item 5"
                    updateStatus "Node inspection has failed: " & strComment, -1
                    lblStatusProgress.Visible = False
                    Exit Function
                End If
            End With
        Next lngIndex1
        lblStatusProgress.Visible = False
        updateStatus "", -1
        updateStatus "Inspection complete", 0
        parseTree_inspect_nodeStructure = True
    End With
End Function
Private Function parseTree_inspect_nodeStructure_operand(ByVal varOperand As Variant, _
                                                         ByVal lngCount As Long, _
                                                         ByVal lngIndex As Long, _
                                                         ByRef strComment As String) _
        As Boolean
    '
    ' Checks the operand field, of a node, that points to an expansion
    '
    '
    Dim booOK As Boolean
    Dim varWork As Variant
    parseTree_inspect_nodeStructure_operand = False
    varWork = Null
    On Error Resume Next
        varWork = Abs(CLng(varOperand))
    On Error GoTo 0
    booOK = False
    If Not IsNull(varWork) Then
        booOK = varWork > 0
    End If
    If Not booOK Then
        strComment = "Node " & _
                     lngIndex & " " & _
                     "contains an invalid index to the parse tree in " & _
                     "item 1"
        updateStatus "Node inspection has failed: " & strComment, -1
        lblStatusProgress.Visible = False
        Exit Function
    End If
    parseTree_inspect_nodeStructure_operand = True
End Function
Private Function parseTree_inspect_symbolList(ByRef colTree As Collection, _
                                              ByVal booNonterminal As Boolean, _
                                              ByRef strComment As String) As Boolean
    '
    ' Inspect the nonterminal/terminal list
    '
    '
    Dim bytIndex1 As Byte
    Dim colHandle(1 To 2) As Collection
    Dim lngIndex1 As Long
    Dim lngIndex2 As Long
    Dim strName As String
    Dim varWork As Variant
    parseTree_inspect_symbolList = False
    strComment = ""
    If booNonterminal Then
        bytIndex1 = 1: strName = "nonterminal"
    Else
        bytIndex1 = 2: strName = "terminal"
    End If
    updateStatus "Inspecting the " & strName & " " & "list of the parse tree", 1
    With colTree
        If Not (TypeOf .item(bytIndex1) Is Collection) Then
            strComment = "The " & strName & " list is not a Collection"
            Exit Function
        End If
        Set colHandle(1) = .item(bytIndex1)
        With colHandle(1)
            For lngIndex1 = 1 To .Count
                updateProgress "Inspecting the " & strName & " " & "list of the parse tree", _
                               strName, _
                               lngIndex1, _
                               .Count
                If Not (TypeOf .item(lngIndex1) Is Collection) Then
                    strComment = strName & " list entry " & _
                                 lngIndex1 & _
                                 "is not a Collection"
                    updateStatus strName & " list inspection has failed: " & strComment, -1
                    lblStatusProgress.Visible = False
                    Exit Function
                End If
                Set colHandle(2) = .item(lngIndex1)
                With colHandle(2)
                    If .Count < 2 Then
                        strComment = strName & " list entry " & _
                                     lngIndex1 & _
                                     "doesn't contain at least two items"
                        updateStatus strName & " list inspection has failed: " & strComment, -1
                        lblStatusProgress.Visible = False
                        Exit Function
                    End If
                    varWork = Null
                    On Error Resume Next
                        varWork = CStr(.item(1))
                    On Error GoTo 0
                    If IsNull(varWork) Then
                        strComment = strName & " list entry " & _
                                     lngIndex1 & _
                                     "doesn't contain a string in item(1)"
                        updateStatus "Nonterminal list inspection has failed: " & strComment, -1
                        lblStatusProgress.Visible = False
                        Exit Function
                    End If
                    If booNonterminal Then
                        lngIndex2 = nonTerminal2NodeIndex(colTree, varWork)
                        If lngIndex2 < 2 Or lngIndex2 > colTree.Count Then
                            strComment = "Nonterminal list entry " & _
                                         lngIndex1 & " " & _
                                         "doesn't contain a valid nonterminal in item(1): " & _
                                         "it contains " & _
                                         OBJutilities.var2Deco(varWork)
                            updateStatus "Nonterminal list inspection has failed: " & strComment, -1
                            lblStatusProgress.Visible = False
                            Exit Function
                        End If
                    End If
                    updateStatus "Checking the index entry for " & _
                                 OBJutilities.enquote(varWork), _
                                 0
                    varWork = 0
                    On Error Resume Next
                        varWork = CLng(.item(2))
                    On Error GoTo 0
                    If varWork <> parseTree_symbol2Index(colTree, booNonterminal, .item(1)) Then
                        strComment = strName & " list entry " & _
                                     lngIndex1 & " " & _
                                     "doesn't contain a valid index in item(2)"
                        updateStatus strName & " list inspection has failed: " & strComment, -1
                        lblStatusProgress.Visible = False
                        Exit Function
                    End If
                    updateStatus "Checking the where-used list for " & strName & _
                                 OBJutilities.enquote(varWork), _
                                 1
                    For lngIndex2 = 3 To .Count
                        varWork = 0
                        On Error Resume Next
                            varWork = CLng(.item(lngIndex2))
                        On Error GoTo 0
                        If varWork < 1 Or varWork > colHandle(1).Count Then
                            strComment = strName & " list entry " & _
                                         lngIndex1 & " " & _
                                         "doesn't contain a valid where-used nonterminal in item" & _
                                         "(" & lngIndex2 & "): " & _
                                         "it contains " & _
                                         OBJutilities.var2Deco(varWork)
                            updateStatus strName & " list inspection has failed: " & strComment, -1
                            lblStatusProgress.Visible = False
                            Exit Function
                        End If
                    Next lngIndex2
                    updateStatus "", -1
                End With
            Next lngIndex1
            updateStatus "", -1
            lblStatusProgress.Visible = False
            parseTree_inspect_symbolList = True
        End With
    End With
End Function
Private Function parseTree_inspection(ByRef colTree As Collection) As Boolean
    '
    ' Internal inspection
    '
    '
    Dim strReport As String
    parseTree_inspection = False
    If Not parseTree_inspect(colTree, strReport) Then
        OBJutilities.errorHandler 0, _
                                  "Parse tree inspection has failed" & _
                                  vbNewLine & vbNewLine & vbNewLine & _
                                  strReport, _
                                  "parseTree_inspection", _
                                  Me.Name
        Exit Function
    End If
    parseTree_inspection = True
End Function
Private Function parseTree_isEmpty(ByRef colTree As Collection) As Boolean
    '
    ' Return True (parse tree is Nothing or empty) or False
    '
    '
    parseTree_isEmpty = True
    If (colTree Is Nothing) Then Exit Function
    If colTree.Count < 3 Then Exit Function
    parseTree_isEmpty = False
End Function
Private Function parseTree_isExtensionName(ByVal strName As String, _
                                           ByRef strPrefix As String, _
                                           ByRef bytSequence As Byte) As Boolean
    '
    ' Returns True when name is in the format name[sequence]: return prefix
    ' and sequence #
    '
    '
    Dim intErr As Integer
    Dim lngIndex1 As Long
    Dim lngIndex2 As Long
    Dim lngLength As Integer
    lngLength = Len(strName)
    If lngLength < 4 Then Exit Function
    lngIndex1 = InStr(strName, "[")
    lngIndex2 = InStr(strName, "]")
    If lngIndex1 < 2 Then Exit Function
    If lngIndex2 < 2 Or lngIndex2 <> lngLength Then Exit Function
    strPrefix = Mid$(strName, 1, lngIndex1 - 1)
    On Error Resume Next
        bytSequence = Mid$(strName, lngIndex1 + 1, lngIndex2 - lngIndex1 - 1)
        intErr = Err.Number
    On Error GoTo 0
    parseTree_isExtensionName = (intErr = 0)
End Function
Private Function parseTree_mkExtensionName(ByVal strName As String, _
                                           ByVal bytSequence As Byte) As String
    '
    ' Create extension name in the format name[sequence]
    '
    '
    parseTree_mkExtensionName = strName & "[" & bytSequence & "]"
End Function
Private Function parseTree_nonTerminal2ListIndex(ByRef colTree As Collection, _
                                                 ByVal strNonterminal As String) As Long
    '
    ' Converts name of the nonterminal to index in the item(1) list, or 0
    '
    '
    Dim colHandle As Collection
    parseTree_nonTerminal2ListIndex = 0
    On Error Resume Next
        Set colHandle = colTree.item(1)
        If Err.Number <> 0 Then Exit Function
        Set colHandle = colHandle.item(symbol2Key(strNonterminal))
        If Err.Number <> 0 Then Exit Function
        parseTree_nonTerminal2ListIndex = colHandle.item(2)
    On Error GoTo 0
End Function
Private Function parseTree_symbol2Index(ByRef colTree As Collection, _
                                        ByVal booNonterminal As Boolean, _
                                        ByVal strSymbol As String) As Long
    '
    ' Locate the terminal or nonterminal in its parse tree list
    '
    '
    If booNonterminal Then
        parseTree_symbol2Index = parseTree_nonTerminal2ListIndex(colTree, strSymbol)
    Else
        parseTree_symbol2Index = parseTree_terminal2Index(colTree, strSymbol)
    End If
End Function
Private Function parseTree_terminal2Index(ByRef colTree As Collection, _
                                          ByVal strTerminal As String) As Long
    '
    ' Converts terminal to index in parse tree item(2), or 0
    '
    '
    Dim colHandle As Collection
    parseTree_terminal2Index = 0
    On Error Resume Next
        Set colHandle = colTree.item(2)
        If Err.Number <> 0 Then Exit Function
        Set colHandle = colHandle.item(symbol2Key(strTerminal))
        If Err.Number <> 0 Then Exit Function
        parseTree_terminal2Index = colHandle.item(2)
    On Error GoTo 0
End Function
Private Function parseTree_whereList(ByRef colTree As Collection, _
                                     ByRef strSymbol As String, _
                                     ByVal booSymbol2Text As Boolean) As String
    '
    ' Return the comma-delimited list of nonterminals using strSymbol
    '
    '
    Dim colHandle(1 To 3) As Collection
    Dim lngIndex1 As Long
    Dim strNext As String
    Dim strOutstring As String
    With colTree
        lngIndex1 = parseTree_nonTerminal2ListIndex(colTree, strSymbol)
        If lngIndex1 <> 0 Then
            Set colHandle(1) = .item(1)
        Else
            lngIndex1 = parseTree_terminal2Index(colTree, strSymbol)
            If lngIndex1 = 0 Then
                OBJutilities.errorHandler 0, _
                                          "Internal programming error: " & _
                                          "undefined symbol " & strSymbol & " ", _
                                          "parseTree_whereList", _
                                          Me.Name
                Exit Function
            End If
            Set colHandle(1) = .item(2)
        End If
        Set colHandle(1) = colHandle(1).item(lngIndex1)
        Set colHandle(2) = .item(1)
        With colHandle(1)
            For lngIndex1 = 3 To .Count
                Set colHandle(3) = colHandle(2).item(.item(lngIndex1))
                strNext = colHandle(3).item(1)
                If booSymbol2Text Then
                    strNext = symbol2Text(strNext)
                    strNext = LCase$(Mid$(strNext, 1, 1)) & Mid$(strNext, 2)
                End If
                OBJutilities.append strOutstring, _
                                    IIf(lngIndex1 = .Count And lngIndex1 > 3, _
                                        " and ", ", "), _
                                    strNext, _
                                    booAppendInPlace:=True
            Next lngIndex1
        End With
        parseTree_whereList = strOutstring
    End With
End Function
#If SIXTY_DAY_EVALUATION Then
Private Function trialInstallation(ByVal STRsection As String, _
                                   ByVal strKey As String) As Boolean
    '
    ' Create (not so) top secret Registry entries
    '
    '
    On Error GoTo trialInstallation_Lbl1_installErrorHandler
        SaveSetting App.EXEName, _
                    STRsection, _
                    encodeInterface("installationDate", strKey), _
                    encodeInterface(Now, strKey)
    On Error GoTo 0
    trialInstallation = True
    Exit Function
trialInstallation_Lbl1_installErrorHandler:
    MsgBox ("The trial installation has failed. Contact: " & _
           vbNewLine & vbNewLine & _
           EGNADDRESS)
End Function
Private Function trialExpirationAnnunciate() As Boolean
    '
    ' Announce first use of trial: get user's buy-in
    '
    '
    MsgBox Replace(Replace(Replace(SIXTYDAY_INFO, "%ISWAS", "was"), _
                           "%EVALUATION_STARTDATE", _
                           Now), _
                   "%EVALUATION_ENDDATE", _
                   calculateEvalEndDate(Now)) & _
            vbNewLine & vbNewLine & _
            "This software has expired. Contact the following for a permanent copy of source " & _
            "or object:" & _
            vbNewLine & vbNewLine & _
            EGNADDRESS
End Function
Private Function trialInstallationAnnunciate() As Boolean
    '
    ' Announce first use of trial: get user's buy-in
    '
    '
    trialInstallationAnnunciate = _
        (MsgBox(ABOUTINFO & _
                vbNewLine & vbNewLine & _
                Replace(Replace(Replace(SIXTYDAY_INFO, "%ISWAS", "is being"), _
                                "%EVALUATION_STARTDATE", _
                                Now), _
                        "%EVALUATION_ENDDATE", _
                        calculateEvalEndDate(Now)) & _
                vbNewLine & vbNewLine & _
                "Click Yes to proceed. Click No if you do not wish to install this software.", _
                vbYesNo) _
         = _
         vbYes)
End Function
Private Function calculateEvalEndDate(ByVal datStart As Date) As Date
    '
    ' Calculates the sixty day end date
    '
    '
    calculateEvalEndDate = DateAdd("d", datStart, 60)
End Function
Private Function getEvaluationStartDate() As Date
    '
    ' Get the evaluation start date
    '
    '
    getEvaluationStartDate = decodeInterface(GetSetting(App.EXEName, _
                                                        Name, _
                                                        STRstartDateKeyEncoded), _
                                             STRencryptionKey)
End Function
Private Function getEvaluationEndDate() As Date
    '
    ' Get the evaluation end date
    '
    '
    getEvaluationEndDate = calculateEvalEndDate(getEvaluationStartDate)
End Function
#End If
Private Function symbols2XML(ByRef colTree As Collection, _
                             ByVal intIndex As Integer, _
                             ByVal strNewline As String) As String
    '
    ' Convert the nonterminals and the terminals to XML
    '
    '
    Dim colEntry As Collection
    Dim colItem As Collection
    Dim lngIndex1 As Long
    Dim strSymbol As String
    Dim strType As String
    Dim strValue As String
    Dim strXML As String
    Select Case intIndex
        Case 1: strType = "nonterminals"
        Case 2: strType = "terminals"
        Case Else
            OBJutilities.errorHandler 0, _
                                      "Internal programming error: invalid index", _
                                      "symbols2XML", _
                                      Name
            Exit Function
    End Select
    With colTree
        Set colEntry = .item(intIndex)     ' This is pretty random as compared to .Net DAMMIT
        With colEntry
            For lngIndex1 = 1 To .Count
                Set colItem = .item(lngIndex1)
                With colItem
                    If intIndex = 1 Then
                        strSymbol = .item(1)
                    Else
                        strSymbol = .item(1)
                        If OBJutilities.isQuoted(strSymbol) Then
                            strSymbol = OBJutilities.dequote(strSymbol)
                        End If
                        strValue = strSymbol
                    End If
                    If intIndex = 2 Or Mid(.item(1), Len(.item(1))) <> "]" Then
                        OBJutilities.append strXML, _
                                            strNewline, _
                                            OBJutilities.mkXMLElement(strSymbol, strValue, 0), _
                                            booAppendInPlace:=True
                    End If
                End With
            Next lngIndex1
            symbols2XML = OBJutilities.mkXMLElement(strType, _
                                                    strNewline & strXML & vbNewLine, _
                                                    0)
        End With
    End With
End Function
Private Function xmlHeader() As String
    '
    ' Return XML frontmatter
    '
    '
    Dim strBNF As String
    With OBJutilities
        If frmRefManOptions.IncludeXMLsource Then
            ' Include the BNF in the header comment if it is rather small
            strBNF = vbNewLine & vbNewLine & getBNF
        End If
        xmlHeader = .mkXMLComment(vbNewLine & _
                                  .string2Box("Backus-Naur Form definition of the language " & _
                                              frmRefManOptions.LanguageName & _
                                              strBNF & _
                                              vbNewLine) & _
                                  vbNewLine, _
                                  booMultipleLineEdit:=False) & vbNewLine & _
                    .mkXMLTag("BNF")
    End With
End Function
Private Function xmlTrailer() As String
    '
    ' Return XML backmatter
    '
    '
    xmlTrailer = OBJutilities.mkXMLTag("BNF", booEndTag:=True)
End Function
Private Function getBNF() As String
    '
    ' Reads the BNF
    '
    '
    With OBJutilities
        getBNF = .file2String(.attachPath(dirBNFlocation, filBNFlocation))
    End With
End Function
Private Function parseTree2XML(ByRef colTree As Collection, _
                               ByVal strXMLnewline As String, _
                               ByVal booInlineBNF As Boolean, _
                               ByVal booEndTagComment As Boolean) As String
    '
    ' Create the XML tree from the parse tree
    '
    '
    ' Unfortunately this code rather parallels parseTree2Rules as well as
    ' parseTree2RRDiagram in the railroad diagram form. This is a misfortune since
    ' if the structure of the parse tree changes all of these routines must change.
    '
    '
    Dim colHandle(1 To 2) As Collection
    Dim lngIndex1 As Long
    Dim lngIndex2 As Long
    Dim strXML As String
    parseTree2XML = ""
    If (colTree Is Nothing) Then
        Exit Function
    End If
    updateStatus "Creating the XML grammar for this language", 1
    With colTree
        Set colHandle(1) = .item(1)
        With colHandle(1)
            lngIndex2 = 1
            For lngIndex1 = 1 To .Count
                updateProgress "Creating the XML grammar for this language", _
                               "nonterminal", _
                               lngIndex1, _
                               .Count
                Set colHandle(2) = .item(lngIndex1)
                With colHandle(2)
                    If nonTerminal2NodeIndex(colTree, .item(1)) <> 0 Then
                        OBJutilities.append _
                            strXML, _
                            strXMLnewline, _
                            parseTree2XML_node2Rules(colTree, _
                                                       .item(1), _
                                                       nonTerminal2NodeIndex _
                                                       (colTree, .item(1)), _
                                                     strXMLnewline, _
                                                     booInlineBNF, _
                                                     booEndTagComment), _
                            booAppendInPlace:=True
                        lngIndex2 = lngIndex2 + 1
                    End If
                End With
            Next lngIndex1
            updateStatus "", -1
            updateStatus "XML grammar is complete"
            lblStatusProgress.Visible = False
        End With
    End With
    parseTree2XML = strXML
End Function
Private Function parseTree2XML_node2Rules(ByRef colTree As Collection, _
                                          ByVal strNonterminal As String, _
                                          ByVal lngNodeIndex As Long, _
                                          ByVal strXMLnewline As String, _
                                          ByVal booInlineBNF As Boolean, _
                                          ByVal booEndTagComment As Boolean) _
        As String
    '
    ' Expand one node to its XML format
    '
    '
    Dim colNodeHandle As Collection
    Dim colSubnodeHandle As Collection
    Dim intLast As Integer
    Dim intOutlineLevel As Integer
    Dim lngIndex1 As Long
    Dim strName As String
    Dim strOp As String
    Dim strSplit() As String
    Dim strTerminal As String
    Dim strWhere As String
    Dim strXML As String
    Dim strXML0 As String
    Dim strXML2 As String
    With colTree
        If lngNodeIndex = 0 Then
            parseTree2XML_node2Rules = "": Exit Function
        End If
        Set colNodeHandle = .item(lngNodeIndex)
        With colNodeHandle
            If (TypeOf .item(3) Is Collection) Then
                Set colSubnodeHandle = .item(3)
                If .item(1) <> "" Then
                    If booInlineBNF Then
                        strXML = OBJutilities.mkXMLTagWithAttributes("GS", _
                                                                     False, _
                                                                     "name", _
                                                                     .item(1), _
                                                                     "BNF", _
                                                                     scannedPhrase(USRscanned, _
                                                                                   .item(4), _
                                                                                   .item(5)))
                    Else
                        strXML = OBJutilities.mkXMLTagWithAttributes("GS", _
                                                                     False, _
                                                                     "name", _
                                                                     .item(1))
                    End If
                End If
                Select Case colSubnodeHandle.item(1)
                    Case OPERATOR_ALTERNATION:
                         strOp = "alternatives"
                    Case OPERATOR_MRE_ONETRIP:
                         strOp = "oneTripRepeat"
                    Case OPERATOR_OPTIONALSEQUENCE:
                         strOp = "optionalSequence"
                    Case OPERATOR_PRODUCTION:
                         strOp = "production"
                    Case OPERATOR_MRE_ZEROTRIP:
                         strOp = "zeroTripRepeat"
                    Case OPERATOR_SEQUENCE:
                         strOp = "sequence"
                    Case Else:
                        OBJutilities.errorHandler 0, _
                                                  "Internal programming error: " & _
                                                  "unsupported operator", _
                                                  "parseTree2XML_node2Rules", _
                                                  Me.Name
                        parseTree2XML_node2Rules = OBJutilities.mkXMLComment("Error")
                        Exit Function
                End Select
                strXML2 = OBJutilities.mkXMLTagWithAttributes("OP", False, "name", strOp)
                strXML0 = strXML
                strXML = IIf(strXML = "", "", strXML & strXMLnewline) & _
                         strXML2 & _
                         strXMLnewline & _
                         parseTree2XML_subnode2XML(colTree, _
                                                   colSubnodeHandle, _
                                                   strXMLnewline, _
                                                   booInlineBNF, _
                                                   booEndTagComment) & _
                         strXMLnewline & _
                         OBJutilities.xmlTag2EndTag(strXML2)
                If booEndTagComment Then
                    strXML = strXML & OBJutilities.mkXMLComment("End " & strOp)
                End If
                If strXML0 <> "" Then
                    strXML = strXML & _
                             vbNewLine & _
                             OBJutilities.xmlTag2EndTag(strXML0)
                    If booEndTagComment Then
                        strXML = strXML & OBJutilities.mkXMLComment("End " & .item(1))
                    End If
                End If
            Else
                lngIndex1 = .item(3)
                If lngIndex1 > 0 Then
                    strXML = _
                        OBJutilities.mkXMLElement(index2Name(colTree, lngIndex1, True), _
                                                  "", _
                                                  0)
                ElseIf lngIndex1 < 0 Then
                    strXML = _
                        OBJutilities.mkXMLElement(index2Name(colTree, Abs(lngIndex1), False), _
                                                  "", _
                                                  0)
                Else
                    strXML = OBJutilities.mkXMLComment("undefined syntax")
                End If
            End If
        End With
    End With
    parseTree2XML_node2Rules = strXML
    Exit Function
parseTree2XML_node2Rules_Lbl1_splitErrorHandler:
    OBJutilities.errorHandler 0, _
                              "Error in split: " & _
                              Err.Number & " " & Err.Description, _
                              "parseTree2XML_node2Rules", _
                              Me.Name
End Function
Private Function parseTree2XML_subnode2XML(ByRef colTree As Collection, _
                                           ByRef colSubnode As Collection, _
                                           ByVal strXMLnewline As String, _
                                           ByVal booInlineBNF As Boolean, _
                                           ByVal booEndTagComment As Boolean) _
        As String
    '
    ' Subnode (operator, operand1, operand2) to XML
    '
    '
    Dim lngIndex1 As Long
    Dim strOutstring As String
    On Error Resume Next
        lngIndex1 = colSubnode.item(2)
    On Error GoTo 0
    If lngIndex1 = 0 Then Exit Function
    strOutstring = parseTree2XML_node2Rules(colTree, _
                                            "", _
                                            lngIndex1, _
                                            strXMLnewline, _
                                            booInlineBNF, _
                                            booEndTagComment)
    lngIndex1 = 0
    On Error Resume Next
        lngIndex1 = colSubnode.item(3)
    On Error GoTo 0
    If lngIndex1 <> 0 Then
        strOutstring = strOutstring & _
                       IIf(strXMLnewline = "", "", vbNewLine) & _
                       parseTree2XML_node2Rules(colTree, _
                                                "", _
                                                lngIndex1, _
                                                strXMLnewline, _
                                                booInlineBNF, _
                                                booEndTagComment)
    End If
    parseTree2XML_subnode2XML = Replace(strOutstring, vbNewLine, strXMLnewline)
End Function
Private Sub mnuToolsRRdiagram_Click()
    If (COLparseTree Is Nothing) Then
        MsgBox "You need to parse the BNF first. Select an input BNF and " & _
               "click Parse the BNF on the menu."
        Exit Sub
    End If
    With frmRRdiagram
        Set .ParseTree = COLparseTree
        .Show vbModal
    End With
End Sub
Private Sub mnuToolsViewBNF_Click()
    OBJutilities.errorHandler 0, _
                              getBNF, _
                              "mnuToolsViewBNF_Click", _
                              Name
End Sub
Private Sub mnuViewAndChange_Click()
    showFile mnuViewAndChange.Tag
End Sub
Private Sub showFile(ByVal strFileid As String)
    With frmShowFile
    End With
End Sub
