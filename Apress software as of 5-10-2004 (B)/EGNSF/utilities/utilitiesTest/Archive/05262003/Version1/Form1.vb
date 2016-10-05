Option Strict

' *********************************************************************
' *                                                                   *
' * utilitiesTest                                                     *
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
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents lblUtilities As System.Windows.Forms.Label
    Friend WithEvents lstUtilities As System.Windows.Forms.ListBox
    Friend WithEvents lstStatus As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdStatusZoom As System.Windows.Forms.Button
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents cmdAbout As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.cmdClose = New System.Windows.Forms.Button
Me.lblUtilities = New System.Windows.Forms.Label
Me.lstUtilities = New System.Windows.Forms.ListBox
Me.lstStatus = New System.Windows.Forms.ListBox
Me.Label1 = New System.Windows.Forms.Label
Me.cmdStatusZoom = New System.Windows.Forms.Button
Me.lblProgress = New System.Windows.Forms.Label
Me.cmdAbout = New System.Windows.Forms.Button
Me.SuspendLayout()
'
'cmdClose
'
Me.cmdClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdClose.Location = New System.Drawing.Point(504, 440)
Me.cmdClose.Name = "cmdClose"
Me.cmdClose.Size = New System.Drawing.Size(120, 24)
Me.cmdClose.TabIndex = 0
Me.cmdClose.Text = "Close"
'
'lblUtilities
'
Me.lblUtilities.BackColor = System.Drawing.Color.Khaki
Me.lblUtilities.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblUtilities.Font = New System.Drawing.Font("Georgia", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblUtilities.Location = New System.Drawing.Point(8, 8)
Me.lblUtilities.Name = "lblUtilities"
Me.lblUtilities.Size = New System.Drawing.Size(616, 16)
Me.lblUtilities.TabIndex = 1
Me.lblUtilities.Text = "Utilities Available: double-click to see information about any utility"
Me.lblUtilities.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'lstUtilities
'
Me.lstUtilities.BackColor = System.Drawing.SystemColors.Control
Me.lstUtilities.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lstUtilities.ItemHeight = 14
Me.lstUtilities.Location = New System.Drawing.Point(8, 24)
Me.lstUtilities.Name = "lstUtilities"
Me.lstUtilities.Size = New System.Drawing.Size(616, 228)
Me.lstUtilities.TabIndex = 2
'
'lstStatus
'
Me.lstStatus.BackColor = System.Drawing.SystemColors.Control
Me.lstStatus.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lstStatus.ItemHeight = 14
Me.lstStatus.Location = New System.Drawing.Point(7, 276)
Me.lstStatus.Name = "lstStatus"
Me.lstStatus.Size = New System.Drawing.Size(616, 130)
Me.lstStatus.TabIndex = 4
'
'Label1
'
Me.Label1.BackColor = System.Drawing.Color.Khaki
Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.Label1.Font = New System.Drawing.Font("Georgia", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.Label1.Location = New System.Drawing.Point(7, 256)
Me.Label1.Name = "Label1"
Me.Label1.Size = New System.Drawing.Size(616, 20)
Me.Label1.TabIndex = 3
Me.Label1.Text = "Status"
Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'cmdStatusZoom
'
Me.cmdStatusZoom.Location = New System.Drawing.Point(560, 256)
Me.cmdStatusZoom.Name = "cmdStatusZoom"
Me.cmdStatusZoom.Size = New System.Drawing.Size(64, 20)
Me.cmdStatusZoom.TabIndex = 5
Me.cmdStatusZoom.Text = "Zoom"
'
'lblProgress
'
Me.lblProgress.BackColor = System.Drawing.Color.Khaki
Me.lblProgress.Font = New System.Drawing.Font("Georgia", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblProgress.Location = New System.Drawing.Point(7, 416)
Me.lblProgress.Name = "lblProgress"
Me.lblProgress.Size = New System.Drawing.Size(616, 16)
Me.lblProgress.TabIndex = 6
Me.lblProgress.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
Me.lblProgress.Visible = False
'
'cmdAbout
'
Me.cmdAbout.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdAbout.Location = New System.Drawing.Point(8, 440)
Me.cmdAbout.Name = "cmdAbout"
Me.cmdAbout.Size = New System.Drawing.Size(120, 24)
Me.cmdAbout.TabIndex = 7
Me.cmdAbout.Text = "About"
'
'Form1
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(630, 467)
Me.ControlBox = False
Me.Controls.Add(Me.cmdAbout)
Me.Controls.Add(Me.lblProgress)
Me.Controls.Add(Me.cmdStatusZoom)
Me.Controls.Add(Me.lstStatus)
Me.Controls.Add(Me.Label1)
Me.Controls.Add(Me.lstUtilities)
Me.Controls.Add(Me.lblUtilities)
Me.Controls.Add(Me.cmdClose)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
Me.Name = "Form1"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
Me.Text = "utilitiesTest"
Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Form data "

    ' ***** Constants *****
    ' --- Filenames
    Private UTILITIES_FILENAME As String = "utilities"
    Private WINDOWSUTILITIES_FILENAME As String = "windowsUtilities"
    ' --- About info  
    Private Const ABOUTINFO As String = _
    "utilitiesTest" & vbNewline & vbNewline & vbNewline & _
    "This form and application tests the utilities and the windowsUtilities libraries." & vbnewline & vbnewline & _
    "The utilities library is a collection of more than fifty general-purpose methods " & vbnewline & _
    "for Web based and Windows parsing, collection management, simple IO and other " & vbNewline & _    
    "common tasks." & vbNewline & vbNewline & _
    "The windowsUtilities library is a smaller collection of tools for classic Windows " & vbnewline & _
    "applications." & vbNewline & vbNewline & _
    "This application was developed by: " & vbnewline & vbNewline & _
    "Edward G. Nilges" & vbNewLine & _
    "Sarita Flats" & vbNewLine & _
    "39 Gordon St." & vbNewLine & _
    "PO Box 16334" & vbNewLine & _
    "Suva, Fiji" & vbNewLine & _
    "Intl dialing: 679 330 0084" & vbNewLine & _
    "spinoza1111@yahoo.COM" & vbNewLine & _
    "http://members.screenz.com/edNilges"
    ' --- File identifiers
    Private Const UTILITIESFILEID As String = "..\..\utilities.VB"
    Private Const WINDOWSUTILITIESFILEID As String = "..\..\..\windowsUtilities\windowsUtilities.VB"

    ' ***** Shared utilities *****
    Private Shared _OBJutilities As utilities.utilities
    Private Shared _OBJwindowsUtilities As windowsUtilities.windowsUtilities

    ' ***** Procedure information *****
    ' --- Procedure type
    Private Enum ENUprocedureType
        subroutineProcedure
        functionProcedure
        propertyProcedure                
        invalidProcedure
    End Enum		
    ' --- Procedure information
    Private Structure TYPoverload
        Dim intStartIndex As Integer        ' Start index from one
        Dim intLength As Integer            ' Length
        Dim enuType As ENUprocedureType     ' Sub, func, property or not valid
        Dim colParameters As Collection     ' Parameters: no key
                                            ' Item(1): return type
                                            ' Items 2..n: subcollection
                                            '      Item(1): parameter name
                                            '      Item(2): True (ByVal) or False (ByRef)
                                            '      Item(3): dimensions or 0 for a scalar
                                            '      Item(4): type
    End Structure
    Private Structure TYPprocedure
        Dim strName As String		        ' Function, subroutine or property name
        Dim strText As String               ' Procedure's text
        Dim usrOverload() As TYPoverload    ' One or more noncontiguous overloads
        Dim strAbstract As String           ' From comment header
        Dim colUses As Collection           ' List of procedures used by this procedure
        Dim colUsed As Collection           ' List of procedures using this procedure
    End Structure
    Private Structure TYPindexedProcedures
        Dim colIndex As Collection          ' Index: key is name: data is index in table
        Dim usrProcedures() As TYPprocedure 
    End Structure    
    Private USRutilitiesProcedures As TYPindexedProcedures
    Private USRwindowsProcedures As TYPindexedProcedures

#End Region ' " Form data "

#Region " Form events "

    Private Sub cmdAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAbout.Click
        showAbout
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        closer
        End
    End Sub

    Private Sub cmdStatusZoom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStatusZoom.Click
        zoomInterface(lstStatus)
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Show
        Refresh
        If Not getProcedures(UTILITIESFILEID, USRutilitiesProcedures) Then closer
    End Sub

    Private Sub lstUtilities_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstUtilities.DoubleClick
        displayTestForm(_OBJutilities.item(CStr(lstUtilities.Items(lstUtilities.SelectedIndex)), _
                                           1, _
                                           ":", _
                                           False)) 
    End Sub

#End Region ' " Form events "

#Region " General procedures "

    ' -----------------------------------------------------------------
    ' Get one-line abstract
    '
    '
    Private Function abstract2Oneline(ByVal strAbstract As String) As String
        Dim strSplit() As String  
        Try
            strSplit = split(strAbstract, vbNewline)
        Catch ex As Exception
            errorHandler("Can't split string: " & Err.Number & " " & Err.Description, _
                         "", _
                         "Returning a null string" & vbNewline & vbNewline & _
                         ex.ToString)
            Return("")                         
        End Try       
        Dim intCount As Integer
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim strNext As String
        For intIndex1 = 0 To UBound(strSplit)
            intCount = 0
            For intIndex2 = 1 To Len(strSplit(intIndex1))
                strNext = Mid(strSplit(intIndex1), intIndex2, 1)
                If strNext >= "A" AndAlso strNext <= "Z" _
                   OrElse _
                   strNext >= "a" AndAlso strNext <= "z" Then
                    intCount += 1
                End If                    
            Next intIndex2   
            If intCount >= Len(strSplit(intIndex1)) * .75 Then 
                Return(strSplit(intIndex1))
            End If                     
        Next intIndex1         
    End Function    

    ' -----------------------------------------------------------------
    ' Initialize the procedure table
    '
    '
    Private Function allocateProcinfo(ByRef usrProcedures As TYPindexedProcedures) As Boolean
        With usrProcedures
            Try
                Redim .usrProcedures(0)
                .colIndex = New Collection
            Catch ex As Exception
                errorHandler("Cannot allocate procedure array or index: " & _
                            Err.Number & " " & Err.Description, _
                            "allocateProcinfo", _
                            ex.toString)
                Return(False)
            End Try
            Return(True)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Close the application and return to Windows
    '
    '
    Private Sub closer  
        Dispose
        End
    End Sub

    ' -----------------------------------------------------------------
    ' Determine procedure type
    '
    '
    Private Function determineProcedureType(ByVal strHeader As String) As ENUprocedureType
        If Instr(strHeader & " ", "SUB ", CompareMethod.Text) <> 0 Then
            Return(ENUprocedureType.subroutineProcedure)
        ElseIf Instr(strHeader & " ", "FUNCTION ", CompareMethod.Text) <> 0 Then
            Return(ENUprocedureType.functionProcedure)
        ElseIf Instr(strHeader & " ", "PROPERTY ", CompareMethod.Text) <> 0 Then
            Return(ENUprocedureType.propertyProcedure)
        End If
        errorHandler("Cannot determine procedure type in " & _
                     _OBJutilities.enquote(_OBJutilities.string2Display(strHeader)), _
                     "determineProcedureType", _
                     "Returning invalid")
        Return(ENUprocedureType.invalidProcedure)
    End Function
    
    ' -----------------------------------------------------------------
    ' Display the test form
    '
    '
    Private Sub displayTestForm(ByVal strUtility As String)

    End Sub    

    ' -----------------------------------------------------------------
    ' Interface to the error handler
    '
    '
    Private Sub errorHandler(ByVal strMessage As String, _
                             ByVal strProcedure As String, _
                             ByVal strHelp As String, _
                             Optional ByVal booInfo As Boolean = False)
        Dim strFullMessage As String
        _OBJutilities.errorHandler(strMessage, _
                                   Me.Name, _
                                   strProcedure, _
                                   strHelp, _
                                   strFullMessage, _
                                   booInfo:=booInfo)
        If Not booInfo Then updateStatusListBox
        updateStatusListBox(_OBJutilities.string2Box(_OBJutilities.soft2HardParagraph(strFullMessage, _
                                                                                      72), _
                                                     CStr(Iif(booInfo, "I N F O", "E R R O R"))), _
                            booDateStamp:=False)
    End Sub

    ' -----------------------------------------------------------------
    ' Expand the procedure table
    '
    '
    Private Function expandProctable(ByRef usrIndexedProcedures As TYPindexedProcedures, _
                                     ByVal strName As String, _
                                     ByVal strText As String, _
                                     ByVal intStartIndex As Integer, _
                                     ByVal intLength As Integer, _
                                     ByVal enuType As ENUprocedureType, _
                                     ByVal strAbstract As String, _
                                     ByVal colParameters As Collection) As Boolean
        With usrIndexedProcedures
            If (.colIndex Is Nothing) AndAlso Not allocateProcinfo(usrIndexedProcedures) Then Return(False)
            Dim intIndex1 As Integer
            Try
                intIndex1 = CInt(.colIndex(strName))
            Catch: End Try            
            If intIndex1 = 0 Then
                ' New procedure
                Try
                    Redim Preserve .usrProcedures(UBound(.usrProcedures) + 1)
                    With .usrProcedures(UBound(.usrProcedures))
                        .colUsed = New Collection: .colUses = New Collection
                        Redim .usrOverload(0)
                    End With
                Catch ex As Exception
                    errorHandler("Cannot expand procedure table: " & _
                                Err.Number & " " & Err.Description, _
                                "expandProctable", _
                                "Returning False" & _
                                vbNewline & vbNewline & _
                                ex.toString)
                    Return(False)
                End Try
                With .usrProcedures(UBound(.usrProcedures))
                    .strName = strName
                    .strAbstract = strAbstract
                    .strText = strText
                End With
                intIndex1 = UBound(.usrProcedures)
                Try
                    .colIndex.Add(intIndex1, strName)
                Catch ex As Exception
                    errorHandler("Cannot expand procedure table's index: " & _
                                Err.Number & " " & Err.Description, _
                                "expandProctable", _
                                "Returning False" & _
                                vbNewline & vbNewline & _
                                ex.toString)
                    Return(False)
                End Try                
            Else
                .usrProcedures(UBound(.usrProcedures)).strText &= vbNewline & strText                
            End If
            With .usrProcedures(intIndex1)
                Try
                    Redim Preserve .usrOverload(UBound(.usrOverload) + 1)
                Catch ex As Exception
                    errorHandler("Cannot expand procedure overload table: " & _
                                Err.Number & " " & Err.Description, _
                                "expandProctable_overload_", _
                                "Returning False" & _
                                vbNewline & vbNewline & _
                                ex.toString)
                    Return(False)
                End Try     
                With .usrOverload(UBound(.usrOverload))   
                    .intStartIndex = intStartIndex
                    .intLength = intLength
                    .enuType = .enuType
                    .colParameters = colParameters
                End With        
            End With                                
            Return(True)
        End With                                            
    End Function
    
    ' -----------------------------------------------------------------
    ' Get directory from file id
    '
    '
    Private Function getFileDirectory(ByVal strFileid As String) As String
        Dim strFileDirectory As String
        Dim strTitle As String
        parseFileid(strFileid, strFileDirectory, strTitle)
        Return(strFileDirectory)
    End Function

    ' -----------------------------------------------------------------
    ' Get name from file id
    '
    '
    Private Function getFileName(ByVal strFileid As String) As String
        Dim strDirectory As String
        Dim strTitle As String
        parseFileid(strFileid, strDirectory, strTitle)
        Dim intIndex1 As Integer = InstrRev(strTitle, ".")
        If intIndex1 = 0 Then Return(strTitle)
        Return(Mid(strTitle, 1, intIndex1 - 1)) 
    End Function

    ' -----------------------------------------------------------------
    ' Get title from file id
    '
    '
    Private Function getFileTitle(ByVal strFileid As String) As String
        Dim strDirectory As String
        Dim strTitle As String
        parseFileid(strFileid, strDirectory, strTitle)
        Return(strTitle)
    End Function

    ' -----------------------------------------------------------------
    ' Get all procedures  
    '
    '
    Private Function getProcedures(ByVal strFileId As String, _
                                   ByRef usrProcedureInfo As TYPindexedProcedures) As Boolean  
        Dim booExists As Boolean
        Dim strDir As String
        Dim strInfile As String
        Try
            strDir = CurDir
            ChDir(Application.StartupPath)
            booExists = _OBJutilities.fileExists(strFileid) 
            If booExists Then 
                strInfile = getProcedures_normalizeNewline(_OBJutilities.file2String(strFileId))
            End If                
            ChDir(strDir)
        Catch: End Try
        If Not (getProcedures_file2Procedures(strInfile, _
                                                usrProcedureInfo, _
                                                getFileName(strFileId))) Then
            Return(False)
        End If
        getProcedures_procedures2Listbox(usrProcedureInfo.usrProcedures, lstUtilities) 
        Return(True)
    End Function

    ' -----------------------------------------------------------------
    ' Get procedures from file on behalf of getProcedures
    '
    '
    Private Function getProcedures_file2Procedures(ByVal strInfile As String, _
                                                   ByRef usrProcedures As TYPindexedProcedures, _
                                                   ByVal strFileName As String) As Boolean
        Dim strActivity As String = "Obtaining procedures from file"
        updateStatusListBox(strActivity, 1)
        Dim objRegex(1) As System.Text.RegularExpressions.Regex
        Try
            objRegex(0) = New System.Text.RegularExpressions.Regex _
                              (_OBJutilities.commonRegularExpressions("vbCommentBlock") & _ 
                               getProcedures_removeScopes _
                               (_OBJutilities.commonRegularExpressions("vbProcedureHeaderNET"))) 
        Catch ex As Exception
            errorHandler("Cannot create regular expression: " & Err.Number & " " & Err.Description, _
                         "instrument_getFiles_getProcedures", _
                         ex.ToString)
            Return(False)
        End Try
        Dim booFoundEnd As Boolean
        Dim enuType As ENUprocedureType
        Dim intIndex1 As Integer = 1
        Dim intStartIndex As Integer
        Dim intEndIndex As Integer
        Dim objMatch As System.Text.RegularExpressions.Match
        Dim strName As String
        Dim strText As String
        Dim strType As String
        Dim strWork As String
        Do While intIndex1 <= Len(strInfile)
            updateProgress(lblProgress, _
                           strActivity, _
                           "character", _
                           intIndex1, _
                           Len(strInfile))
            Try
                objMatch = objRegex(0).Match(strInfile, _
                                             getProcedures_findProbableStart(strInfile, _
                                                                             intIndex1) - 1)
            Catch ex As Exception
                errorHandler("Regex match for start of procedure failed: " & _
                             Err.Number & " " & Err.Description, _
                             "instrument_getFiles_getProcedures", _
                             "Terminating search" & vbNewline & vbNewline & ex.ToString)
                Exit Do
            End Try
            With objMatch
                If .Length <> 0 Then
                    intStartIndex = .Index + 1
                    strName = _OBJutilities.word(.Value, _OBJutilities.words(.Value))
                Else
                    intIndex1 = Len(strInfile) + 1: Exit Do
                End If 
                enuType = determineProcedureType(.Value)
                booFoundEnd = True
                strType = procedureType2Keyword(enuType)
                Try
                    objRegex(1) = Nothing
                    objRegex(1) = New System.Text.RegularExpressions.Regex _
                                      (_OBJutilities.commonRegularExpressions("newlineWebWindows") & _
                                       "[ ]*End " & strType)
                    objMatch = objRegex(1).Match(strInfile, intStartIndex) 
                Catch ex As Exception
                    errorHandler("Regex match for end of procedure failed: " & _
                                 Err.Number & " " & Err.Description, _
                                 "instrument_getFiles_getProcedures", _
                                 "Ending procedure at eof" & vbNewline & vbNewline & ex.ToString)
                    booFoundEnd = False                                 
                End Try       
                With objMatch
                    booFoundEnd = booFoundEnd AndAlso .Length <> 0
                    If booFoundEnd Then 
                        intEndIndex = .Index + 1 + .Length
                    Else
                        intEndIndex = Len(strInfile) + 1
                    End If                        
                    strText = Mid(strInfile, intStartIndex, intEndIndex - intStartIndex)
                End With                         
                If Not expandProctable(usrProcedures, _ 
                                       strName, _
                                       strText, _
                                       intStartIndex, _
                                       Len(strText), _
                                       enuType, _
                                       getProcedures_parseAbstract(Mid(strInfile, .Index + 1, 8192)), _ 
                                       getProcedures_parseParameters(Mid(strInfile, .Index + 1 + .Length, 8192))) Then 
                    Exit Do
                End If
                updateStatusListBox("Found " & strType & " " & strName)
                intIndex1 = intEndIndex + 1
            End With
        Loop
        updateStatusListBox(-1, strActivity & " complete")
        lblProgress.Visible = False
        Return(intIndex1 > Len(strInfile))
    End Function
    
    ' -----------------------------------------------------------------
    ' Find probable start of procedure based on structure of the actual
    ' utilities file
    '
    '
    ' We locate the comment bar (consisting of apostrophe and dashes
    ' as above) that precedes the string "\n    Public". This assists
    ' the regular expression scan in finding the next procedure quickly.
    '
    ' intIndex is returned without change if the probable start cannot be
    ' located.
    '
    ' Note that this method will skip over Private and Friend procedures as 
    ' intended.
    '
    '
    Private Function getProcedures_findProbableStart(ByVal strInfile As String, _
                                                     ByVal intIndex As Integer) _
                     As Integer
        Dim intIndex1 As Integer = Instr(intIndex, strInfile, vbNewline & "    Public") 
        If intIndex1 = 0 Then Return(intIndex)
        Dim intIndex2 As Integer = InstrRev(strInfile, vbNewline & "    ' ----------", intIndex1)
        If intIndex2 > intIndex Then Return(intIndex2)
        Return(intIndex1) 
    End Function                     

    ' -----------------------------------------------------------------
    ' Normalize the newline on behalf of getProcedures
    '
    '
    Private Function getProcedures_normalizeNewline(ByVal strInfile As String) As String
        updateStatusListBox("Normalizing the new line string", 1)
        Dim strNewline As String = _OBJutilities.determineNewline(strInfile, intMaxLength:=1024)
        If strNewline = "" Then
            errorHandler("Input file does not appear to have a consistent newline", _
                         "getProcedures_normalizeNewline", _
                         "Using the standard value")
            Return(strInfile)                         
        End If        
        updateStatusListBox(-1, "The newline string is " & _OBJutilities.string2Display(strNewline))
        Return(Replace(strInfile, strNewline, vbNewline))
    End Function
        
    ' -----------------------------------------------------------------
    ' Parse the abstract on behalf of getProcedures
    '
    '
    Private Function getProcedures_parseAbstract(ByVal strSource As String) As String
        ' --- Get the abstract
        Dim intIndex1 As Integer  
        Dim strSplit() As String 
        Try
            strSplit = Split(Mid(strSource, intIndex1 + Len(vbNewline)), vbNewline)
        Catch ex As Exception
            errorHandler("Cannot split string: " & Err.Number & " " & Err.Description, _
                         "getProcedures_parseAbstract", _
                         "Returning a null abstract")
            Return("")
        End Try
        Dim strOutstring As String
        Dim strText As String
        For intIndex1 = 0 To UBound(strSplit)
            If Not getProcedures_parseComment(strSplit(intIndex1), strText) Then Exit For
            _OBJutilities.append(strOutstring, vbNewline, strText)
        Next intIndex1
        Return(strOutstring)
    End Function

    ' -----------------------------------------------------------------
    ' Parse comment on behalf of getProcedures
    '
    '
    Private Function getProcedures_parseComment(ByVal strSourceLine As String, _
                                                ByRef strText As String) As Boolean
        Dim strTextTrim As String = Trim(strSourceLine)
        If Mid(strTextTrim, 1, 1) <> "'" Then Return(False)
        strText = Trim(Mid(strTextTrim, 2))
        Return(True)
    End Function
        
    ' -----------------------------------------------------------------
    ' Parse the parameters and the As clause on behalf of getProcedures
    '
    '
    ' This method returns a new collection. Item(1) will be Nothing or
    ' the As clause's type (Object when the variable type is not 
    ' specified). Items(2)..n will be subcollections:
    '
    '
    '      *  Item(1) of each subcollection will be the variable name
    '
    '      *  Item(2) will be True (item is passed ByVal) or False 
    '         (item is passed ByRef)
    '
    '      *  Item(3) will be its number of dimensions (0 for nonarrays)
    '
    '      *  Item(4) will be its variable type (Object when the variable type
    '         is not specified, even for subroutines)
    '
    '
    Private Function getProcedures_parseParameters(ByVal strSource As String) As Collection
        Dim colNew As Collection
        Try
            colNew = New Collection
        Catch ex As Exception
            errorHandler("Can't create collection: " & Err.Number & " " & Err.Description, _
                         "", _
                         "Returning Nothing")
            Return(Nothing)                         
        End Try        
        ' --- Assemble possibly continued statement
        Dim intIndex1 As Integer = 1
        Dim intIndex2 As Integer
        Dim intIndex3 As Integer
        Dim strHeader As String
        Do While intIndex1 <= Len(strSource)
            intIndex2 = Instr(intIndex1, strSource & " _" & vbNewline, " _" & vbNewline)
            intIndex3 = Instr(intIndex1, strSource & vbNewline, vbNewline)
            strHeader &= Mid(strSource, intIndex1, Math.Min(intIndex2, intIndex3) - intIndex1)
            If intIndex2 > intIndex3 Then Exit Do
            intIndex1 = intIndex2 + Len(" _" & vbNewline)
        Loop        
        ' --- Parse the parameters
        Dim booByVal As Boolean
        Dim colSubentry as Collection
        Dim intDimensions As Integer
        Dim objRE As New System.Text.RegularExpressions.Regex _
            (_OBJutilities.commonRegularExpressions("vbParameterDefNet"))
        Dim objMatch As System.Text.RegularExpressions.Match  
        Dim strName As String
        Dim strSplit() As String
        Dim strType As String
        intIndex1 = 1
        Do While intIndex1 <= Len(strSource)
            Try
                objMatch = objRE.Match(strHeader, intIndex1)
            Catch ex As Exception
                errorHandler("Can't apply regular expression: " & _
                             Err.Number & " " & Err.Description, _
                             "getProcedures_parseParameters", _
                             "Returning Nothing")
                Return(Nothing)                         
            End Try   
            With objMatch
                If .Length = 0 Then Exit Do
                Try
                    colSubentry = New Collection
                Catch ex As Exception
                    errorHandler("Can't create subcollection: " & _
                                Err.Number & " " & Err.Description, _
                                "getProcedures_parseParameters", _
                                "Returning Nothing")
                    Return(Nothing)                         
                End Try                
                booByVal = True
                intIndex2 = 1
                Select Case UCase(_OBJutilities.word(.Value, intIndex2))
                    Case "BYVAL": intIndex2 += 1
                    Case "BYREF": booByVal = False: intIndex2 += 1
                    Case Else:
                End Select
                strName = _OBJutilities.word(.Value, intIndex2)
                intDimensions = 0
                Try
                    strSplit = split(strName, "(")
                Catch ex As Exception
                    errorHandler("Can't split: " & _
                                Err.Number & " " & Err.Description, _
                                "getProcedures_parseParameters", _
                                "Returning Nothing")
                    Return(Nothing)                         
                End Try                    
                strName = strSplit(0)
                intDimensions = 0
                If UBound(strSplit) > 0 Then
                    Try
                        strSplit = split(strSplit(1), ",")
                    Catch ex As Exception
                        errorHandler("Can't split: " & _
                                    Err.Number & " " & Err.Description, _
                                    "getProcedures_parseParameters", _
                                    "Returning Nothing")
                        Return(Nothing)                         
                    End Try                    
                    intDimensions = UBound(strSplit) + 1
                End If
                intIndex2 += 1      
                strType = "Object"
                If intIndex2 < _OBJutilities.words(.Value) _
                   AndAlso _
                   UCase(_OBJutilities.word(.Value, intIndex2)) = "AS" Then
                    strType = _OBJutilities.word(.Value, intIndex2 + 1)
                End If                
                With colSubentry
                    Try
                        .Add(strName) 
                        .Add(booByVal)
                        .Add(intDimensions)
                        .Add(strType)
                        colNew.Add(colSubentry)
                    Catch ex As Exception
                        errorHandler("Can't make subentry and/or add it to collection: " & _
                                    Err.Number & " " & Err.Description, _
                                    "getProcedures_parseParameters", _
                                    "Nothing")
                        Return(Nothing)                         
                    End Try                    
                End With
                intIndex1 = .Index + 1 + .Length
                If Mid(strHeader, intIndex1, 1) <> "," Then Exit Do                            
            End With                 
        Loop            
        strType = "Object"
        intIndex1 = Instr(intIndex1, strSource, " As ")
        If intIndex1 <> 0 Then strType = _OBJutilities.item(Mid(strSource, intIndex1), _
                                                            2, _
                                                            _OBJutilities.range2String(0, 32), _ 
                                                            True)
        Try
            colNew.Add(strType, , 1)
        Catch ex As Exception
            errorHandler("Can't add return type to collection: " & _
                        Err.Number & " " & Err.Description, _
                        "getProcedures_parseParameters", _
                        "Nothing")
            Return(Nothing)                         
        End Try    
        Return(colNew)    
    End Function

    ' -----------------------------------------------------------------
    ' Transfer procedures to List box on behalf of getProcedures
    '
    '
    Private Sub getProcedures_procedures2Listbox(ByRef usrProcedureTable() As TYPprocedure, _
                                                 ByRef lstBox As ListBox)  
        With lstBox
            .Items.Clear: .Refresh
            Dim intIndex1 As Integer
            For intIndex1 = 1 To UBound(usrProcedureTable)
                With usrProcedureTable(intIndex1)
                    lstBox.Items.Add(.strName & ": " & abstract2OneLine(.strAbstract))
                End With
                .SelectedIndex = .Items.Count - 1
                .Refresh
            Next intIndex1
            .SelectedIndex = CInt(Iif(.Items.Count = 0, -1, 0))
            .Refresh
        End With
    End Sub
    
    ' -----------------------------------------------------------------
    ' Remove Private and Friend scopes from the procedures re
    '
    '
    Private Function getprocedures_removeScopes(ByVal strRE As String) As String
        Dim strReplace As String = "(Public )|(Private )|(Friend )"
        Dim intIndex1 As Integer = Instr(strRE, strReplace)
        If intIndex1 = 0 Then
            errorHandler("Cannot remove scopes from procedure regular expression: " & _
                         "expected contents not found", _
                         "", _
                         "Continuing with unmodified regular expression")
            Return(strRE)                         
        End If        
        Return(Replace(strRE, strReplace, "(Public )"))
    End Function    

    ' -----------------------------------------------------------------
    ' Parse path into directory and file title
    '
    '
    Private Sub parseFileid(ByVal strPath As String, _
                            ByRef strDirectory As String, _
                            ByRef strFileTitle As String)
        Dim intIndex1 As Integer = InstrRev(strPath, "\")
        strDirectory = "": strFileTitle = strPath
        If intIndex1 = 0 Then Return
        strDirectory = Mid(strPath, 1, intIndex1 - 1)
        strFileTitle = Mid(strPath, intIndex1 + 1)
    End Sub
    
    ' -----------------------------------------------------------------
    ' Convert procedure type to keyword
    '
    '
    Private Function procedureType2Keyword(ByVal enuType As ENUprocedureType) As String
        Select Case enuType
            Case ENUprocedureType.functionProcedure: Return("Function")
            Case ENUprocedureType.invalidProcedure: Return("")
            Case ENUprocedureType.propertyProcedure: Return("Property")
            Case ENUprocedureType.subroutineProcedure: Return("Sub")
            Case Else
                errorHandler("Programming error: unexpected procedure type " & _
                             enuType.ToString, _
                             "procedureType2Keyword", _
                             "Returning a null string")
                Return("")                             
        End Select
    End Function    

    ' -----------------------------------------------------------------
    ' Convert procedure type enumerator to VB-compatible name
    '
    '
    Private Function procType2String(ByVal enuType As ENUprocedureType) As String
        Select Case enuType
            Case ENUprocedureType.functionProcedure: Return("Function")
            Case ENUprocedureType.propertyProcedure: Return("Property")
            Case ENUprocedureType.subroutineProcedure: Return("Sub")
            Case Else: Return("Invalid")
        End Select
    End Function

    ' -----------------------------------------------------------------
    ' Show the Easter Egg
    '
    '
    Private Sub showAbout 
        MsgBox(ABOUTINFO)
    End Sub

    ' -----------------------------------------------------------------
    ' Adjust form boundaries, around bottom and right side of control
    '
    '
    Private Sub shrinkWrap(ByVal ctlControl As Control)
        Width =  _OBJwindowsUtilities.controlRight(ctlControl) _
                 + _
                 _OBJwindowsUtilities.Grid * 2
        Height = _OBJwindowsUtilities.controlBottom(ctlControl) _
                 + _
                 _OBJwindowsUtilities.Grid * 4
        CenterToParent
        Refresh
    End Sub

    ' -----------------------------------------------------------------
    ' Update progress report
    '
    '
    Private Sub updateProgress(ByVal lblProgress As Label, _
                               ByVal strActivity As String, _
                               ByVal strEntity As String, _
                               ByVal intEntityNumber As Integer, _
                               ByVal intEntityCount As Integer)
        Dim strPercentComplete As String
        If intEntityCount <> 0 Then
            strPercentComplete = ": " & _
                                 CStr(Math.Round(intEntityNumber/intEntityCount * 100, 2)) & _
                                 "% complete"
        End If
        updateStatusListBox(strActivity & " at " & strEntity & " " & _
                            intEntityNumber & " of " & intEntityCount & _
                            strPercentComplete)
        With lblProgress
            .Width = CInt(_OBJutilities.histogram(intEntityNumber, _
                                             dblRangeMax:=lstStatus.Width, _
                                             dblValueMax:=intEntityCount))
            .Visible = True
            .Refresh
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Update the status list box
    '
    '
    ' --- No level change
    Private Overloads Sub updateStatusListBox(ByVal strMessage As String, _
                                              Optional ByVal booDateStamp As Boolean = True)
        updateStatusListBox(0, strMessage, 0)
    End Sub
    ' --- Change level before displaying the message
    Private Overloads Sub updateStatusListBox(ByVal intLevelChangeBefore As Integer, _
                                              ByVal strMessage As String, _
                                              Optional ByVal booDateStamp As Boolean = True)
        updateStatusListBox(intLevelChangeBefore, strMessage, 0)
    End Sub
    ' --- Change level after displaying the message
    Private Overloads Sub updateStatusListBox(ByVal strMessage As String, _
                                              ByVal intLevelChangeAfter As Integer, _
                                              Optional ByVal booDateStamp As Boolean = True)
        updateStatusListBox(0, strMessage, intLevelChangeAfter)
    End Sub
    ' --- Reset level with no display
    Private Overloads Sub updateStatusListBox()
        lstStatus.Tag = 0
    End Sub
    ' --- Change level before and after displaying the message
    Private Overloads Sub updateStatusListBox(ByVal intLevelChangeBefore As Integer, _
                                              ByVal strMessage As String, _
                                              ByVal intLevelChangeAfter As Integer, _
                                              Optional ByVal booDateStamp As Boolean = True)
        Dim intLevel As Integer
        Try
            intLevel = CInt(lstStatus.Tag)
        Catch: End Try
        intLevel += intLevelChangeBefore
        If intLevel < 0 Then intLevel = 0
        If strMessage <> "" Then
            With lstStatus
                Dim strSplit() As String 
                Try
                    strSplit = split(strMessage, vbNewline)
                Catch ex As Exception
                    MsgBox("Error in " & _
                           Me.Name & ".updateStatusListBox" & ": " & _
                           "cannot split: " & _
                           Err.Number & " " & Err.Description)
                    Return
                End Try
                Dim strDateStamp As String
                If booDateStamp Then strDateStamp = Trim(CStr(Now))
                Dim strIndent As String = _OBJutilities.copies(" ", intLevel * 5)
                Dim intIndex1 As Integer
                If UBound(strSplit) = 0 Then
                    .Items.Add(strIndent & strDateStamp & " " & _
                               Replace(_OBJutilities.soft2HardParagraph(strSplit(0), bytLineWidth:=50), _
                                       vbNewline, _
                                       vbNewline & _
                                       strIndent & _
                                       _OBJutilities.copies(" ", Len(strDateStamp) + 1)))
                Else
                    For intIndex1 = 0 To UBound(strSplit)
                        .Items.Add(strIndent & _
                                   strDateStamp & " " & _
                                   strSplit(intIndex1))
                    Next intIndex1
                End If
                .SelectedIndex = .Items.Count - 1
                .Focus
                .Refresh
            End With
        End If
        intLevel += intLevelChangeAfter
        lstStatus.Tag = intLevel
    End Sub

    ' -----------------------------------------------------------------
    ' Zoom control
    '
    '
    Private Sub zoomInterface(ByVal ctlZoomed As Control)
        Dim ctlZoom As zoom.zoom
        Try
            ctlZoom = New zoom.zoom(ctlZoomed, CDbl(1.5), CDbl(5))
        Catch ex As Exception
            errorHandler("Cannot create zoom control: " & _
                         Err.Number & " " & Err.Description, _
                         "zoomInterface", _
                         ex.ToString)
            Return
        End Try
        With ctlZoom
            .ShowZoom
            .Dispose
        End With
    End Sub

#End Region ' " General procedures "

#Region " My own objects "

' *********************************************************************
' *                                                                   *
' * testForm     Utilities test form                                  *
' *                                                                   *
' *                                                                   *
' * This object exposes the following properties and methods.         *
' *                                                                   *
' *                                                                   *
' *      addTestControl: this method adds one testControl to the      *
' *           collection of test controls displayed for the utility:  *
' *           often, only one such control is displayed, but more than*
' *           one control will be displayed when there are SeeAlso    *
' *           utilities that perform the inverse of the utility       *
' *           functions.                                              *
' *                                                                   *
' *      clearTestControls: this method clears the set of test        *
' *           controls that are associated with the utility.          *
' *                                                                   *
' *      Abstract: this write-only property can be set to the utility *
' *           abstract: it defaults to "No information is available." *
' *                                                                   *
' *      SeeAlso: this write-only property can be set to a list of    *
' *           one or more related utilities: it defaults to a null    *
' *           string.  Usually the SeeAlso utilities are inverse      *
' *           functions, where the base utility has the name a2b, and *
' *           the inverse utility has the name b2a.                   *
' *                                                                   *
' *      UsedBy: this write-only property can be set to a list of     *
' *           one or more utilities that call the displayed utility:  *
' *           it defaults to a null string.                           *
' *                                                                   *
' *      Uses: this write-only property can be set to a list of       *
' *           one or more utilities that the displayed utility calls: *
' *           it defaults to a null string.                           *
' *                                                                   *
' *                                                                   *
' *********************************************************************

Private Class testForm
    Inherits System.Windows.Forms.Form
    
    ' ***** Object state *****
    Private COLtestControls As Collection   ' Of testing controls
    
    ' -----------------------------------------------------------------
    ' Add test control
    '
    '
    Friend Function addTestControl(ByVal objTestControl As testControl) _
            As Boolean

    End Function            
    
End Class

' *********************************************************************
' *                                                                   *
' * testControl     Utilities test control                            *
' *                                                                   *
' *********************************************************************

Private Class testControl
    Inherits Control
    
    ' ***** Object state *****
End Class

#End Region ' " My own objects "

End Class
