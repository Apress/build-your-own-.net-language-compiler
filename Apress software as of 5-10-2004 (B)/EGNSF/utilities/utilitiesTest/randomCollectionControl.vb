Option Strict

' *********************************************************************
' *                                                                   *
' * Random collection and stress testing control                      *
' *                                                                   *
' *                                                                   *
' * This form and application exposes the following Friendly          *
' * properties.                                                       *
' *                                                                   *
' *                                                                   *
' *     *  EntryTypes: this read-only property returns the blank-de-  *
' *        limited list of entry types to be added to the collection. *
' *                                                                   *
' *        The members of this list are taken from the following:     *
' *        Boolean, Byte, Short, Integer, Single Long, String,        *
' *        Collection and StringBuilder.                              *
' *                                                                   *
' *     *  MaxDepth: this read-only property returns the maximum      *
' *        collection depth. When the user indicates unlimited depth  *
' *        (by moving the mac depth scrollbar as far as possible to   *
' *        the right) this property returns -1.                       *
' *                                                                   *
' *     *  MaxLength: this read-only property returns the maximum     *
' *        collection length, at each level                           *
' *                                                                   *
' *     *  ReadabilityProbability: this read-only property returns the*
' *        probability that a collection2String will return a readable*
' *        string, as a double-precision value, between 0 (never) and *
' *        1 (always).                                                *
' *                                                                   *
' *     *  StressTestCount: this read-only property returns the       *
' *        required number of stress tests to be conducted.           *
' *                                                                   *
' *                                                                   *
' *********************************************************************
Public Class randomCollectionControl
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        
        registry2Form

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
    Friend WithEvents lblMaxLength As System.Windows.Forms.Label
    Friend WithEvents hscMaxLength As System.Windows.Forms.HScrollBar
    Friend WithEvents hscMaxDepth As System.Windows.Forms.HScrollBar
    Friend WithEvents lblMaxDepth As System.Windows.Forms.Label
    Friend WithEvents lblEntryTypes As System.Windows.Forms.Label
    Friend WithEvents lstEntryTypes As System.Windows.Forms.ListBox
    Friend WithEvents cmdRandomize As System.Windows.Forms.Button
    Friend WithEvents cmdForm2Registry As System.Windows.Forms.Button
    Friend WithEvents cmdRegistry2Form As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents hscStressTestCount As System.Windows.Forms.HScrollBar
    Friend WithEvents lblStressTestCount As System.Windows.Forms.Label
    Friend WithEvents cmdClearSettings As System.Windows.Forms.Button
    Friend WithEvents hscReadabilityProbability As System.Windows.Forms.HScrollBar
    Friend WithEvents lblReadabilityProbability As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.lblMaxLength = New System.Windows.Forms.Label
Me.hscMaxLength = New System.Windows.Forms.HScrollBar
Me.hscMaxDepth = New System.Windows.Forms.HScrollBar
Me.lblMaxDepth = New System.Windows.Forms.Label
Me.lblEntryTypes = New System.Windows.Forms.Label
Me.lstEntryTypes = New System.Windows.Forms.ListBox
Me.cmdRandomize = New System.Windows.Forms.Button
Me.cmdForm2Registry = New System.Windows.Forms.Button
Me.cmdRegistry2Form = New System.Windows.Forms.Button
Me.cmdCancel = New System.Windows.Forms.Button
Me.cmdClose = New System.Windows.Forms.Button
Me.hscStressTestCount = New System.Windows.Forms.HScrollBar
Me.lblStressTestCount = New System.Windows.Forms.Label
Me.cmdClearSettings = New System.Windows.Forms.Button
Me.hscReadabilityProbability = New System.Windows.Forms.HScrollBar
Me.lblReadabilityProbability = New System.Windows.Forms.Label
Me.SuspendLayout()
'
'lblMaxLength
'
Me.lblMaxLength.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
Me.lblMaxLength.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblMaxLength.Location = New System.Drawing.Point(8, 8)
Me.lblMaxLength.Name = "lblMaxLength"
Me.lblMaxLength.Size = New System.Drawing.Size(680, 16)
Me.lblMaxLength.TabIndex = 0
Me.lblMaxLength.Text = "Maximum Length: %LENGTH"
Me.lblMaxLength.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'hscMaxLength
'
Me.hscMaxLength.Location = New System.Drawing.Point(8, 24)
Me.hscMaxLength.Name = "hscMaxLength"
Me.hscMaxLength.Size = New System.Drawing.Size(680, 16)
Me.hscMaxLength.TabIndex = 1
'
'hscMaxDepth
'
Me.hscMaxDepth.Location = New System.Drawing.Point(8, 64)
Me.hscMaxDepth.Name = "hscMaxDepth"
Me.hscMaxDepth.Size = New System.Drawing.Size(680, 16)
Me.hscMaxDepth.TabIndex = 3
'
'lblMaxDepth
'
Me.lblMaxDepth.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
Me.lblMaxDepth.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblMaxDepth.Location = New System.Drawing.Point(8, 48)
Me.lblMaxDepth.Name = "lblMaxDepth"
Me.lblMaxDepth.Size = New System.Drawing.Size(680, 16)
Me.lblMaxDepth.TabIndex = 2
Me.lblMaxDepth.Text = "Maximum Depth: %DEPTH"
Me.lblMaxDepth.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'lblEntryTypes
'
Me.lblEntryTypes.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
Me.lblEntryTypes.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblEntryTypes.Location = New System.Drawing.Point(8, 168)
Me.lblEntryTypes.Name = "lblEntryTypes"
Me.lblEntryTypes.Size = New System.Drawing.Size(680, 16)
Me.lblEntryTypes.TabIndex = 4
Me.lblEntryTypes.Text = "Entry Types"
Me.lblEntryTypes.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'lstEntryTypes
'
Me.lstEntryTypes.BackColor = System.Drawing.SystemColors.Control
Me.lstEntryTypes.Location = New System.Drawing.Point(8, 184)
Me.lstEntryTypes.Name = "lstEntryTypes"
Me.lstEntryTypes.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
Me.lstEntryTypes.Size = New System.Drawing.Size(680, 108)
Me.lstEntryTypes.TabIndex = 5
'
'cmdRandomize
'
Me.cmdRandomize.Location = New System.Drawing.Point(8, 296)
Me.cmdRandomize.Name = "cmdRandomize"
Me.cmdRandomize.Size = New System.Drawing.Size(112, 24)
Me.cmdRandomize.TabIndex = 6
Me.cmdRandomize.Text = "Randomize"
'
'cmdForm2Registry
'
Me.cmdForm2Registry.Location = New System.Drawing.Point(128, 296)
Me.cmdForm2Registry.Name = "cmdForm2Registry"
Me.cmdForm2Registry.Size = New System.Drawing.Size(112, 24)
Me.cmdForm2Registry.TabIndex = 7
Me.cmdForm2Registry.Text = "Save Settings"
'
'cmdRegistry2Form
'
Me.cmdRegistry2Form.Location = New System.Drawing.Point(248, 296)
Me.cmdRegistry2Form.Name = "cmdRegistry2Form"
Me.cmdRegistry2Form.Size = New System.Drawing.Size(108, 24)
Me.cmdRegistry2Form.TabIndex = 8
Me.cmdRegistry2Form.Text = "Restore Settings"
'
'cmdCancel
'
Me.cmdCancel.Location = New System.Drawing.Point(472, 296)
Me.cmdCancel.Name = "cmdCancel"
Me.cmdCancel.Size = New System.Drawing.Size(108, 24)
Me.cmdCancel.TabIndex = 9
Me.cmdCancel.Text = "Cancel"
'
'cmdClose
'
Me.cmdClose.Location = New System.Drawing.Point(584, 296)
Me.cmdClose.Name = "cmdClose"
Me.cmdClose.Size = New System.Drawing.Size(108, 24)
Me.cmdClose.TabIndex = 10
Me.cmdClose.Text = "Close"
'
'hscStressTestCount
'
Me.hscStressTestCount.Location = New System.Drawing.Point(8, 104)
Me.hscStressTestCount.Name = "hscStressTestCount"
Me.hscStressTestCount.Size = New System.Drawing.Size(680, 16)
Me.hscStressTestCount.TabIndex = 12
'
'lblStressTestCount
'
Me.lblStressTestCount.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
Me.lblStressTestCount.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblStressTestCount.Location = New System.Drawing.Point(8, 88)
Me.lblStressTestCount.Name = "lblStressTestCount"
Me.lblStressTestCount.Size = New System.Drawing.Size(680, 16)
Me.lblStressTestCount.TabIndex = 11
Me.lblStressTestCount.Text = "Stress Test Count: %COUNT"
Me.lblStressTestCount.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'cmdClearSettings
'
Me.cmdClearSettings.Location = New System.Drawing.Point(360, 296)
Me.cmdClearSettings.Name = "cmdClearSettings"
Me.cmdClearSettings.Size = New System.Drawing.Size(108, 24)
Me.cmdClearSettings.TabIndex = 13
Me.cmdClearSettings.Text = "Clear Settings"
'
'hscReadabilityProbability
'
Me.hscReadabilityProbability.Location = New System.Drawing.Point(8, 144)
Me.hscReadabilityProbability.Name = "hscReadabilityProbability"
Me.hscReadabilityProbability.Size = New System.Drawing.Size(680, 16)
Me.hscReadabilityProbability.TabIndex = 15
'
'lblReadabilityProbability
'
Me.lblReadabilityProbability.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
Me.lblReadabilityProbability.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblReadabilityProbability.Location = New System.Drawing.Point(8, 128)
Me.lblReadabilityProbability.Name = "lblReadabilityProbability"
Me.lblReadabilityProbability.Size = New System.Drawing.Size(680, 16)
Me.lblReadabilityProbability.TabIndex = 14
Me.lblReadabilityProbability.Text = "Probability that collection2String is readable: %PROBABILITY"
Me.lblReadabilityProbability.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'randomCollectionControl
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(696, 325)
Me.ControlBox = False
Me.Controls.Add(Me.hscReadabilityProbability)
Me.Controls.Add(Me.lblReadabilityProbability)
Me.Controls.Add(Me.cmdClearSettings)
Me.Controls.Add(Me.hscStressTestCount)
Me.Controls.Add(Me.lblStressTestCount)
Me.Controls.Add(Me.cmdClose)
Me.Controls.Add(Me.cmdCancel)
Me.Controls.Add(Me.cmdRegistry2Form)
Me.Controls.Add(Me.cmdForm2Registry)
Me.Controls.Add(Me.cmdRandomize)
Me.Controls.Add(Me.lstEntryTypes)
Me.Controls.Add(Me.lblEntryTypes)
Me.Controls.Add(Me.hscMaxDepth)
Me.Controls.Add(Me.lblMaxDepth)
Me.Controls.Add(Me.hscMaxLength)
Me.Controls.Add(Me.lblMaxLength)
Me.Name = "randomCollectionControl"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
Me.Text = "Random Collection and Stress Test Parameters"
Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Form data "

Private _OBJutilities As utilities.utilities
Private _OBJwindowsUtilities As windowsUtilities.windowsUtilities

Private Const ENTRY_TYPES As String = "Byte Boolean Short Integer " & _
                                      "Long Single Double String " & _
                                      "Collection StringBuilder"
Private Const DEFAULT_STRESSTEST_COUNT As Integer = 10    
Private Const READABILITY_PROBABILITIES As String = _
        "Never,Almost never,Rarely,Somewhat rarely,Sometimes,Often," & _
        "Frequently,Very frequently,Nearly always,Always"
Private Const DEFAULT_READABILITY_PROBABILITY As Double = .15        
                                      
Private INTmaxLength As Integer
Private INTmaxDepth As Integer
Private INTstressTestCount As Integer
Private DBLreadabilityProbability As Double
Private STRentryTypes As String                                      

#End Region ' Form data

#Region " Form Friend properties and methods "

    Friend ReadOnly Property EntryTypes As String
        Get
            Return STRentryTypes
        End Get        
    End Property    
    
    Friend ReadOnly Property MaxDepth As Integer
        Get
            Return(INTmaxDepth)
        End Get        
    End Property    
    
    Friend ReadOnly Property MaxLength As Integer
        Get
            Return(INTmaxLength)
        End Get        
    End Property    
    
    Friend ReadOnly Property ReadabilityProbability As Double
        Get
            With hscReadabilityProbability
                Return(roundLimitProbability(_OBJutilities.histogram(.Value, _
                                                                     dblRangeMin:=0, _
                                                                     dblRangeMax:=1, _
                                                                     dblValueMin:=CDbl(.Minimum), _
                                                                     dblValueMax:=CDbl(.Minimum))))                                                     
            End With
        End Get        
    End Property    
    
    Friend ReadOnly Property StressTestCount As Integer
        Get
            Return(hscStressTestCount.Value)
        End Get        
    End Property    

#End Region ' Form properties

#Region " Form events "

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        registry2Form
        Hide
    End Sub

    Private Sub cmdClearSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearSettings.Click
        clearRegistry
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        form2Registry
        Hide
    End Sub

    Private Sub cmdForm2Registry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdForm2Registry.Click
        form2Registry
    End Sub

    Private Sub cmdRandomize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRandomize.Click
        Randomize
    End Sub

    Private Sub cmdRegistry2Form_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRegistry2Form.Click
        registry2Form
    End Sub

    Private Sub hscMaxDepth_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles hscMaxDepth.ValueChanged
        maxDepth2Object
    End Sub

    Private Sub hscReadabilityProbability_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles hscReadabilityProbability.ValueChanged
        readabilityProbability2Object                          
    End Sub

    Private Sub hscMaxLength_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles hscMaxLength.ValueChanged
        maxLength2Object       
    End Sub

    Private Sub hscStressTestCount_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles hscStressTestCount.ValueChanged
        stressTestCount2Object
    End Sub

    Private Sub lstEntryTypes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstEntryTypes.Click
        entryTypes2Object
    End Sub

    Private Sub randomCollectionControl_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        properties2Screen
    End Sub

#End Region ' Form events

#Region " General procedures "
    
    ' -----------------------------------------------------------------
    ' Clear the Registry and set defaults
    '
    '
    Private Function clearRegistry As Boolean
        _OBJwindowsUtilities.clearRegistryKey(Application.ProductName, _
                                              Me.Name)
        registry2Form    
        properties2Screen                                          
    End Function    
    
    ' -----------------------------------------------------------------
    ' Transfer the selected entry types to the object
    '
    '
    Private Sub entryTypes2Object
        Dim intIndex1 As Integer
        Dim strEntries As String
        STRentryTypes = ""
        With lstEntryTypes
            For intIndex1 = 0 To .SelectedIndices.Count - 1
                _OBJutilities.append(STRentryTypes, _
                                     " ", _
                                     CStr(.Items(.SelectedIndices.Item(intIndex1))))
            Next intIndex1                 
        End With 
    End Sub    
       
    ' -----------------------------------------------------------------
    ' Form settings to Registry
    '
    '    
    Private Function form2Registry As Boolean
        Try
            With hscMaxDepth
                SaveSetting(Application.ProductName, _
                            Me.Name, _
                            .Name, _
                            CStr(.Value)) 
            End With        
            With hscMaxLength
                SaveSetting(Application.ProductName, _
                            Me.Name, _
                            .Name, _
                            CStr(.Value)) 
            End With        
            With hscStressTestCount
                SaveSetting(Application.ProductName, _
                            Me.Name, _
                            .Name, _
                            CStr(.Value)) 
            End With   
            SaveSetting(Application.ProductName, _
                        Me.Name, _
                        hscReadabilityProbability.Name, _
                        CStr(readabilityProbability2Double)) 
            With lstEntryTypes
                Dim intIndex1 As Integer
                Dim strEntries As String
                For intIndex1 = 0 To .SelectedIndices.Count - 1
                    _OBJutilities.append(strEntries, _
                                         " ", _
                                         CStr(.Items(.SelectedIndices.Item(intIndex1))))
                Next                 
                SaveSetting(Application.ProductName, _
                            Me.Name, _
                            .Name, _
                            strEntries) 
            End With                
        Catch
            _OBJutilities.errorHandler("Could not save Registry settings: " & _
                                       Err.Number & " " & Err.Description, _
                                       Me.Name, _
                                       "form2Registry", _
                                       "Returning False")
            Return(False)                                       
        End Try          
        Return(True)  
    End Function    

    ' -----------------------------------------------------------------
    ' Return the selected entry types
    '
    '
    Private Function getEntryTypes As String
        With lstEntryTypes.SelectedIndices
            Dim intIndex1 As Integer
            Dim strOutstring As String
            For intIndex1 = 1 To .Count
                _OBJutilities.append(strOutstring, _
                                     " ", _
                                     CStr(lstEntryTypes.Items(.Item(intIndex1))))
            Next intIndex1      
            Return(strOutstring)      
        End With        
    End Function      
    
    ' ------------------------------------------------------------------
    ' Load the entry types  
    '
    '  
    Private Sub loadEntryTypes
        Dim strSplit() As String
        Try
            strSplit = split(ENTRY_TYPES, " ")
        Catch  
            _OBJutilities.errorHandler("Could not split: " & _
                                       Err.Number & " " & Err.Description, _
                                       Me.Name, _
                                       "loadEntryTypes", _
                                       "Cannot load entry types")
            Return                                       
        End Try        
        Dim intIndex1 As Integer
        lstEntryTypes.Items.Clear
        For intIndex1 = 1 To UBound(strSplit)
            Try
                With lstEntryTypes
                    .Items.Add(strSplit(intIndex1))
                    .SelectedIndex = .Items.Count - 1
                    .Focus
                    .Refresh
                End With                
            Catch  
                _OBJutilities.errorHandler("Could not split: " & _
                                           Err.Number & " " & Err.Description, _
                                           Me.Name, _
                                           "loadEntryTypes", _
                                           "Cannot load entry types")
                Return                                       
            End Try            
        Next intIndex1        
        With lstEntryTypes
            .SelectedIndex = CInt(Iif(.Items.Count = 0, -1, 0))
        End With        
    End Sub
    
    ' -----------------------------------------------------------------
    ' Maximum depth to its label and object state
    '
    '
    Private Sub maxDepth2Object
        With hscMaxDepth
            scrollValue2Object(CStr(Iif(.Value >= CInt(.Maximum * .96), "no limit", .Value)), _
                               lblMaxDepth, _
                               "%DEPTH")
            INTmaxDepth = .Value                               
        End With                          
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Maximum length to its label and object state
    '
    '
    Private Sub maxLength2Object
        With hscMaxLength
            scrollValue2Object(CStr(Iif(.Value = .Maximum, "no limit", .Value)), _
                              lblMaxLength, _
                              "%LENGTH")
            INTmaxLength = .Value                              
        End With                          
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Properties to screen
    '
    '
    Private Sub properties2Screen
        Try
            hscMaxDepth.Value = INTmaxDepth
        Catch: End Try
        maxDepth2Object
        Try
            hscMaxLength.Value = INTmaxLength
        Catch: End Try
        maxLength2Object
        Try
            hscStressTestCount.Value = INTstressTestCount
        Catch: End Try
        stressTestCount2Object
        Try
            With hscReadabilityProbability
                .Value = CInt(_OBJutilities.histogram(DBLreadabilityProbability, _
                                                      dblRangeMin:=.Minimum, _
                                                      dblRangeMax:=.Maximum, _
                                                      dblValueMax:=1))
            End With                                        
        Catch: End Try
        readabilityProbability2Object
        loadEntryTypes
        With lstEntryTypes
            .Focus
            .ClearSelected
            Dim intIndex1 As Integer
            Dim intIndex2 As Integer
            For intIndex1 = 1 To _OBJutilities.words(STRentryTypes)
                intIndex2 = _OBJwindowsUtilities.searchListBox(lstEntryTypes, _
                                                               _OBJutilities.word(STRentryTypes, _
                                                                                  intIndex1))
                If intIndex2 >= 0 Then
                    .SelectedIndex = intIndex2
                    .Refresh
                End If                                                                               
            Next intIndex1            
        End With        
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Convert the readability probability from integer to double
    ' precision
    '
    '
    Private Function readabilityProbability2Double As Double
        With hscReadabilityProbability
            Return(_OBJutilities.histogram(CDbl(.Value), _
                                            dblRangeMin:=0, _
                                            dblRangeMax:=1, _
                                            dblValueMin:=.Minimum, _
                                            dblValueMax:=.Maximum))
        End With                                                                   
    End Function    
    
    ' -----------------------------------------------------------------
    ' Readability to its label
    '
    '
    Private Sub readabilityProbability2Object
        With hscReadabilityProbability
            DBLreadabilityProbability = roundLimitProbability(readabilityProbability2Double)                                                                               
            scrollValue2Object(CStr(Math.Round(roundLimitProbability(DBLreadabilityProbability), 3) * 100) & "% " & _
                              "(" & readabilityProbability2String(DBLreadabilityProbability) & ")", _
                              lblReadabilityProbability, _
                              "%PROBABILITY")
        End With                          
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Return the identifier for the probability
    '
    '
    Private Function readabilityProbability2String(ByVal dblProbability As Double) _
            As String
        With _OBJutilities
            Return(.item(READABILITY_PROBABILITIES, _
                         CInt(.histogram(dblProbability, _
                                         dblRangeMin:=1, _
                                         dblRangeMax:=.items(READABILITY_PROBABILITIES, _
                                                             ",", _
                                                             False), _
                                         dblValueMin:=0, _
                                         dblValueMax:=1)), _
                         ",", _
                         False))
        End With                                                           
    End Function    
    
    ' -----------------------------------------------------------------
    ' Return readability granularity based on the probability words
    '
    '
    Private Function readabilityProbabilityTolerance As Double
        With _OBJutilities
            Return(.histogram(2, _
                              dblRangeMin:=0, _
                              dblRangeMax:=1, _
                              dblValueMin:=1, _
                              dblValueMax:=.items(READABILITY_PROBABILITIES, _
                                                  ",", _
                                                  False)))
        End With                                                           
    End Function    
    
    ' -----------------------------------------------------------------
    ' Registry settings to form
    '
    '    
    Private Function registry2Form As Boolean
        Try
            With hscMaxDepth
                INTmaxDepth = CInt(GetSetting(Application.ProductName, _
                                              Me.Name, _
                                              .Name, _
                                              CStr(CInt((.Maximum - .Minimum + 1) * .10 _
                                                         + _
                                                         .Minimum _
                                                         - _
                                                         1))))
            End With        
            With hscMaxLength
                INTmaxLength = CInt(GetSetting(Application.ProductName, _
                                               Me.Name, _
                                               .Name, _
                                               CStr(CInt((.Maximum - .Minimum + 1) * .10 _
                                                         + _
                                                         .Minimum _
                                                         - _
                                                         1)))) 
            End With        
            With hscStressTestCount
                INTstressTestCount = CInt(GetSetting(Application.ProductName, _
                                                     Me.Name, _
                                                     .Name, _
                                                     CStr(DEFAULT_STRESSTEST_COUNT)))
            End With   
            With hscReadabilityProbability
                DBLreadabilityProbability = CDbl(GetSetting(Application.ProductName, _
                                                 Me.Name, _
                                                 .Name, _
                                                 CStr(DEFAULT_READABILITY_PROBABILITY)))
            End With   
            STRentryTypes = GetSetting(Application.ProductName, _
                                       Me.Name, _
                                       lstEntryTypes.Name, _
                                       ENTRY_TYPES)
        Catch
            _OBJutilities.errorHandler("Could not get Registry settings: " & _
                                       Err.Number & " " & Err.Description, _
                                       Me.Name, _
                                       "registry2Form", _
                                       "Returning False")
            Return(False)                                       
        End Try          
        Return(True)  
    End Function    
    
    ' -----------------------------------------------------------------
    ' Round the probability, close to 0 and 1
    '
    '
    Private Function roundLimitProbability(ByVal dblProbability As Double) As Double
        Dim dblTolerance As Double = readabilityProbabilityTolerance
        If dblProbability <= dblTolerance Then
            Return(0)
        ElseIf dblProbability >= 1 - dblTolerance Then
            Return(1)            
        End If        
        Return(dblProbability)
    End Function    
    
    ' -----------------------------------------------------------------
    ' Move scrollbar's value to its label
    '
    '
    ' Note that this method stores the original value of the scroll
    ' bar's label in the Tag of that label.
    '
    '
    Private Sub scrollValue2Object(ByVal strValue As String, _
                                  ByVal lblLabel As Label, _
                                  ByVal strKeyword As String)
        With lblLabel
            If (.Tag Is Nothing) Then .Tag = .Text
            .Text = Replace(CStr(.Tag), strKeyword, strValue)
        End With                                          
    End Sub                   
    
    ' -----------------------------------------------------------------
    ' Transfer value to label
    '
    '
    Private Sub stressTestCount2Object
        scrollValue2Object(CStr(hscStressTestCount.Value), _
                          lblStressTestCount, _
                          "%COUNT")
        INTstressTestCount = hscStressTestCount.Value                          
    End Sub                   
    
#End Region ' General procedures

End Class
