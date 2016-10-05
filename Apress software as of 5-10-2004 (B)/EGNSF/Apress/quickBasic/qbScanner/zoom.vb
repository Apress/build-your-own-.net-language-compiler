Option Strict

' ********************************************************************
' *                                                                  *
' * zoom     Zoom information in list box or other control           *
' *                                                                  *
' *                                                                  *
' * This class creates and modally displays a dynamic form, that     *
' * shows the item list of a list box, or, when passed a control     *
' * other than a list box, the Text of that control.  For more       *
' * information, see zoom.DOC.                                       *
' *                                                                  *
' ********************************************************************

Public Class zoom
    Inherits Control
    
    ' ***** Shared *****
    Private Shared _OBJutilities As utilities 
    Private Shared _OBJwindowsUtilities As windowsUtilities 
    
    ' ***** Object state *****
    Private Structure TYPstate
        Dim booUsable As Boolean            ' True: object may be used
        Dim frmZoom As Form                 ' The zoomer form
        Dim txtZoom As TextBox              ' The zoomer text bo
        Dim cmdClose As Button              ' The close button: see also CMDclose below
        Dim ctlControl As Control           ' Listbox or other control
        Dim objSize As zoomSize             ' Just in time size
        Dim strItemSep As String            ' Newline or other size
    End Structure    
    Private USRstate As TYPstate
    Private WithEvents CMDclose As Button    ' Linked to cmdClose above
    
    ' ***** Public procedures ****************************************
    
    ' ----------------------------------------------------------------
    ' Assign and set the control
    '
    '
    Public Property Control() As Control
        Get
            If Not checkUsable_("Control get") Then Return(Nothing)
            Return(USRstate.ctlControl)
        End Get
        Set(ByVal ctlNewValue As Control)
            If Not checkUsable_("Control set") Then Return
            With USRstate
                .ctlControl = ctlNewValue
                Dim objWidth As Object = .objSize.Width
                Dim objHeight As Object = .objSize.Height
                .objSize = Nothing
                .objSize = New zoomSize(ctlNewValue.Size)
                With .objSize
                    .Width = objWidth: .Height = objHeight
                End With     
            End With       
        End Set                
    End Property    
    
    ' ----------------------------------------------------------------
    ' Dispose object and its base
    '
    '
    ' Note that we do not call checkUsable_ to check usability, since
    ' in the normal course of events, the user will click the Close
    ' button and this will dispose the object and its form.  We merely
    ' check the usability of the object and return if the object is not
    ' usable.
    '
    '
    Public Overloads Sub dispose
        With USRstate
            If Not .booUsable Then Return
            .booUsable = False
            .cmdClose.Visible = False: .cmdClose.Dispose: .cmdClose = Nothing
            .txtZoom.Visible = False: .txtZoom.Dispose: .txtZoom = Nothing
            .frmZoom.Visible = False: .frmZoom.Dispose: .frmZoom = Nothing
        End With        
        USRstate.booUsable = False
        MyBase.Dispose
    End Sub    
    
    ' ----------------------------------------------------------------
    ' Return and set the list box item separator
    '
    '
    Public Property ItemSep() As String
        Get
            If Not checkUsable_("ItemSep get") Then Return("")
            Return(USRstate.strItemSep)
        End Get        
        Set(ByVal strNewValue As String) 
            If Not checkUsable_("ItemSep set") Then Return
            usrSTATE.strItemSep = strNewValue
        End Set        
    End Property    
    
    ' ----------------------------------------------------------------
    ' Object constructor
    '
    '
    ' --- Create object with no zoomed control
    Public Sub new() 
        new_
        setDefaultSize_
    End Sub    
    ' --- Create object with control (recommended)
    Public Sub new(ByVal ctlControl As Control) 
        new_
        setDefaultSize_
        Me.Control = ctlControl
    End Sub    
    ' --- Create object with control (recommended) and relative dimensions
    Public Sub new(ByVal ctlControl As Control, _
                   ByVal dblWidthChange As Double, _
                   ByVal dblHeightChange As Double) 
        new_
        If Not Me.setSize(dblWidthChange, dblHeightChange) Then
            USRstate.booUsable = False: Return
        End If        
        Me.Control = ctlControl
    End Sub                
    ' --- Create object with control (recommended) and absolute dimensions as pair
    Public Sub new(ByVal ctlControl As Control, _
                   ByVal intWidth As Integer, _
                   ByVal intHeight As Integer) 
        new_
        If Not Me.setSize(intWidth, intHeight) Then
            USRstate.booUsable = False: Return
        End If
        Me.Control = ctlControl
    End Sub                
    ' --- Create object with control (recommended) and absolute dimensions as Size
    Public Sub new(ByVal ctlControl As Control, _
                   ByVal objSize As Size) 
        new_
        If Not Me.setSize(objSize) Then
            USRstate.booUsable = False: Return
        End If
        Me.Control = ctlControl
    End Sub                
    ' --- Common logic
    Private Sub new_
        With USRstate
            .strItemSep = vbNewline
            .booUsable = mkZoomForm_
            Dim objSize As New Size(.frmZoom.Width, .frmZoom.Height)
            .objSize = New zoomSize(objSize)
        End With        
    End Sub    

    ' ----------------------------------------------------------------
    ' Set the size
    '
    '
    ' --- Set to size object
    Public Overloads Function setSize(ByVal objSize As Size) As Boolean
        If Not checkUsable_("setSize") Then Return(False)
        USRstate.objSize.Width = CInt(objSize.Width)
        USRstate.objSize.Height = CInt(objSize.Height)
        Return(True)
    End Function   
    ' --- Set to fixed values 
    Public Overloads Function setSize(ByVal intWidth As Integer, _
                                      ByVal intHeight As Integer) As Boolean
        If Not checkUsable_("setSize") Then Return(False)
        USRstate.objSize.Width = intWidth
        USRstate.objSize.Height = intHeight
        Return(True)
    End Function                           
    ' --- Set to relative values        
    Public Overloads Function setSize(ByVal dblWidthChange As Double, _
                                      ByVal dblHeightChange As Double) As Boolean
        If Not checkUsable_("setSize") Then Return(False)
        USRstate.objSize.Width = dblWidthChange
        USRstate.objSize.Height = dblHeightChange
        Return(True)
    End Function                                 
    
    ' ---------------------------------------------------------------
    ' Show the form
    '
    '
    Public Function showZoom As Boolean 
        If Not checkUsable_("showZoom") Then Return(False)
        With USRstate
            If Not (.ctlControl Is Nothing) Then
                If (Typeof .ctlControl Is ListBox) Then
                    .txtZoom.Text = _
                    _OBJwindowsUtilities.listBox2String(CType(.ctlControl, ListBox), _
                                                        strSep:=.strItemSep)
                Else
                    .txtZoom.Text = .ctlControl.Text                            
                End If          
            End If                
            .frmZoom.Size = .objSize.Size
            resizeControls_
            .frmZoom.ShowDialog
            Return(True)
        End With        
    End Function    
    
    ' ---------------------------------------------------------------
    ' Return the zoom text box
    '
    '
    Public ReadOnly Property ZoomTextBox As TextBox 
        Get
            If Not checkUsable_("ZoomTextBox get") Then Return(Nothing)
            Return(USRstate.txtZoom)
        End Get        
    End Property      

    ' ***** Private procedures ***************************************
    
    ' ----------------------------------------------------------------
    ' Check usability
    '
    '
    Private Function checkUsable_(ByVal strProcedure As String) As Boolean
        If Not USRstate.booUsable Then
            errorHandler_("Object is not usable", _
                          strProcedure, _
                          "This object has encountered a serious internal " & _
                          "error and can't be used")
            Return(False)                          
        End If        
        Return(True)
    End Function    
        
    ' ----------------------------------------------------------------
    ' Interface to error handler
    '
    '
    Private Sub errorHandler_(ByVal strMessage As String, _
                              ByVal strProcedure As String, _
                              ByVal strHelp As String)
        _OBJutilities.errorHandler(strMessage, _
                                   "zoom", strProcedure, _
                                   strHelp)
    End Sub            
    
    ' ----------------------------------------------------------------
    ' Make the zoomer form
    '
    '
    Private Function mkZoomForm_ As Boolean
        Try
            With USRstate
                ' --- Create the form
                .frmZoom = New Form
                With .frmZoom
                    .Visible = False
                    .StartPosition = FormStartPosition.CenterScreen
                    .FormBorderStyle = FormBorderStyle.Fixed3D
                    .ControlBox = False
                    .Text = "Zoom"
                End With            
                ' --- Create the close button
                CMDclose = New Button
                AddHandler CMDclose.Click, AddressOf cmdClose_Click
                .cmdClose = CMDclose
                .frmZoom.Controls.Add(.cmdClose)
                With .cmdClose
                    .Text = "Close"
                    .Width *= 2
                End With                
                ' --- Create the text box
                .txtZoom = New TextBox
                .frmZoom.Controls.Add(.txtZoom)
                With .txtZoom
                    .ReadOnly = True
                    .Multiline = True
                    .ScrollBars = ScrollBars.Both
                    .Text = " "
                    .Font = New Font("Courier New", 8.25, FontStyle.Regular)
                    .Left = _OBJwindowsUtilities.Grid
                    .Top = .Left
                End With                
            End With        
        Catch
            errorHandler_("Cannot create the zoom form: " & _
                          Err.Number & " " & Err.Description, _
                          "mkZoomForm_", _
                          "Marking object not usable")
            USRstate.booUsable = False: Return(False)                          
        End Try
        Return(True)
    End Function    
    
    ' ----------------------------------------------------------------
    ' Proportionally, resize the controls
    '
    '
    Private Sub resizeControls_
        With USRstate.cmdClose
            .Left = USRstate.frmZoom.Width \ 2 _
                    - _
                    .Width \ 2
            .Top = USRstate.frmZoom.Height _
                    - _
                    _OBJwindowsUtilities.Grid * 5 _
                    - _
                    .Height                            
        End With        
        With USRstate.txtZoom
            .Width = USRstate.frmZoom.Width - .Left * 3
            .Height = USRstate.cmdClose.Top - .Left * 2
        End With        
    End Sub    
    
    ' ----------------------------------------------------------------
    ' Assign the default size
    '
    '
    Private Sub setDefaultSize_
        With USRstate
            If Not (.ctlControl Is Nothing) Then
                .objSize = New zoomSize(.ctlControl.Size): Return
            End If
            .objSize = New zoomSize(.frmZoom.Size): Return                        
        End With        
    End Sub                          
    
    ' ***** Event handlers *******************************************
    
    ' ----------------------------------------------------------------
    ' User clicks my close button
    '
    '
    Private Sub cmdClose_Click(ByVal objSender As Object, _
                               ByVal objEventArgs As System.EventArgs)
        If Not checkUsable_("cmdClose_Click") Then Return
        USRstate.frmZoom.Hide
    End Sub                               
       
    ' ****************************************************************
    ' *                                                              *
    ' * zoomSize     Just-in-time zoom form size                     *
    ' *                                                              *
    ' *                                                              *
    ' * This class represents the intended size of the zoom form.  It*
    ' * exposes the following methods and properties.                *
    ' *                                                              *
    ' *                                                              *
    ' *      *  Height: this read-write property returns and may be  *
    ' *         set to the planned height of the zoom form, and it   *
    ' *         may be an integer (representing the planned height)  *
    ' *         or a Double precision value (representing the        *
    ' *         intended multiple of the default zoom height that    *
    ' *         obtains the actual height.)                          *
    ' *                                                              *
    ' *      *  new(defaultSize): the object constructor must specify*
    ' *         the current and default size of the zoom form, or    *
    ' *         if available, that of the zoomed control: this       *
    ' *         size is used when Width or Height are multiples of   *
    ' *         the default sizes.                                   * 
    ' *                                                              *
    ' *      *  Size: this read-only property returns the resolved   *
    ' *         size as a Size object.  This will consist of the     *
    ' *         original, planned integer width and height, or, the  *
    ' *         double width or double height times the default size.)
    ' *                                                              *
    ' *      *  Width: this read-write property returns and may be   *
    ' *         set to the planned width of the zoom form, and it    *
    ' *         may be an integer (representing the planned width)   *
    ' *         or a Double precision value (representing the        *
    ' *         intended multiple of the default zoom width that     *
    ' *         obtains the actual width.)                           *
    ' *                                                              *
    ' *                                                              *
    ' ****************************************************************
    Private Class zoomSize
    
        ' ***** Shared *****
        Private Shared _OBJutilities As utilities 
        
        ' ***** Object state *****
        Private Structure TYPstate
            Dim intDefaultHeight As Integer
            Dim intDefaultWidth As Integer
            Dim objWidth As Object         ' Object or double
            Dim objHeight As Object        ' Object or double
        End Structure        
        Private USRstate As TYPstate
        
        ' ***** Public procedures *************************************
        
        ' -------------------------------------------------------------
        ' Return and change height
        '
        '
        Public Property Height As Object
            Get
                Return(USRstate.objHeight)
            End Get            
            Set(ByVal objNewValue As Object)
                If Not (TypeOf objNewValue Is Integer) _
                   AndAlso _
                   Not (TypeOf objNewValue Is Double) Then
                    errorHandler_("Height has invalid type", _
                                  "Height set", _
                                  "Object instance not changed")
                    Return                                  
                End If    
                USRstate.objHeight = objNewValue               
            End Set            
        End Property        
        
        ' -------------------------------------------------------------
        ' Object constructor
        '
        '
        Public Sub new(ByVal objDefaultSize As Size)
            With USRstate
                .intDefaultWidth = objDefaultSize.Width
                .intDefaultHeight = objDefaultSize.Height
                .objWidth = CDbl(1): .objHeight = CDbl(1)
            End With            
        End Sub        
        
        ' -------------------------------------------------------------
        ' Resolve and return the size
        '
        '
        Public ReadOnly Property Size As Size
            Get
                Dim objSize As New Size(resolveWidth_, resolveHeight_)
                Return(objSize)
            End Get            
        End Property        
        
        ' -------------------------------------------------------------
        ' Return and change width
        '
        '
        Public Property Width As Object
            Get
                Return(USRstate.objWidth)
            End Get            
            Set(ByVal objNewValue As Object)
                If Not (TypeOf objNewValue Is Integer) _
                   AndAlso _
                   Not (TypeOf objNewValue Is Double) Then
                    errorHandler_("Width has invalid type", _
                                  "Width set", _
                                  "Object instance not changed")
                    Return                                  
                End If    
                USRstate.objWidth = objNewValue               
            End Set            
        End Property        
        
        ' ***** Private procedures ***********************************
        
        ' ------------------------------------------------------------
        ' Interface to error handler
        '
        '
        Private Sub errorHandler_(ByVal strMessage As String, _
                                  ByVal strProcedure As String, _
                                  ByVal strHelp As String)
            _OBJutilities.errorHandler(strMessage, _
                                       "zoomSize", strProcedure, _
                                       strHelp)
        End Sub                                  
        
        ' -------------------------------------------------------------
        ' Resolve height to an integer value
        '
        '
        Private Function resolveHeight_ As Integer
            With USRstate
                If (TypeOf .objHeight Is Integer) Then 
                    Return(CInt(.objHeight))
                End If
                .objHeight = CInt(.intDefaultHeight * CDbl(.objHeight))
                Return CInt(.objHeight)
            End With                                
        End Function        
        
        ' -------------------------------------------------------------
        ' Resolve width to an integer value
        '
        '
        Private Function resolveWidth_ As Integer
            With USRstate
                If (TypeOf .objWidth Is Integer) Then 
                    Return(CInt(.objWidth))
                End If
                .objWidth = CInt(.intDefaultWidth * CDbl(.objWidth))
                Return CInt(.objWidth)
            End With                                
        End Function        
        
    End Class    
End Class
