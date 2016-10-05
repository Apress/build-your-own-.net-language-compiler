Option Strict

Imports System
Imports Microsoft.VisualBasic
Imports System.Windows.Forms
Imports System.Threading

' **************************************************************************
' *                                                                        *
' * A collection of utility methods (for Windows applications)             *
' *                                                                        *
' *                                                                        *
' * This stateless object exposes a number of Shared utility methods as    *
' * well as Shared constant-valued properties including GRID (the standard *
' * grid size for Windows forms, at least as far as my applications are    *
' * concerned.)                                                            *
' *                                                                        *
' * The rest of this header block addresses the following topics.          *
' *                                                                        *
' *                                                                        *
' *      *  Methods and properties exposed by this object                  *
' *      *  Multithreading considerations                                  *
' *      *  Change record                                                  *
' *      *  Issues                                                         *
' *                                                                        *
' *                                                                        *
' * METHODS AND PROPERTIES  EXPOSED BY THIS OBJECT ----------------------- *
' *                                                                        *
' *                                                                        *
' *      *  addUniqueListItem: add list item (make sure it's not already   *
' *         present)                                                       *
' *                                                                        *
' *      *  clearRegistryKey: clear the Registry keys specified            *
' *                                                                        *
' *      *  controlBottom: return the base of the control                  *
' *                                                                        *
' *      *  controlChange: modify the Enabled and/or Visible status of a   *
' *         set of controls                                                *
' *                                                                        *
' *      *  controlRight: return the right side of the control             *
' *                                                                        *
' *      *  controlSemiclone: clone or copy  the useful properties of the  *
' *         more common controls                                           *
' *                                                                        *
' *      *  disposeControl: dispose control and its constituents           *
' *                                                                        *
' *      *  extendTextBox: extend the Text box                             *
' *                                                                        *
' *      *  Grid: return a standardized form grid value                    *
' *                                                                        *
' *      *  listBox2Registry: save the list box in the Registry (Windows)  *
' *                                                                        *
' *      *  listBox2String: convert the list box's item list to a string   *
' *                                                                        *
' *      *  MaxTextBoxLength: return maximum text box length as supported  *
' *         by extendTextBox                                               *
' *                                                                        *
' *      *  registry2Listbox: restore list box from the Registry (Windows) *
' *                                                                        *
' *      *  resizeControls: resize the constituent controls of a form      *
' *         or other container                                             *
' *                                                                        *
' *      *  screenHeight: returns the height of the primary screen         *
' *                                                                        *
' *      *  screenSize: returns the width and height of the primary screen *
' *                                                                        *
' *      *  screenWidth: returns the width of the primary screen           *
' *                                                                        *
' *      *  searchListBox: searches list box for a string value (Windows)  *
' *                                                                        *
' *      *  setControl2GroupBox: initializes a control to a GroupBox       *
' *                                                                        *
' *      *  sortUsingListBox: sorts lines in a string, using a temp list   *
' *         box                                                            *
' *                                                                        *
' *      *  string2ListBox: transfers a string to a list box               *
' *                                                                        *
' *      *  test: test the procedures in this object                       *
' *                                                                        *
' *      *  trimContainer: adjust form/container width and height (Windows)*
' *                                                                        *
' *                                                                        *
' * MULTITHREADING CONSIDERATIONS ---------------------------------------- *
' *                                                                        *
' * This object is stateless, containing no Shared or class-level variables*
' * other than constants and an unallocated Shared utilities object.       *
' * All of its methods are shared (it contains no properties, and no non-  *
' * shared methods or properties should be added.                          *
' *                                                                        *
' * Therefore this object may be used in multiple threading applications   *
' * including multiple copies running simultaneously and a single copy,    *
' * with multiple methods executing simultaneously.                        *
' *                                                                        *
' *                                                                        *
' * C H A N G E   R E C O R D -------------------------------------------- *
' *   DATE     PROGRAMMER     DESCRIPTION OF CHANGE                        *
' * --------   ----------     -------------------------------------------- *
' * 01 07 03   Nilges         1.  Version 1.0 spun off from utilities      *
' * 01 11 03   Nilges         1.  Changed controlRight/trimContainer       *
' * 02 07 03   Nilges         1.  Added controlSemiclone                   *
' *                           2.  Added disposeRecursive                   *
' * 02 09 03   Nilges         1.  Added setControlToGroupBox               *
' * 02 11 03   Nilges         1.  Added zoom                               *
' *                           2.  Added listbox2String                     *
' * 05 06 03   Nilges         1.  Converted to Visual Studio Pro 2003      *
' * 05 09 03   Nilges         1.  Removed references to zoom, which is an  *
' *                               independent class                        *  
' * 05 12 03   Nilges         1.  Added sortUsingListBox                   *  
' *                           2.  Added string2ListBox                     *
' * 05 18 03   Nilges         1.  Changed registry2Listbox                 *
' *                           2.  Changed searchListbox                    *
' * 06 01 03   Nilges         1.  Added unimplemented test method          *
' * 06 09 03   Nilges         1.  disposeRecursive has been changed to     *
' *                               disposeControl                           *
' *                           2.  Assign Nothing to _OBJutilities to avoid *
' *                               subsequent, temporary, failures in the   *
' *                               compiler, where _OBJutilities hasn't been*
' *                               unloaded yet by the Framework.           *
' * 07 03 03   Nilges         1.  Added addUniqueListItem                  *
' * 07 03 03   Nilges         1.  Added clearRegistryKey                   *
' * 12 31 03   Nilges         1.  Added controlChange                      *
' * 01 05 04   Nilges         1.  Added updateStatusListBox                *
' * 01 19 04   Nilges         1.  Added issues and solution readme         *
' * 01 20 04   Nilges         1.  Changed updateStatusListBox: added       *
' *                               booSplitLines                            *
' * 01 24 04   Nilges         1.  Changed updateStatusListBox: added       *
' *                               booIncludeDate                           *
' *                                                                        *
' * I S S U E S ---------------------------------------------------------- *
' *   DATE     PROGRAMMER     DESCRIPTION                                  *
' * --------   ----------     -------------------------------------------- *
' * 01 19 04   Nilges         bytLineWidth parameter of                    *
' *                           updateStatusListBox does not work properly   *
' **************************************************************************

Public Class windowsUtilities

    Private Shared _OBJutilities As utilities.utilities 
    
    Private Const GRID_ As Integer = 8

    ' ---------------------------------------------------------------------
    ' Return a standardized form grid value
    '
    '
    Public Shared ReadOnly Property Grid As Integer
        Get
            Return(GRID_)
        End Get        
    End Property    

    ' ---------------------------------------------------------------------
    ' Extend list box (unless entry is already present)
    '
    '
    ' This method adds a string item to lstBox unless the item is already
    ' present. It returns False when the item is already present or True
    ' when the item is new.
    '
    ' By default the list box is searched for a precise match.  However, the
    ' optional parameter booCaseSensitive:=False parameter may be passed to
    ' search the list ignoring case, and the optional parameter booTrim:=True
    ' may be passed to search the list while removing leading and ending spaces
    ' from list entries and the new item.
    '
    '
    Public Shared Function addUniqueListItem(ByRef lstBox As ListBox, _
                                             ByVal strNewItem As String, _
                                             Optional ByVal booCaseSensitive _
                                                      As Boolean = True, _
                                             Optional ByVal booTrim _
                                                      As Boolean = False) _
           As Boolean
        Dim intIndex1 As Integer = searchListBox(lstBox, _
                                                 strNewItem, _
                                                 booCaseSensitive:=booCaseSensitive, _
                                                 booTrim:=booTrim)
        If intIndex1 > -1 Then Return(False)
        Try
            lstBox.Items.Add(strNewItem)
        Catch  
            _OBJutilities.errorHandler("Cannot add item to list box: " & _
                                       Err.Number & " " & Err.Description, _
                                       "windowsUtilities", "addUniqueListItem", _
                                       "Returning False")
            Return(False)                                       
        End Try          
        Return(True)                                               
    End Function           
    
    ' -----------------------------------------------------------------
    ' Clear Registry values with Windows-style MsgBox prompting
    '
    '
    ' This method clears (within HKEY_CURRENT_USER\Software\VB and VBA
    ' program settings) the key identified in strMainKey. If the overload
    ' clearRegistryKey(strMainKey, strSubkey) is called, then only the
    ' key strMainKey\strSubkey is cleared.
    '
    ' By default, this method will use a Windows message box, to prompt
    ' the user for approval to clear the keys. The message will be, by
    ' default, "Do you want to clear your form selections?".
    '
    ' Use the optional overload syntax clearRegistryKey(main,sub,prompt)
    ' to override the default prompt. If a prompt is specified that is
    ' null or blank, no prompt is made.
    '
    '
    ' --- Clears all subkeys
    Public Shared Overloads Function clearRegistryKey(ByVal strMainKey As String) _
           As Boolean
        Return(clearRegistryKey(strMainKey, ""))
    End Function    
    ' --- Clears one subkeys
    Public Shared Overloads Function clearRegistryKey(ByVal strMainKey As String, _
                                                      ByVal strSubKey As String) _
           As Boolean
        Return(clearRegistryKey(strMainKey, _
                                strSubKey, _
                                "Do you want to clear your form selections?"))
    End Function           
    ' --- Clears one subkey
    Public Shared Overloads Function clearRegistryKey(ByVal strMainKey As String, _
                                                      ByVal strSubKey As String, _
                                                      ByVal strPrompt As String) _
           As Boolean
        If Trim(strMainKey) = "" Then
            _OBJutilities.errorHandler("Main key cannot be blank or null", _
                                       "windowsUtilities", "clearRegistry", _
                                       "Returning false")
        End If                   
        If Trim(strPrompt) <> "" Then
            Select Case MsgBox(strPrompt, MsgBoxStyle.YesNo)
                Case MsgBoxResult.Yes:
                Case MsgBoxResult.No: Return(False)
                Case Else:
                    _OBJutilities.errorHandler("Unexpected case", _
                                               "windowsUtilities", _
                                               "clearRegistry", _
                                               "Won't clear form selections")
                    Return(False)                             
            End Select 
        End If                                      
        Try
            If strSubKey = "" Then
                DeleteSetting(strMainKey)
            Else
                DeleteSetting(strMainKey, strSubKey)
            End If                
        Catch            
            _OBJutilities.errorHandler("Cannot delete form settings: " & _
                                        Err.Number & " " & Err.Description, _
                                        "windowsUtilities", _
                                        "clearRegistry", _
                                        "Won't clear form selections")
            Return(False)                             
        End Try        
        Return(True)
    End Function

    ' ---------------------------------------------------------------------
    ' Return base of control
    '
    '
    Public Shared Function controlBottom(ByVal ctlControl As Control) As Integer
        With ctlControl
            Return (.Top + .Height)
        End With
    End Function

    ' ---------------------------------------------------------------------
    ' Modify the status of a set of controls
    '
    '
    ' This method changes the Visible and/or Enabled status of one or more
    ' controls on a Windows form. It is passed an "example" control
    ' (ctlExample) and a new Visible and a new Enabled status, it 
    ' modifies a matching set of controls to match the new statuses, and,
    ' it creates a collection that identifies the old status and each
    ' control which was modified...which makes it easy to restore controls.
    '
    ' This method will modify each control belonging to the parent of the example
    ' that has its type and the identical Visible and Enabled status. It can
    ' modify controls inside container controls that belong to the parent of the
    ' example when a recursion flag is used.
    '
    ' Here is the synopsis of this method:
    '
    '
    '      Public Function controlChange(ByVal ctlExample As Control, _
    '                                    ByVal booNewVisible As Boolean, _
    '                                    ByVal booNewEnabled As Boolean, _
    '                                    ByVal booRecursion As Boolean) _
    '             As Collection
    '
    '
    ' In the above:
    '
    '
    '      *  ctlExample is the control which represents the set
    '
    '      *  booNewVisible is the desired new visibility setting
    '
    '      *  booNewEnabled is the desired new enabled setting
    '
    '      *  booRecursion should be True or False:
    '
    '         + If booRecursion is True, then controls in containers owned by
    '           the parent of ctlExample are changed at any level they occur.
    '
    '         + If booRecursion is False, then controls in containers owned by
    '           the parent of ctlExample are not changed. Only controls owned
    '           directly by the parent are modified.
    '
    '
    ' This method returns a collection with the following structure.
    '
    '
    '      *  Item(1) will be the old Visible status of each control 
    '         modified.
    '
    '      *  Item(2) will be the old Enabled status of each control 
    '         modified.
    '
    '      *  Item(3)..Item(n) will be the handle to each modified control.
    '
    '
    ' You can use this collection to safely restore all modified controls to
    ' their original status.
    '
    ' Note that by default, only controls belonging directly to the parent
    ' form or parent container are modified. Pass the optional parameter
    ' booRecursion:=True to modify all controls that belong directly or to
    ' a container control belonging to the parent at any level.
    '
    ' For example, supposed we want to set all buttons on a form to Visible 
    ' but not Enabled (including buttons inside containers on the form, such 
    ' as group boxes). After a while we want to safely restore all buttons.
    '
    ' Use this code, where cmdClose is a Close button on the form.
    '
    '
    '      ' -- Code to modify the buttons
    '      Dim colRestore As Collection = controlChange(cmdClose, _
    '                                                   True, _
    '                                                   False, _
    '                                                   True)
    '      .
    '      .
    '      .
    '      ' -- Code to restore the buttons
    '      With colRestore
    '          Dim intIndex1 As Integer
    '          Dim booOldVisibility As Boolean = CType(.Item(1), Boolean)
    '          Dim booOldEnabled As Boolean = CType(.Item(2), Boolean)
    '          For intIndex1 = 3 To .Count
    '              With CType(.Item(intIndex1), Control)
    '                  .Visible = booOldVisibility
    '                  .Enabled = booOldEnabled
    '              End With
    '          Next intIndex1
    '      End With
    '
    '
    Public Shared Function controlChange(ByVal ctlExample As Control, _
                                         ByVal booNewVisible As Boolean, _
                                         ByVal booNewEnabled As Boolean, _
                                         ByVal booRecursion As Boolean) _
                           As Collection
        Dim colRestore As Collection
        Try
            colRestore = New Collection
            With colRestore
                .Add(ctlExample.Visible)
                .Add(ctlExample.Enabled)
            End With
        Catch
            _OBJutilities.errorHandler("Cannot create restore collection", _
                                       "windowsUtilities", _
                                       "controlChange", _
                                       "Returning Nothing")
            Return (Nothing)
        End Try
        controlChange_(ctlExample.Parent, _
                        ctlExample, _
                        ctlExample.Visible, _
                        ctlExample.Enabled, _
                        booNewVisible, _
                        booNewEnabled, _
                        booRecursion, _
                        colRestore)
        Return (colRestore)
    End Function

    ' -----------------------------------------------------------------
    ' Recursion on behalf of controlChange
    '
    '
    Private Shared Sub controlChange_(ByVal ctlParent As Control, _
                                      ByVal ctlExample As Control, _
                                      ByVal booOldVisible As Boolean, _
                                      ByVal booOldEnabled As Boolean, _
                                      ByVal booNewVisible As Boolean, _
                                      ByVal booNewEnabled As Boolean, _
                                      ByVal booRecursion As Boolean, _
                                      ByRef colRestore As Collection)
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim strType As String = UCase(ctlExample.GetType.ToString)
        With ctlParent
            For intIndex1 = 0 To .Controls.Count - 1
                With .Controls(intIndex1)
                    If UCase(.GetType.ToString) = strType _
                       AndAlso _
                       booOldVisible = booOldVisible _
                       AndAlso _
                       booOldEnabled = booOldEnabled Then
                        .Visible = booNewVisible
                        .Enabled = booNewEnabled
                        .Refresh()
                        Try
                            colRestore.Add(ctlParent.Controls(intIndex1))
                        Catch ex As Exception
                            _OBJutilities.errorHandler("Cannot expand restore collection: " & _
                                                       Err.Number & " " & Err.Description, _
                                                        "windowsUtilities", _
                                                        "controlChange", _
                                                        "Change is incomplete")
                            Return
                        End Try
                    End If
                    If booRecursion Then
                        For intIndex2 = 0 To .Controls.Count - 1
                            controlChange_(ctlParent.Controls.Item(intIndex1), _
                                           ctlExample, _
                                           booOldVisible, _
                                           booOldEnabled, _
                                           booNewVisible, _
                                           booNewEnabled, _
                                           True, _
                                           colRestore)
                        Next intIndex2
                    End If
                End With
            Next intIndex1
        End With
    End Sub


    ' ---------------------------------------------------------------------
    ' Return right side of control
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   --------------------------------------------
    ' 01 11 03     Nilges       Method should be Shared
    '
    '
    Public Shared Function controlRight(ByVal ctlControl As Control) As Integer
        With ctlControl
            Return (.Left + .Width)
        End With
    End Function

    ' ---------------------------------------------------------------------
    ' Clone the useful properties of some controls
    '
    '
    ' This method clones command buttons, labels, list boxes, text boxes,
    ' and group boxes (that contain, exclusively, command buttons, labels,
    ' list boxes, text boxes and recursive group boxes.)  It is able to
    ' copy properties from one of these controls to another of identical type
    ' without "cloning", which involves the creation of a new instance.
    '
    ' In copying and cloning, only the common properties of each control 
    ' are copied and cloned and the copying and cloning is not guaranteed to
    ' be precise for this reason.
    '
    ' This method has two overloads:
    '
    '
    '      *  controlSemiclone(ctl) creates a clone and returns it, or Nothing
    '         on failure
    '
    '      *  controlSemiclone(ctl1, ctl2) copies the common properties of
    '         ctl1 to ctl2, returning True on success and False on failure
    '
    '
    ' --- Clone control 
    Public Overloads Function controlSemiclone(ByVal ctlHandle As Control) As Control
        Dim ctlHandle2 As Control
        Dim ctlHandleConstituent As Control
        Dim intIndex1 As Integer
        If (TypeOf ctlHandle Is Button) Then
            ctlHandle2 = New Button
        ElseIf (TypeOf ctlHandle Is GroupBox) Then
            ctlHandle2 = New GroupBox
            With CType(ctlHandle, GroupBox)
                For intIndex1 = 0 To .Controls.Count
                    Try
                        ctlHandleConstituent = controlSemiclone(.Controls.Item(intIndex1))
                    Catch : End Try
                    If (ctlHandleConstituent Is Nothing) Then
                        _OBJutilities.errorHandler("Group box has constituent controls " & _
                                                   "which cannot be cloned", _
                                                   "windowsUtilities", _
                                                   "controlSemiclone", _
                                                   "Returning Nothing")
                        disposeControl(ctlHandle2) : ctlHandle2 = Nothing
                        Return (Nothing)
                    End If
                    Try
                        ctlHandle2.Controls.Add(ctlHandleConstituent)
                    Catch
                        _OBJutilities.errorHandler("Cannot add constituent control " & _
                                                   _OBJutilities.object2String(ctlHandleConstituent) & " " & _
                                                   "to cloned group box", _
                                                   "", _
                                                   "controlSemiclone", _
                                                   "Returning Nothing")
                        disposeControl(ctlHandle2) : ctlHandle2 = Nothing
                        Return (Nothing)
                    End Try
                Next intIndex1
            End With
        ElseIf (TypeOf ctlHandle Is Label) Then
            ctlHandle2 = New Label
        ElseIf (TypeOf ctlHandle Is ListBox) Then
            ctlHandle2 = New ListBox
        ElseIf (TypeOf ctlHandle Is TextBox) Then
            ctlHandle2 = New TextBox
        Else
            _OBJutilities.errorHandler("Control " & _
                                       _OBJutilities.object2String(ctlHandle) & " " & _
                                       "has unsupported type", _
                                       "windowsUtilities", _
                                       "controlSemiclone", _
                                       "Returning Nothing")
            Return (Nothing)
        End If
        If Not controlSemiclone(ctlHandle, ctlHandle2) Then
            ctlHandle2.dispose() : ctlHandle2 = Nothing
            Return (Nothing)
        End If
        Return (ctlHandle2)
    End Function
    ' --- Copy control without cloning
    Public Overloads Function controlSemiclone(ByVal ctlHandle1 As Control, _
                                               ByVal ctlHandle2 As Control) As Boolean
        Dim intIndex1 As Integer
        Try
            With ctlHandle1
                .BackColor = ctlHandle1.BackColor
                .Enabled = ctlHandle1.Enabled
                .ForeColor = ctlHandle1.ForeColor
                .Location = ctlHandle1.Location
                .Size = ctlHandle1.Size
                .Text = ctlHandle1.Text
                .Visible = ctlHandle1.Visible
            End With
            If (TypeOf ctlHandle1 Is Label) Then
                CType(ctlHandle2, Label).BorderStyle = CType(ctlHandle1, Label).BorderStyle
                CType(ctlHandle2, Label).TextAlign = CType(ctlHandle1, Label).TextAlign
            ElseIf (TypeOf ctlHandle1 Is ListBox) Then
                With CType(ctlHandle2, ListBox)
                    Dim lstHandle1 As ListBox = CType(ctlHandle1, ListBox)
                    .Items.Clear()
                    For intIndex1 = 0 To lstHandle1.Items.Count - 1
                        .Items.Add(lstHandle1.Items(intIndex1))
                    Next intIndex1
                End With
            ElseIf (TypeOf ctlHandle1 Is TextBox) Then
                CType(ctlHandle2, TextBox).BorderStyle = CType(ctlHandle1, TextBox).BorderStyle
                CType(ctlHandle2, TextBox).ReadOnly = CType(ctlHandle1, TextBox).ReadOnly
                CType(ctlHandle2, TextBox).TextAlign = CType(ctlHandle1, TextBox).TextAlign
                ctlHandle2 = New TextBox
            Else
                _OBJutilities.errorHandler("Control " & _
                                           _OBJutilities.object2String(ctlHandle1) & " " & _
                                           "has unsupported type", _
                                           "windowsUtilities", _
                                           "controlSemiclone", _
                                           "Returning False")
                Return (False)
            End If
        Catch
            _OBJutilities.errorHandler("Cannot clone text box: " & Err.Number & " " & Err.Description, _
                                       "cloneTextBox_", _
                                       "Returning Nothing")
            Return (Nothing)
        End Try
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Dispose of control and any constituents as well
    '
    '
    Public Function disposeControl(ByVal ctlHandle As Control) As Boolean
        Dim intIndex1 As Integer
        With ctlHandle.Controls
            For intIndex1 = 0 To ctlHandle.Controls.Count - 1
                If Not disposeControl(.Item(intIndex1)) Then
                    _OBJutilities.errorHandler("Cannot recursively dispose of constituent control " & _
                                               _OBJutilities.object2String(.Item(intIndex1)) & ": " & _
                                               Err.Number & " " & Err.Description, _
                                               "windowsUtilities", "disposeControl", _
                                               "Returning False and not disposing any more " & _
                                               "constituents")
                    Return (False)
                End If
            Next intIndex1
            Try
                ctlHandle.dispose()
            Catch
                _OBJutilities.errorHandler("Cannot dispose of control " & _
                                            _OBJutilities.object2String(ctlHandle) & ": " & _
                                            Err.Number & " " & Err.Description, _
                                            "windowsUtilities", "disposeControl", _
                                            "Returning False")
                Return (False)
            End Try
            Return (True)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Add prompts and other material to the display
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   --------------------------------------------
    ' 12 01 01     Nilges       1.  Limit size 
    ' 03 17 02     Nilges       1.  Don't refresh by default: runs too slow 
    '
    '
    Public Overloads Shared Sub extendTextBox(ByRef txtBox As TextBox, ByVal strNew As String)
        extendTextBox(txtBox, strNew, False)
    End Sub
    Public Overloads Shared Sub extendTextBox(ByRef txtBox As TextBox, _
                                              ByVal strNew As String, _
                                              ByVal booRefresh As Boolean)
        With txtBox
            Dim booEnabled As Boolean = .Enabled
            .Enabled = False
            Dim intExcess As Integer = Len(.Text) + Len(strNew) - MaxTextBoxLength
            If intExcess > 0 Then
                .Text = Mid(.Text, intExcess + 1) & strNew
            Else
                .AppendText(strNew)
            End If
            .Enabled = booEnabled
            If booRefresh Then .Refresh()
        End With
    End Sub

    ' ----------------------------------------------------------------------
    ' Save list box or combo box (entries, index, and when list is combo, 
    ' text) in the Registry
    '
    '
    ' This method is able to store a list box (with a SelectionMode of
    ' either One or None only) or a Combo box in a set of Registry keys.
    ' It is useful for saving the contents of list and combo boxes which
    ' can vary dynamically.
    '
    ' This method creates the following keys in the Registry folder
    ' HKEY_CURRENT_USER\Software\VB and VBA Program Settings\application\
    ' section, where <name> is the name of the list-type box.
    '
    '
    '      <name>_TYPE: contains the type of the control:
    '
    '           *  LISTBOX_NOMULTI: a ListBox with a SelectionMode of either
    '              One or None
    '
    '           *  LISTBOX_MULTI: a ListBox with a SelectionMode of either
    '              MultiExtended or MultiSimple
    '
    '           *  COMBOBOX: a ComboBox
    '
    '      <name>_COUNT: contains the number of entries stored from
    '           the list or combo box
    '
    '      <name>_INDEX: contains the index of the currently selected
    '           entry in the list box.  This key is not stored for a 
    '           list box with a SelectionMode of None.
    '
    '      <name>_ITEM_<n>: contains the value of the nth entry in the list
    '           or combo box
    '
    '      <name>_COMBOTEXT: contains the text of a combo box, not stored for
    '           a list box
    '
    '
    ' The parameters strApplicationName and strSectionName are optional.  The application
    ' name defaults to the ProductName of the Application object unless the ProductName
    ' is null; when the ProductName is a null string, application detaults to LISTBOX2REGISTRY.
    '
    ' The section name defaults to the Name of the list box or combo box control prefixed
    ' by all of its available parent's names separated by underscores.  Usually this will be
    ' form_control.
    '
    '
    ' --- Store combo box
    Public Overloads Shared Function listBox2Registry(ByRef cboListBox As ComboBox, _
                                                      Optional ByVal strApplication As String = "", _
                                                      Optional ByVal strSection As String = "") As Boolean
        Return (listBox2Registry_(False, _
                                 cboListBox, _
                                 strApplication, _
                                 strSection, _
                                 cboListBox.Items.Count, _
                                 cboListBox.SelectedIndex, _
                                 cboListBox.Text))
    End Function
    ' --- Store list box
    Public Overloads Shared Function listBox2Registry(ByRef lstListBox As ListBox, _
                                                      Optional ByVal strApplication As String = "", _
                                                      Optional ByVal strSection As String = "") As Boolean
        Dim objUniqueIndex As Object
        Select Case lstListBox.SelectionMode
            Case SelectionMode.None
            Case SelectionMode.One : objUniqueIndex = lstListBox.SelectedIndex
            Case Else
                _OBJutilities.errorHandler("Cannot store this type of list box in the Registry", _
                                          "listBox2Registry")
                Return (False)
        End Select
        Return (listBox2Registry_(True, _
                                 lstListBox, _
                                 strApplication, _
                                 strSection, _
                                 lstListBox.Items.Count, _
                                 objUniqueIndex, _
                                 Nothing))
    End Function
    ' --- Common logic to update the Registry
    Private Shared Function listBox2Registry_(ByVal booListbox As Boolean, _
                                                ByVal ctlListBox As ListControl, _
                                                ByVal strApplicationName As String, _
                                                ByVal strSectionName As String, _
                                                ByVal intItems As Integer, _
                                                ByVal objUniqueIndex As Object, _
                                                ByVal objComboText As Object) As Boolean
        Dim strApplicationNameWork As String = _
            CStr(IIf(strApplicationName = "", Application.ProductName, strApplicationName))
        If strApplicationNameWork = "" Then
            strApplicationNameWork = "LISTBOX2REGISTRY"
        End If
        Dim strSectionNameWork As String = _
            CStr(IIf(strSectionName = "", listBox2Registry__defaultName_(ctlListBox), strSectionName))
        Dim strType As String = _
            CStr(IIf(booListbox, _
                     CStr(IIf(objUniqueIndex Is Nothing, "LISTBOX_MULTI", "LISTBOX_NOMULTI")), _
                     "LISTBOX_COMBO"))
        Dim strPrefix As String = ctlListBox.Name & "_"
        Dim strCountKeyName As String = strPrefix & "COUNT"
        Dim strIndexKeyName As String = strPrefix & "INDEX"
        Dim strTypeKeyName As String = strPrefix & "TYPE"
        Dim intIndex1 As Integer
        Try
            DeleteSetting(strApplicationNameWork, strSectionNameWork, strIndexKeyName)
            DeleteSetting(strApplicationNameWork, strSectionNameWork, strTypeKeyName)
            Dim strPreviousCount As String = GetSetting(strApplicationNameWork, _
                                                        strSectionNameWork, _
                                                        strCountKeyName)
            Dim intPreviousCount As Integer
            Try
                intPreviousCount = CInt(strPreviousCount)
            Catch : End Try
            For intIndex1 = 0 To intPreviousCount - 1
                Try
                    DeleteSetting(strApplicationNameWork, strSectionNameWork, strPrefix & "ITEM_" & intIndex1)
                Catch : End Try
            Next intIndex1
            DeleteSetting(strApplicationNameWork, strSectionNameWork, strCountKeyName)
        Catch : End Try
        Try
            SaveSetting(strApplicationNameWork, _
                        strSectionNameWork, _
                        strTypeKeyName, _
                        strType)
            SaveSetting(strApplicationNameWork, _
                        strSectionNameWork, _
                        strCountKeyName, _
                        CStr(intItems))
            If Not (objUniqueIndex Is Nothing) Then
                SaveSetting(strApplicationNameWork, _
                            strSectionNameWork, _
                            strIndexKeyName, _
                            CStr(objUniqueIndex))
            End If
        Catch
            Return (listBox2Registry__errorHandler_(strApplicationNameWork, strSectionNameWork))
        End Try
        Dim cboListBoxHandle As ComboBox
        Dim lstListBoxHandle As ListBox
        Dim strNext As String
        If booListbox Then
            lstListBoxHandle = CType(ctlListBox, ListBox)
        Else
            cboListBoxHandle = CType(ctlListBox, ComboBox)
        End If
        For intIndex1 = 0 To intItems - 1
            If booListbox Then
                strNext = CStr(lstListBoxHandle.Items(intIndex1))
            Else
                strNext = CStr(cboListBoxHandle.Items(intIndex1))
            End If
            Try
                SaveSetting(strApplicationNameWork, strSectionNameWork, strPrefix & "ITEM_" & CStr(intIndex1), strNext)
            Catch
                Return (listBox2Registry__errorHandler_(strApplicationNameWork, strSectionNameWork))
            End Try
        Next intIndex1
        If Not booListbox Then
            Try
                SaveSetting(strApplicationNameWork, strSectionNameWork, strPrefix & "COMBOTEXT", cboListBoxHandle.Text)
            Catch
                Return (listBox2Registry__errorHandler_(strApplicationNameWork, strSectionNameWork))
            End Try
        End If
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' On behalf of listbox2Registry, return default section name
    '
    '
    Private Shared Function listBox2Registry__defaultName_(ByVal ctlListbox As ListControl) As String
        Dim ctlControl As Control = ctlListbox
        Dim objStringBuilder As New System.Text.StringBuilder
        Dim strNext As String = ctlListbox.Name
        Do
            _OBJutilities.append(objStringBuilder, "_", strNext, booToStart:=True)
            ctlControl = ctlControl.Parent
            If (ctlControl Is Nothing) Then Exit Do
            strNext = ctlControl.Name
        Loop Until strNext = ""
        Return (objStringBuilder.ToString)
    End Function

    ' ----------------------------------------------------------------------
    ' On behalf of listBox2Registry, display error and cleanup
    '
    '
    ' This method removes keys created by a failed attempt to store a list
    ' or combo box in the Registry and always returns False.
    '
    '
    Private Shared Function listBox2Registry__errorHandler_(ByVal strApplicationName As String, _
                                                            ByVal strSectionNameWork As String) As Boolean
        _OBJutilities.errorHandler("Unable to store values of list or combo box", _
                                   "listBox2Registry")
        Try
            DeleteSetting(strApplicationName, strSectionNameWork)
        Catch : End Try
        Return (False)
    End Function

    ' ---------------------------------------------------------------------
    ' Convert a list box to a string
    '
    '
    ' This method forms a string from the item list in a list box.  Each
    ' item is separated by newline, or, the value of the optional
    ' strSep parameter.
    '
    '
    Public Shared Function listBox2String(ByVal lstBox As ListBox, _
                                          Optional ByVal strSep As String = vbNewLine) _
           As String
        Dim objStringBuilder As System.Text.StringBuilder
        Try
            objStringBuilder = New System.Text.StringBuilder
        Catch
            _OBJutilities.errorHandler("Cannot create string builder: " & _
                                       Err.Number & " " & Err.Description, _
                                       "windowsUtilities", "listBox2String", _
                                       "Returning a null string")
            Return ("")
        End Try
        With lstBox
            Dim intIndex1 As Integer
            For intIndex1 = 0 To .Items.Count - 1
                Try
                    _OBJutilities.append(objStringBuilder, strSep, CStr(.Items(intIndex1)))
                Catch
                    _OBJutilities.errorHandler("Cannot append to string builder: " & _
                                               Err.Number & " " & Err.Description, _
                                               "windowsUtilities", "listBox2String", _
                                               "Returning an incomplete string")
                    Exit For
                End Try
            Next intIndex1
            Return (objStringBuilder.ToString)
        End With
    End Function

    ' ---------------------------------------------------------------------
    ' Return a standardized max length for the text box
    '
    '
    Public Shared ReadOnly Property MaxTextBoxLength() As Integer
        Get
            Return (30000)
        End Get
    End Property

    ' ----------------------------------------------------------------------
    ' Restore list box or combo box (entries, index, and when list is combo, 
    ' text) from the Registry
    '
    '
    ' This method expects the following keys in the Registry folder
    ' HKEY_CURRENT_USER\Software\VB and VBA Program Settings\application\
    ' section, where <name> is the name of the list-type box.
    '
    '
    '      <name>_TYPE: should contain the type of the control:
    '
    '           *  LISTBOX_NOMULTI: a ListBox with a SelectionMode of either
    '              One or None
    '
    '           *  LISTBOX_MULTI: a ListBox with a SelectionMode of either
    '              MultiExtended or MultiSimple
    '
    '           *  COMBOBOX: a ComboBox
    '
    '      <name>_COUNT: should contain the number of entries stored from
    '           the list or combo box
    '
    '      <name>_INDEX: should the index of the currently selected
    '           entry in the list box or of the combo box.  This key
    '           is not stored for a list box with a SelectionMode of
    '           None, MultiExtended, or MultiSimple.
    '
    '      <name>_ITEM_<n>: should the value of the nth entry in the list
    '           or combo box
    '
    '      <name>_ITEM_<n>_SELECTED: stored only when the TYPE is LISTBOX_MULTI,
    '           this key should contain the selection status (Y for selected: N for
    '           not selected) of each list box entry. 
    '
    '      <name>_COMBOTEXT: should contain the text of a combo box, not stored for
    '           a list box
    '
    '
    ' The parameters strApplicationName and strSectionName are optional.  The application
    ' name defaults to the ProductName of the Application object unless the ProductName
    ' is null; when the ProductName is a null string, application defaults to LISTBOX2REGISTRY.
    '
    ' The section name defaults to the Name of the list box or combo box control.  If this name
    ' is not available for some silly reason the section name defaults to LISTBOX2REGISTRY.
    '
    ' By default, the list box is cleared and filled with the Registry's values unless the optional
    ' booAppend parameter is present and True.
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   --------------------------------------------
    ' 05 18 03     Nilges       1.  Bug in list box handling 
    '                           2.  Supply default count
    '
    '
    ' --- Control type enumerator
    Private Enum ENUregistry2ListBoxCtlType
        nomulti
        combo
        notValid
    End Enum
    ' --- Store combo box
    Public Overloads Shared Function registry2ListBox(ByRef cboListBox As ComboBox, _
                                                        Optional ByVal strApplication As String = "", _
                                                        Optional ByVal strSection As String = "", _
                                                        Optional ByVal booAppend As Boolean = False) As Boolean
        If Not booAppend Then cboListBox.Items.Clear()
        Return (registry2ListBox_(False, _
                                 cboListBox, _
                                 strApplication, _
                                 strSection))
    End Function
    ' --- Store list box
    Public Overloads Shared Function registry2ListBox(ByRef lstListBox As ListBox, _
                                                      Optional ByVal strApplication As String = "", _
                                                      Optional ByVal strSection As String = "", _
                                                      Optional ByVal booAppend As Boolean = False) As Boolean
        If Not booAppend Then lstListBox.Items.Clear()
        Return (registry2ListBox_(True, lstListBox, strApplication, strSection))
    End Function
    ' --- Common logic to update the Registry
    Private Shared Function registry2ListBox_(ByVal booListbox As Boolean, _
                                              ByVal ctlListBox As ListControl, _
                                              ByVal strApplicationName As String, _
                                              ByVal strSectionName As String) As Boolean
        Dim strApplicationNameWork As String = _
            CStr(IIf(strApplicationName = "", Application.ProductName, strApplicationName))
        If strApplicationNameWork = "" Then
            strApplicationNameWork = "LISTBOX2REGISTRY"
        End If
        Dim strPrefix As String = ctlListBox.Name & "_"
        Dim strSectionNameWork As String = _
            listBox2Registry__defaultName_(ctlListBox)
        If strSectionNameWork = "" Then
            strSectionNameWork = ctlListBox.Parent.Name
            If strSectionNameWork = "" Then strSectionNameWork = "LISTBOX2REGISTRY"
        End If
        If strSectionName <> "" Then strSectionNameWork = strSectionName
        Dim enuType As ENUregistry2ListBoxCtlType = _
                       registry2Listbox_type2Enum_(GetSetting(strApplicationNameWork, _
                                                              strSectionNameWork, _
                                                              strPrefix & "TYPE"))
        Dim booOK As Boolean
        Dim cboListBoxHandle As ComboBox
        Dim lstListBoxHandle As ListBox
        Select Case enuType
            Case ENUregistry2ListBoxCtlType.combo
                cboListBoxHandle = CType(ctlListBox, ComboBox)
                booOK = Not booListbox
            Case ENUregistry2ListBoxCtlType.nomulti
                lstListBoxHandle = CType(ctlListBox, ListBox)
                booOK = booListbox
            Case Else
                Select Case UCase(ctlListBox.GetType.ToString)
                    Case "SYSTEM.WINDOWS.FORMS.LISTBOX"
                        lstListBoxHandle = CType(ctlListBox, ListBox)
                        enuType = CType(IIf(lstListBoxHandle.SelectionMode = SelectionMode.One, _
                                            ENUregistry2ListBoxCtlType.nomulti, _
                                            ENUregistry2ListBoxCtlType.notValid), _
                                        ENUregistry2ListBoxCtlType)
                        booOK = booListbox
                    Case "SYSTEM.WINDOWS.FORMS.COMBOBOX"
                        cboListBoxHandle = CType(ctlListBox, ComboBox)
                        enuType = ENUregistry2ListBoxCtlType.combo
                        booOK = Not booListbox
                    Case Else : Return (False)
                End Select
        End Select
        If Not booOK Then
            _OBJutilities.errorHandler("List/combo box type does not agree with the value in the Registry", _
                                        "registry2Listbox")
            Return (False)
        End If
        Dim intIndex1 As Integer
        Dim strIndex As String = GetSetting(strApplicationNameWork, _
                                            strSectionNameWork, _
                                            strPrefix & "INDEX")
        Dim intIndex As Integer
        Try
            intIndex = CInt(strIndex)
        Catch
            intIndex = -1                   ' Minor repair to Registry
        End Try
        If intIndex < -1 Then intIndex = -1
        Dim intItems As Integer
        Dim strNext As String
        Try
            intItems = CInt(GetSetting(strApplicationNameWork, _
                                       strSectionNameWork, _
                                       strPrefix & "COUNT", _
                                       "0"))
        Catch
            _OBJutilities.errorHandler("Unexpected Registry entries exist: no COUNT available", _
                                      "registry2Listbox")
            Return (False)
        End Try
        For intIndex1 = 0 To intItems - 1
            Try
                strNext = GetSetting(strApplicationNameWork, strSectionNameWork, strPrefix & "ITEM_" & CStr(intIndex1))
            Catch
                _OBJutilities.errorHandler("Unexpected Registry entries exist: cannot Add item " & intIndex1, _
                                            "registry2Listbox")
                Return (False)
            End Try
            Try
                If booListbox Then
                    lstListBoxHandle.Items.Add(strNext)
                Else
                    cboListBoxHandle.Items.Add(strNext)
                End If
            Catch : End Try
        Next intIndex1
        If booListbox Then
            lstListBoxHandle.SelectedIndex = intIndex
        Else
            cboListBoxHandle.SelectedIndex = intIndex
            Try
                cboListBoxHandle.Text = GetSetting(strApplicationNameWork, strSectionNameWork, strPrefix & "COMBOTEXT")
            Catch
                Return (False)
            End Try
        End If
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' On behalf of registry2Listbox, convert string type to an enumeration
    '
    '
    Private Shared Function registry2Listbox_type2Enum_(ByVal strType As String) As ENUregistry2ListBoxCtlType
        Select Case UCase(strType)
            Case "LISTBOX_NOMULTI" : Return ENUregistry2ListBoxCtlType.nomulti
            Case "LISTBOX_COMBO" : Return ENUregistry2ListBoxCtlType.combo
        End Select
        Return ENUregistry2ListBoxCtlType.notValid
    End Function

    ' ----------------------------------------------------------------------
    '
    ' Resize form or container controls proportionally
    '
    '
    ' This tool stretches or shrinks controls on a form or other container
    ' proportionally to a change in form height and form width.  
    '
    ' ctlHandle should identify the container control: note that all forms
    ' are Controls in .Net. For example, pass Me as the control handle in the 
    ' Resize event code of a form, to adjust the controls in your form or 
    ' container.
    '
    ' Pass the PREVIOUS height and the PREVIOUS width of the container in 
    ' intOldHeight and intOldWidth.  
    '
    ' All controls are resized proportionally whether they appear in the
    ' container or one if its contained controls.
    '
    ' To proportionally adjust the Font of each control that has a Font size,
    ' pass True in the optional booChangeFontSize parameter.  However, the font
    ' size is never reduced below 8.25.
    '
    ' The optional parameter sglMinFontSize can be passed to enforce a mininum
    ' font size even when font sizes are changed: the optional sglMaxFontSize
    ' parameter will enforce a maximum font size.  Use these parameters to avoid
    ' silly visual results, or pass a value of -1 in either sglMinFontSize or
    ' sglMaxFontSize to avoid any enforcement of mininum/maximum size.  The
    ' default mininum is 8.25 (points, where a point is 1/72") and the default
    ' maximum is 42 points.
    '
    ' The sglFontSizeTuner parameter can be used for fine adjustments to the
    ' font size.  After the font size is calculated using the height-width
    ' ration and the mininum and maximum values, it can be multiplied by this
    ' tuner value.  sglFontSizeTuner defaults to 1 (no tuning.)
    '
    ' Note that the font size parameters sglMinFontSize, sglMaxFontSize and
    ' sglFontSizeTuner are ignored when booChangeFontSize is False.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------
    ' 01 02 99   Nilges         1.  Added booChangeFontSize parameter
    '                           2.  Added documentation
    ' 02 25 99   Nilges         1.  Bug: font size change incorrect
    ' 07 26 99   Nilges         1.  Bug: division by zero
    ' 10 30 99   Nilges         1.  Added sglMinFontSize and
    '                               sglMaxFontSize parameters
    '                           2.  Added sglFontSizeTuner
    ' 11 23 99   Nilges         1.  Added varProgress and
    '                               varProgressBar
    ' 11 25 99   Nilges         1.  Added booProgressDoEvents
    ' 12 09 99   Nilges         1.  Added varShrinkWrapTolerance
    ' 02 22 99   Nilges         1.  Added form cache
    ' 07 14 00   Nilges         1.  Bug: tried to change r/o drive
    '                               Height property like an idiot
    ' 07 15  00  Nilges         1.  Added selected controls count to
    '                               progress report
    '                           2.  Allow read-only height error
    '                               but flag other height change
    '                               errors
    '                           3.  Added varAnimationColor and
    '                               varAnimationTimer
    ' 07 16 00   Nilges         1.  Added bytResizeSequence
    ' 02 27 04   Nilges         1.  Converted from VB-6 and simplified
    ' 02 28 04   Nilges         1.  Bug: negative dimensions
    '                               1.1 Use absolute value
    ' 03 17 04   Nilges         1.  Bug: font based on incorrect control
    '                               1.1 Fixed						      
    '
    '
    Public Shared Function resizeConstituentControls(ByRef ctlHandle As Control, _
                                                     ByVal intOldWidth As Double, _
                                                     ByVal intOldHeight As Double, _
                                                     Optional ByVal booChangeFontSize As Boolean = False, _
                                                     Optional ByVal sglMinFontSize As Single = 8.25, _
                                                     Optional ByVal sglMaxFontSize As Single = 42, _
                                                     Optional ByVal sglFontSizeTuner As Single = 1) _
           As Boolean
        With ctlHandle
            Dim dblHeightRatio As Double
            Dim dblWidthRatio As Double
            Dim intOldConstituentHeight As Integer
            Dim intOldConstituentWidth As Integer
            Dim sglEmsize As Single
            Dim intIndex1 As Integer
            dblWidthRatio = ctlHandle.Width / intOldWidth
            dblHeightRatio = ctlHandle.Height / intOldHeight
            For intIndex1 = 0 To .Controls.Count - 1
                With .Controls(intIndex1)
                    intOldConstituentWidth = .Width
                    intOldConstituentHeight = .Height
                    .Width = CInt(.Width * dblWidthRatio)
                    .Height = CInt(.Height * dblHeightRatio)
                    .Left = CInt(.Left * dblWidthRatio)
                    .Top = CInt(.Top * dblHeightRatio)
                    If Not resizeConstituentControls(ctlHandle.Controls(intIndex1), _
                                                     intOldConstituentWidth, _
                                                     intOldConstituentHeight, _
                                                     booChangeFontSize:=booChangeFontSize, _
                                                     sglMinFontSize:=sglMinFontSize, _
                                                     sglMaxFontSize:=sglMaxFontSize, _
                                                     sglFontSizeTuner:=sglFontSizeTuner) Then
                        Return False
                    End If
                    If booChangeFontSize Then
                        sglEmsize = .Font.Size * CSng(dblHeightRatio)
                        If sglMinFontSize <> -1 Then
                            sglEmsize = Math.Max(sglEmsize, sglMinFontSize)
                        End If
                        If sglMaxFontSize <> -1 Then
                            sglEmsize = Math.Min(sglEmsize, sglMaxFontSize)
                        End If
                        sglEmsize = sglEmsize * sglFontSizeTuner
                        .Font = New System.Drawing.Font _
                                    (ctlHandle.Controls(intIndex1).Font.FontFamily, _
                                     sglEmsize)
                    End If
                End With
            Next intIndex1
            Return True
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Return the screen height
    '
    '
    ' This method returns the height of the primary screen as its function
    ' value (see also screenWidth and screenSize). It will return -1 for 
    ' the height if there is no primary screen. It will return -2 for 
    ' the height if for some reason, there are multiple primary screens.
    '
    ' The width will be the working width.  In particular, note that the 
    ' height will not include the space occupied by a visible Task bar.
    '
    '
    Public Shared Function screenHeight() As Integer
        Dim intHeight As Integer
        Dim intWidth As Integer
        screenSize(intWidth, intHeight)
        Return (intHeight)
    End Function

    ' ----------------------------------------------------------------------
    ' Return the screen size
    '
    '
    ' This method returns the width and the height of the primary screen in
    ' reference parameters (see also screenWidth and screenHeight). It will
    ' return -1 for width and for height if there is no primary screen. It
    ' will return -2 for width and for height if for some reason, there are
    ' multiple primary screens.
    '
    ' Both the width and the height are the working width and height. In
    ' particular, note that the height will not include the space occupied by
    ' a visible Task bar.
    '
    '
    Public Shared Sub screenSize(ByRef intWidth As Integer, _
                                 ByRef intHeight As Integer)
        Dim intIndex1 As Integer
        Dim intUpperBound As Integer
        Dim objScreens() As System.Windows.Forms.Screen = _
            System.Windows.Forms.Screen.AllScreens
        intWidth = -1 : intHeight = -1
        For intIndex1 = 0 To objScreens.GetUpperBound(0)
            If objScreens(intIndex1).Primary Then
                If intWidth = -1 AndAlso intHeight = -1 Then
                    intWidth = objScreens(intIndex1).WorkingArea.Width
                    intHeight = objScreens(intIndex1).WorkingArea.Height
                Else
                    intWidth = -2 : intHeight = -2
                End If
            End If
        Next intIndex1
        Select Case intWidth
            Case -1
                _OBJutilities.errorHandler("Cannot find the screen dimensions", _
                                           "windowsUtilities", _
                                           "screenSize")
            Case -2
                _OBJutilities.errorHandler("Multiple screen dimensions found", _
                                           "windowsUtilities", _
                                           "screenSize")
        End Select
    End Sub

    ' ----------------------------------------------------------------------
    ' Return the screen width
    '
    '
    ' This method returns the width of the primary screen as its function
    ' value (see also screenHeight and screenSize). It will return -1 for 
    ' the width if there is no primary screen. It will return -2 for 
    ' the width if for some reason, there are multiple primary screens.
    '
    ' The width will be the working width.
    '
    '
    Public Shared Function screenWidth() As Integer
        Dim intHeight As Integer
        Dim intWidth As Integer
        screenSize(intWidth, intHeight)
        Return (intWidth)
    End Function

    ' ----------------------------------------------------------------------
    ' Search a list box
    '
    '
    ' This method searches the items in lstBox for the target in strTarget,
    ' returning the index of the first match or -1 when the target cannot
    ' be found.
    '
    ' By default the search ignores case as well as leading and trailing 
    ' spaces in both the list box and the target; but the optional parameter
    ' booCaseSensitive can be passed as True to force a case-sensitive search,
    ' and the optional parameter booTrim can be passed as False to avoid
    ' ignoring leading and trailing spaces.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------------
    ' 05 18 03   Nilges         Bug in searching an empty list box, dammit
    '
    '
    Public Shared Function searchListBox(ByRef lstBox As ListBox, _
                                         ByVal strTarget As String, _
                                         Optional ByVal booCaseSensitive As Boolean = False, _
                                         Optional ByVal booTrim As Boolean = True) As Integer
        Dim intIndex1 As Integer
        Dim strTargetWork As String = strTarget
        Dim strNext As String
        If Not booCaseSensitive Then strTargetWork = UCase(strTargetWork)
        If booTrim Then strTargetWork = Trim(strTargetWork)
        With lstBox
            For intIndex1 = 0 To .Items.Count - 1
                strNext = CStr(.Items(intIndex1))
                If Not booCaseSensitive Then strNext = UCase(strNext)
                If booTrim Then strNext = Trim(strNext)
                If strTargetWork = strNext Then Exit For
            Next intIndex1
            If intIndex1 >= .Items.Count Then Return (-1)
            Return (intIndex1)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Initialize control to a group box
    '
    '
    ' This method deposits a group box in a control such that this group box
    ' fills the control, except for a surrounding border.  It sets the dimensions
    ' of the control to a specified width and height or the default width and
    ' height of a group box plus the border size.  It returns the group box in
    ' a ByRef parameter.
    '
    ' This method has several overloads:
    '
    '
    '      *  setControlToGroupBox(c, g) sets the existing control c to contain 
    '         the new group box g, where g has default width and height, and is
    '         surrounded by a border of default width and height (this will be
    '         the value returned by the Grid property.)
    '
    '      *  setControlToGroupBox(c, g, w) sets the existing control c to contain 
    '         the new group box g with the specified width w, and default height 
    '         and border.
    '
    '      *  setControlToGroupBox(c, g, w, h) sets the existing control c to contain 
    '         the new group box g with the specified width w, the specified height
    '         h, and the default border.
    '
    '      *  setControlToGroupBox(c, g, w, h, b) sets the existing control c to contain 
    '         the new group box g with the specified width w, the specified height
    '         h, and the border size b.
    '
    '
    ' Any one of the parameters w, h or b can be a negative integer.  If w, h or b is a
    ' negative integer, the absolute value of this integer is multiplied times the 
    ' default height or width of a Group box, or Grid, to obtain the corresponding
    ' value.  For example, setControlToGroupBox(c, g, -1, -2) creates a group box of
    ' default width and twice the default height.
    '
    '
    Public Overloads Shared Function setControlToGroupBox(ByVal ctlControl As Control, _
                                                          ByRef gbxGroupBox As GroupBox) As Boolean
        Return (setControlToGroupBox(ctlControl, _
                                    gbxGroupBox, _
                                    -1, -1, _
                                    -1))
    End Function
    Public Overloads Shared Function setControlToGroupBox(ByVal ctlControl As Control, _
                                                          ByRef gbxGroupBox As GroupBox, _
                                                          ByVal intWidth As Integer) As Boolean
        Return (setControlToGroupBox(ctlControl, _
                                    gbxGroupBox, _
                                    intWidth, -1, _
                                    -1))
    End Function
    Public Overloads Shared Function setControlToGroupBox(ByVal ctlControl As Control, _
                                                          ByRef gbxGroupBox As GroupBox, _
                                                          ByVal intWidth As Integer, _
                                                          ByVal intHeight As Integer) As Boolean
        Return (setControlToGroupBox(ctlControl, _
                                    gbxGroupBox, _
                                    intWidth, intHeight, _
                                    -1))
    End Function
    Public Overloads Shared Function setControlToGroupBox(ByVal ctlControl As Control, _
                                                          ByRef gbxGroupBox As GroupBox, _
                                                          ByVal intWidth As Integer, _
                                                          ByVal intHeight As Integer, _
                                                          ByVal intBorder As Integer) As Boolean
        Try
            If Not (gbxGroupBox Is Nothing) Then
                ' Cleanly get rid of existing box
                gbxGroupBox.Dispose()
                gbxGroupBox = Nothing
            End If
            gbxGroupBox = New GroupBox
            With gbxGroupBox
                .Left = CInt(IIf(intBorder < 0, Grid * -intBorder, intBorder))
                .Top = .Left
                .Width = CInt(IIf(intWidth < 0, .Width * -intWidth, intWidth))
                .Height = CInt(IIf(intHeight < 0, .Height * -intHeight, intHeight))
                ctlControl.Width = .Width + .Left * 2
                ctlControl.Height = .Height + .Left * 2
            End With
            ctlControl.Controls.Add(gbxGroupBox)
        Catch
            _OBJutilities.errorHandler("Cannot set control to group box: " & _
                                       Err.Number & " " & Err.Description, _
                                       "windowsUtilities", _
                                       "setControlToGroupBox", _
                                       "Returning False")
            Return (False)
        End Try
        Return (True)
    End Function

    ' ---------------------------------------------------------------------
    ' Sort string in ascending order using a list box
    '
    '
    ' Sorting is always in ascending order using the entire string and
    ' the rules applicable to Sorted list boxes. By default, lines are
    ' assumed to be separated by newline: the optional strSep parameter
    ' may specify a different separator.
    '
    '
    Public Shared Function sortUsingListbox(ByVal strInstring As String, _
                                            Optional ByVal strSep As String = vbNewLine) _
           As String
        Dim lstTemp As ListBox
        Try
            lstTemp = New ListBox
        Catch ex As Exception
            _OBJutilities.errorHandler("Can't create temporary list box", _
                                       "windowsUtilities", _
                                       "sortUsingListbox", _
                                       Err.Number & " " & Err.Description)
            Return (strInstring)
        End Try
        With lstTemp
            .Sorted = True
            If Not string2Listbox(strInstring, _
                                  lstTemp, _
                                  objProgress:=False, _
                                  strSep:=strSep) Then Return (strInstring)
            Return (listBox2String(lstTemp, strSep:=strSep))
        End With
    End Function

    ' ---------------------------------------------------------------------
    ' String to list box
    '
    '
    ' This method places a string into a list box.  By default, it assumes
    ' that lines are separated by newlines, or, the strSep parameter can
    ' define a different line separator.  
    '
    ' By default, the list box is cleared, or, the booAppend:=True parameter 
    ' will cause the lines to be appended to the list box.
    '
    ' By default, when the list box does not have the Sorted property, progress 
    ' is shown by highlighting each add, and, for each add, refreshing and 
    ' focusing-on the control. The objProgress:=False parameter will suppress 
    ' this progress report for non-Sorted list boxes. This parameter defaults to
    ' False when the progress report is Sorted.
    '
    '
    Public Shared Function string2Listbox(ByVal strInstring As String, _
                                            ByVal lstBox As ListBox, _
                                            Optional ByVal strSep As String = vbNewLine, _
                                            Optional ByVal booAppend As Boolean = False, _
                                            Optional ByVal objProgress As Object = Nothing) As Boolean
        If strSep = "" Then
            _OBJutilities.errorHandler("Null line separator is not allowed", _
                                       "windowsUtilities", _
                                       "string2Listbox", _
                                       "No change will be made to the list box")
            Return (False)
        End If
        Dim booProgressWork As Boolean
        If (objProgress Is Nothing) Then
            booProgressWork = Not lstBox.Sorted
        ElseIf (TypeOf objProgress Is System.Boolean) Then
            booProgressWork = CType(objProgress, Boolean)
        Else
            _OBJutilities.errorHandler("objProgress has invalid type " & _
                                       objProgress.GetType.ToString & ": " & _
                                       "type must be Boolean or a Nothing object", _
                                       "windowsUtilities", _
                                       "string2Listbox", _
                                       "")
        End If
        Dim strSplit() As String
        Try
            strSplit = Split(strInstring, strSep)
        Catch ex As Exception
            _OBJutilities.errorHandler("Can't split: " & _
                                       Err.Number & " " & Err.Description, _
                                       "windowsUtilities", _
                                       "string2Listbox", _
                                       "No change will be made to the list box")
            Return (False)
        End Try
        With lstBox
            If Not booAppend Then .Items.Clear()
            Dim intIndex1 As Integer
            For intIndex1 = 0 To UBound(strSplit)
                Try
                    .Items.Add(strSplit(intIndex1))
                Catch ex As Exception
                    _OBJutilities.errorHandler("Can't split: " & _
                                               Err.Number & " " & Err.Description, _
                                               "windowsUtilities", _
                                               "string2Listbox", _
                                               "No change will be made to the list box")
                    Return (False)
                End Try
                If booProgressWork Then
                    .SelectedIndex = .Items.Count - 1
                    .Focus()
                    .Refresh()
                End If
            Next intIndex1
        End With
        Return (True)
    End Function

    ' ---------------------------------------------------------------------
    ' Unimplemented test method
    '
    '
    Public Shared Function test(ByRef strReport As String) As Boolean
        strReport = "Test functionality hasn't been implemented"
        test = True
    End Function

    ' ---------------------------------------------------------------------
    ' Trim the form or container control
    '
    '
    ' This method adjusts a form or other container control around the
    ' bottom and the rightmost contained control. 
    '
    ' By default, a space equal to GRID pixels (where GRID is defined in
    ' this class as the GRID property) is left below the bottom control
    ' and to the right of the rightmost control.  You may pass overriding
    ' values for the bottom and for the right side using the optional parameters 
    ' intBottomBorder:=n and intRightBorder:=m.
    '
    ' Note that when these parameters have negative values, controls at the 
    ' bottom and right side of your form or container may overlap the edges 
    ' with no error indication.
    '
    ' Note that depending on the presence of a menu bar or other special conditions
    ' you may have to experiment a bit, passing GRID * c where c is a constant
    ' such as 2, 4 or 8, to get the desired effect.
    '
    ' If the progressReport class is available and the UTILITIES_PROGRESS
    ' compile-time symbol is True then you may pass an object of type
    ' progressReport in the optional objProgressReport parameter, and this
    ' object will provide a visual indication of progress during this method's
    ' search for the bottom and the right-most controls.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------
    ' 11 15 02   Nilges         Don't finish the progressReport
    ' 01 11 03   Nilges         Method should be Shared
    '
    '
    Public Overloads Shared Function trimContainer(ByRef frmContainer As Form, _
                                                    Optional ByVal intBottomBorder As Integer = GRID_, _
                                                    Optional ByVal intRightBorder As Integer = GRID_, _
                                                    Optional ByRef objProgressReport As Object = Nothing) _
           As Boolean
        Return (trimContainer(CType(frmContainer, Control), _
                             intBottomBorder:=intBottomBorder, _
                             intRightBorder:=intRightBorder, _
                             objProgressReport:=objProgressReport))
    End Function
    Public Overloads Shared Function trimContainer(ByRef ctlContainer As Control, _
                                                    Optional ByVal intBottomBorder As Integer = GRID_, _
                                                    Optional ByVal intRightBorder As Integer = GRID_, _
                                                    Optional ByRef objProgressReport As Object = Nothing) _
           As Boolean
        Dim objProgressReportHandle As Object
#If UTILITIES_PROGRESS Then
            If Not (objProgressReport Is Nothing) _
               AndAlso _
               Not (TypeOf objProgressReport Is progressReport.progressReport) Then
                errorHandler("objProgressReport isn't of type progressReport", _
                             "utilities", "trimContainer") 
                Return(False)                             
            End If
            objProgressReportHandle = objProgressReport
            With CType(objProgressReportHandle, progressReport.progressReport)
                .Activity = "Searching for container boundaries"
                .Entity = "contained control"
                .EntityCount = ctlContainer.Controls.Count                    
            End With                
#End If
        Dim ctlBottom As Control
        Dim ctlRightMost As Control
        Dim intBottom As Integer
        Dim intRight As Integer
        Dim intIndex1 As Integer
        With ctlContainer
            For intIndex1 = 0 To .Controls.Count - 1
                If ctlBottom Is Nothing Then
                    ctlBottom = .Controls(intIndex1)
                    ctlRightMost = .Controls(intIndex1)
                Else
                    intBottom = controlBottom(.Controls(intIndex1))
                    If intBottom > controlBottom(ctlBottom) Then
                        ctlBottom = .Controls(intIndex1)
                    End If
                    intRight = controlRight(.Controls(intIndex1))
                    If intRight > controlBottom(ctlRightMost) Then
                        ctlRightMost = .Controls(intIndex1)
                    End If
                End If
#If UTILITIES_PROGRESS Then
                    If Not (objProgressReportHandle Is Nothing) Then
                        CType(objProgressReportHandle, progressReport.progressReport).refresh(vbNewline & vbNewline & _
                                                        "At control " & _
                                                        enquote(.Controls.Item(intIndex1).Name) & _
                                                        vbNewline & vbNewline & _
                                                        "Bottom control at this time is " & _
                                                        enquote(ctlBottom.Name) & _
                                                        vbNewline & _
                                                        "Rightmost control at this time is " & _
                                                        enquote(ctlRightMost.Name))
                    End If                    
#End If
            Next intIndex1
            .Width = controlRight(ctlRightMost) + Grid * 2
            .Height = controlBottom(ctlBottom) + Grid * 8
            .Refresh()
        End With
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Update the status list box
    '
    '
    ' This method updates a status list box with status reports that are
    ' time and date stamped, and which may be indented to show the
    ' progress of nested processes. This method can format the status
    ' report into multiple lines for narrower list box.
    '
    ' The message is appended to the end of the list box, and the new
    ' entry is focused. 
    '
    ' This method uses the Tag field of the status list box to contain
    ' the current indentation level, and use of this method assumes your
    ' code won't use this Tag.
    '
    ' This method has the following overloads.
    '
    '
    '      *  updateStatusListBox(statusListBox, levelChangeBefore, message, 
    '         levelChangeAfter): appends message to the end of the status report. 
    '         The level is changed by the signed integer value in levelChangeBefore 
    '         BEFORE the message is added: the level is changed by the signed integer 
    '         value in levelChangeAfter AFTER the message is added.
    '
    '      *  updateStatusListBox(statusListBox, levelChangeBefore, message): 
    '         appends message to the end of the status report. The level is changed 
    '         by the signed integer value in levelChangeBefore BEFORE the message is added.
    '
    '      *  updateStatusListBox(statusListBox, message, levelChangeAfter): appends 
    '         message to the end of the status report. The level is changed by the signed 
    '         integer value in levelChangeAfter AFTER the message is added.
    '
    '      *  updateStatusListBox(statusListBox, message): appends message to the end of the 
    '         status report. The level is not changed.
    '
    '      *  updateStatusListBox(statusListBox): appends no message, and instead sets the
    '         level to 0.
    '
    '      *  updateStatusListBox(statusListBox, levelChange): appends no message, and 
    '         instead adds or subtracts levelChange to or from the level.
    '
    '
    ' All overloads except the last two above support the optional bytLineWidth,
    ' booSplitLines, booIncludeDate and strIndent parameters.
    '
    ' If the bytLineWidth parameter is present and not 0, the line added is restricted to 
    ' the specified width. When the message (and its date and time prefix) exceed the 
    ' specified width the line is broken up into two or more lines.
    '
    ' If the booSplitLines parameter is absent or False, newlines in the status message
    ' are replaced by \n: if this parameter is present and True, then multiple lines are
    ' added to the report.
    '
    ' If the booIncludeDate parameter is absent or True, each line appended to the status
    ' report will contain the date and the time. If this parameter is present and False,
    ' the date and the time won't be added.
    '
    ' The strIndent parameter is the spacing or other character string that is used to
    ' prefix the line added to the status list box, with 1 copy of this string for each
    ' level. strIndent defaults to "    " (four spaces).
    '
    ' Note that for best results, the Font of the list box should be a monospaced font such
    ' as Courier New.
    '
    '
    ' --- Change the level, add the message, change the level
    Public Overloads Shared Function updateStatusListBox(ByRef lstStatus As ListBox, _
                                                            ByVal intBefore As Integer, _
                                                            ByVal strMessage As String, _
                                                            ByVal intAfter As Integer, _
                                                            Optional ByVal bytLineWidth As Byte = 0, _
                                                            Optional ByVal booSplitLines As Boolean = False, _
                                                            Optional ByVal booIncludeDate As Boolean = True, _
                                                            Optional ByVal strIndent As String = "    ") _
           As Boolean
        Return (updateStatusListBox_(lstStatus, _
                                     intBefore, _
                                     True, _
                                     strMessage, _
                                     intAfter, _
                                     False, _
                                     0, _
                                     bytLineWidth, _
                                     booSplitLines, _
                                     booIncludeDate, _
                                     strIndent))
    End Function
    ' --- Change the level, add the message 
    Public Overloads Shared Function updateStatusListBox(ByRef lstStatus As ListBox, _
                                                            ByVal intBefore As Integer, _
                                                            ByVal strMessage As String, _
                                                            Optional ByVal bytLineWidth As Byte = 0, _
                                                            Optional ByVal booSplitLines As Boolean = False, _
                                                            Optional ByVal booIncludeDate As Boolean = True, _
                                                            Optional ByVal strIndent As String = "    ") _
           As Boolean
        Return (updateStatusListBox_(lstStatus, _
                                     intBefore, _
                                     True, _
                                     strMessage, _
                                     0, _
                                     False, _
                                     0, _
                                     bytLineWidth, _
                                     booSplitLines, _
                                     booIncludeDate, _
                                     strIndent))
    End Function
    ' --- Add the message, change the level
    Public Overloads Shared Function updateStatusListBox(ByRef lstStatus As ListBox, _
                                                            ByVal strMessage As String, _
                                                            ByVal intAfter As Integer, _
                                                            Optional ByVal bytLineWidth As Byte = 0, _
                                                            Optional ByVal booSplitLines As Boolean = False, _
                                                            Optional ByVal booIncludeDate As Boolean = True, _
                                                            Optional ByVal strIndent As String = "    ") _
           As Boolean
        Return (updateStatusListBox_(lstStatus, _
                                     0, _
                                     True, _
                                     strMessage, _
                                     intAfter, _
                                     False, _
                                     0, _
                                     bytLineWidth, _
                                     booSplitLines, _
                                     booIncludeDate, _
                                     strIndent))
    End Function
    ' --- Add the message 
    Public Overloads Shared Function updateStatusListBox(ByRef lstStatus As ListBox, _
                                                         ByVal strMessage As String, _
                                                         Optional ByVal bytLineWidth As Byte = 0, _
                                                         Optional ByVal booSplitLines As Boolean = False, _
                                                         Optional ByVal booIncludeDate As Boolean = True, _
                                                         Optional ByVal strIndent As String = "    ") _
           As Boolean
        Return (updateStatusListBox_(lstStatus, _
                                     0, _
                                     True, _
                                     strMessage, _
                                     0, _
                                     False, _
                                     0, _
                                     bytLineWidth, _
                                     booSplitLines, _
                                     booIncludeDate, _
                                     strIndent))
    End Function
    ' --- Reset the level
    Public Overloads Shared Function updateStatusListBox(ByRef lstStatus As ListBox) _
           As Boolean
        Return (updateStatusListBox_(lstStatus, _
                                     0, _
                                     False, _
                                     "", _
                                     0, _
                                     False, _
                                     0, _
                                     0, _
                                     False, _
                                     False, _
                                     ""))
    End Function
    ' --- Common logic
    Private Shared Function updateStatusListBox_(ByRef lstStatus As ListBox, _
                                                 ByVal intBefore As Integer, _
                                                 ByVal booMessage As Boolean, _
                                                 ByVal strMessage As String, _
                                                 ByVal intAfter As Integer, _
                                                 ByVal booNewlevel As Boolean, _
                                                 ByVal intNewlevel As Integer, _
                                                 ByVal bytLineWidth As Byte, _
                                                 ByVal booSplitLines As Boolean, _
                                                 ByVal booIncludeDate As Boolean, _
                                                 ByVal strIndent As String) _
           As Boolean
        With lstStatus
            Dim intLevel As Integer
            Try
                intLevel = CInt(.Tag)
            Catch : End Try
            intLevel += intBefore
            intLevel = Math.Max(intLevel, 0)
            If booMessage Then
                Dim strPrefix As String = _OBJutilities.copies(strIndent, intLevel)
                If booIncludeDate Then strPrefix &= Now
                strPrefix &= " "
                If booSplitLines Then
                    Dim intIndex1 As Integer
                    For intIndex1 = 1 To _OBJutilities.items(strMessage, _
                                                             vbNewLine, _
                                                             False)
                        updateStatusListBox__displayOneLine_(lstStatus, _
                                                             strPrefix, _
                                                             _OBJutilities.item(strMessage, _
                                                                                intIndex1, _
                                                                                vbNewLine, _
                                                                                False), _
                                                             bytLineWidth)
                    Next intIndex1
                Else
                    updateStatusListBox__displayOneLine_(lstStatus, _
                                                         strPrefix, _
                                                         strMessage, _
                                                         bytLineWidth)
                End If
            End If
            intLevel += intAfter
            If booNewlevel Then intLevel = intNewlevel
            .Tag = intLevel
            Return (True)
        End With
    End Function

    ' ---------------------------------------------------------------------
    ' Display one line on behalf of updateStatusListBox
    '
    '
    Private Shared Sub updateStatusListBox__displayOneLine_(ByRef lstStatus As ListBox, _
                                                            ByVal strPrefix As String, _
                                                            ByVal strMessage As String, _
                                                            ByVal bytLineWidth As Byte)
        With lstStatus
            Dim strFullMessage As String = strPrefix & Replace(strMessage, vbNewLine, "\n")
            Dim strIndent As String = _OBJutilities.copies(" ", Len(strPrefix))
            If bytLineWidth <> 0 AndAlso Len(strFullMessage) > bytLineWidth Then
                strFullMessage = _OBJutilities.soft2HardParagraph(strFullMessage, _
                                                                  bytLineWidth:=bytLineWidth)
            End If
            Dim intIndex1 As Integer
            For intIndex1 = 1 To _OBJutilities.items(strFullMessage, vbNewLine, False)
                .Items.Add(_OBJutilities.item(strFullMessage, _
                            intIndex1, _
                            vbNewLine, _
                            False))
            Next intIndex1
            .Focus()
            .SelectedIndex = .Items.Count - 1
            .Refresh()
        End With
    End Sub

End Class

