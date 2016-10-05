Option Strict

' *********************************************************************
' *                                                                   *
' * collectionTest: form to experiment with the collection utilities  *
' *                                                                   *
' *********************************************************************

Public Class collectionTest
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
    Friend WithEvents lblCollectionsAvailable As System.Windows.Forms.Label
    Friend WithEvents lstCollectionsAvailable As System.Windows.Forms.ListBox
    Friend WithEvents cmdCollection2String As System.Windows.Forms.Button
    Friend WithEvents cmdRegistry2Form As System.Windows.Forms.Button
    Friend WithEvents cmdForm2Registry As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents cmdCollectionCompare As System.Windows.Forms.Button
    Friend WithEvents cmdCopyCollection As System.Windows.Forms.Button
    Friend WithEvents cmdStressTest As System.Windows.Forms.Button
    Friend WithEvents lstStatus As System.Windows.Forms.ListBox
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents cmdStatusZoom As System.Windows.Forms.Button
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTools As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsRandomCollectionControl As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelp As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAbout As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileExitNoSave As System.Windows.Forms.MenuItem
    Friend WithEvents cmdCollectionClear As System.Windows.Forms.Button
    Friend WithEvents cmdCollectionTypes As System.Windows.Forms.Button
    Friend WithEvents cmdClearRegistry As System.Windows.Forms.Button
    Friend WithEvents cmdMkRandomCollections As System.Windows.Forms.Button
    Friend WithEvents gbxCollection2String As System.Windows.Forms.GroupBox
    Friend WithEvents chkCollection2StringReadable As System.Windows.Forms.CheckBox
    Friend WithEvents chkCollection2StringInspect As System.Windows.Forms.CheckBox
    Friend WithEvents cmdString2Collection As System.Windows.Forms.Button
    Friend WithEvents lstSerializedCollection As System.Windows.Forms.ListBox
    Friend WithEvents txtSerializedCollection As System.Windows.Forms.TextBox
    Friend WithEvents lblSerializedCollection As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.lblCollectionsAvailable = New System.Windows.Forms.Label
Me.lstCollectionsAvailable = New System.Windows.Forms.ListBox
Me.cmdCollection2String = New System.Windows.Forms.Button
Me.cmdRegistry2Form = New System.Windows.Forms.Button
Me.cmdForm2Registry = New System.Windows.Forms.Button
Me.cmdClose = New System.Windows.Forms.Button
Me.cmdCopyCollection = New System.Windows.Forms.Button
Me.cmdCollectionCompare = New System.Windows.Forms.Button
Me.cmdStressTest = New System.Windows.Forms.Button
Me.lstStatus = New System.Windows.Forms.ListBox
Me.lblProgress = New System.Windows.Forms.Label
Me.lblStatus = New System.Windows.Forms.Label
Me.cmdStatusZoom = New System.Windows.Forms.Button
Me.MainMenu1 = New System.Windows.Forms.MainMenu
Me.mnuFile = New System.Windows.Forms.MenuItem
Me.mnuFileExitNoSave = New System.Windows.Forms.MenuItem
Me.mnuFileExit = New System.Windows.Forms.MenuItem
Me.mnuTools = New System.Windows.Forms.MenuItem
Me.mnuToolsRandomCollectionControl = New System.Windows.Forms.MenuItem
Me.mnuHelp = New System.Windows.Forms.MenuItem
Me.mnuAbout = New System.Windows.Forms.MenuItem
Me.cmdCollectionClear = New System.Windows.Forms.Button
Me.cmdCollectionTypes = New System.Windows.Forms.Button
Me.cmdClearRegistry = New System.Windows.Forms.Button
Me.cmdMkRandomCollections = New System.Windows.Forms.Button
Me.gbxCollection2String = New System.Windows.Forms.GroupBox
Me.chkCollection2StringInspect = New System.Windows.Forms.CheckBox
Me.chkCollection2StringReadable = New System.Windows.Forms.CheckBox
Me.cmdString2Collection = New System.Windows.Forms.Button
Me.lstSerializedCollection = New System.Windows.Forms.ListBox
Me.txtSerializedCollection = New System.Windows.Forms.TextBox
Me.lblSerializedCollection = New System.Windows.Forms.Label
Me.gbxCollection2String.SuspendLayout()
Me.SuspendLayout()
'
'lblCollectionsAvailable
'
Me.lblCollectionsAvailable.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(192, Byte))
Me.lblCollectionsAvailable.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblCollectionsAvailable.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblCollectionsAvailable.Location = New System.Drawing.Point(8, 4)
Me.lblCollectionsAvailable.Name = "lblCollectionsAvailable"
Me.lblCollectionsAvailable.Size = New System.Drawing.Size(889, 64)
Me.lblCollectionsAvailable.TabIndex = 2
Me.lblCollectionsAvailable.Text = "Collections Available"
Me.lblCollectionsAvailable.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
'
'lstCollectionsAvailable
'
Me.lstCollectionsAvailable.BackColor = System.Drawing.SystemColors.Control
Me.lstCollectionsAvailable.Location = New System.Drawing.Point(8, 68)
Me.lstCollectionsAvailable.Name = "lstCollectionsAvailable"
Me.lstCollectionsAvailable.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
Me.lstCollectionsAvailable.Size = New System.Drawing.Size(889, 108)
Me.lstCollectionsAvailable.Sorted = True
Me.lstCollectionsAvailable.TabIndex = 3
'
'cmdCollection2String
'
Me.cmdCollection2String.BackColor = System.Drawing.SystemColors.Control
Me.cmdCollection2String.Enabled = False
Me.cmdCollection2String.Location = New System.Drawing.Point(8, 15)
Me.cmdCollection2String.Name = "cmdCollection2String"
Me.cmdCollection2String.Size = New System.Drawing.Size(120, 20)
Me.cmdCollection2String.TabIndex = 7
Me.cmdCollection2String.Text = "Collection to String"
'
'cmdRegistry2Form
'
Me.cmdRegistry2Form.Location = New System.Drawing.Point(120, 312)
Me.cmdRegistry2Form.Name = "cmdRegistry2Form"
Me.cmdRegistry2Form.Size = New System.Drawing.Size(104, 24)
Me.cmdRegistry2Form.TabIndex = 9
Me.cmdRegistry2Form.Text = "Restore Settings"
'
'cmdForm2Registry
'
Me.cmdForm2Registry.Location = New System.Drawing.Point(8, 312)
Me.cmdForm2Registry.Name = "cmdForm2Registry"
Me.cmdForm2Registry.Size = New System.Drawing.Size(104, 24)
Me.cmdForm2Registry.TabIndex = 10
Me.cmdForm2Registry.Text = "Save Settings"
'
'cmdClose
'
Me.cmdClose.Location = New System.Drawing.Point(792, 312)
Me.cmdClose.Name = "cmdClose"
Me.cmdClose.Size = New System.Drawing.Size(104, 24)
Me.cmdClose.TabIndex = 11
Me.cmdClose.Text = "Close"
'
'cmdCopyCollection
'
Me.cmdCopyCollection.Enabled = False
Me.cmdCopyCollection.Location = New System.Drawing.Point(768, 12)
Me.cmdCopyCollection.Name = "cmdCopyCollection"
Me.cmdCopyCollection.Size = New System.Drawing.Size(120, 20)
Me.cmdCopyCollection.TabIndex = 12
Me.cmdCopyCollection.Text = "Copy Collection"
'
'cmdCollectionCompare
'
Me.cmdCollectionCompare.Enabled = False
Me.cmdCollectionCompare.Location = New System.Drawing.Point(640, 12)
Me.cmdCollectionCompare.Name = "cmdCollectionCompare"
Me.cmdCollectionCompare.Size = New System.Drawing.Size(120, 20)
Me.cmdCollectionCompare.TabIndex = 13
Me.cmdCollectionCompare.Text = "Compare Collections"
'
'cmdStressTest
'
Me.cmdStressTest.Location = New System.Drawing.Point(528, 312)
Me.cmdStressTest.Name = "cmdStressTest"
Me.cmdStressTest.Size = New System.Drawing.Size(104, 24)
Me.cmdStressTest.TabIndex = 15
Me.cmdStressTest.Text = "Stress Test"
'
'lstStatus
'
Me.lstStatus.BackColor = System.Drawing.SystemColors.Control
Me.lstStatus.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lstStatus.ItemHeight = 14
Me.lstStatus.Location = New System.Drawing.Point(7, 364)
Me.lstStatus.Name = "lstStatus"
Me.lstStatus.Size = New System.Drawing.Size(889, 88)
Me.lstStatus.TabIndex = 16
'
'lblProgress
'
Me.lblProgress.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(192, Byte))
Me.lblProgress.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblProgress.Location = New System.Drawing.Point(8, 460)
Me.lblProgress.Name = "lblProgress"
Me.lblProgress.Size = New System.Drawing.Size(888, 16)
Me.lblProgress.TabIndex = 17
Me.lblProgress.Visible = False
'
'lblStatus
'
Me.lblStatus.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(192, Byte))
Me.lblStatus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblStatus.Location = New System.Drawing.Point(8, 344)
Me.lblStatus.Name = "lblStatus"
Me.lblStatus.Size = New System.Drawing.Size(888, 20)
Me.lblStatus.TabIndex = 18
Me.lblStatus.Text = "Status"
Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
'
'cmdStatusZoom
'
Me.cmdStatusZoom.Location = New System.Drawing.Point(792, 344)
Me.cmdStatusZoom.Name = "cmdStatusZoom"
Me.cmdStatusZoom.Size = New System.Drawing.Size(104, 20)
Me.cmdStatusZoom.TabIndex = 19
Me.cmdStatusZoom.Text = "Zoom"
'
'MainMenu1
'
Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuTools, Me.mnuHelp})
'
'mnuFile
'
Me.mnuFile.Index = 0
Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFileExitNoSave, Me.mnuFileExit})
Me.mnuFile.Text = "File"
'
'mnuFileExitNoSave
'
Me.mnuFileExitNoSave.Index = 0
Me.mnuFileExitNoSave.Text = "Exit (don't save form settings)"
'
'mnuFileExit
'
Me.mnuFileExit.Index = 1
Me.mnuFileExit.Text = "Exit (save form settings)"
'
'mnuTools
'
Me.mnuTools.Index = 1
Me.mnuTools.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuToolsRandomCollectionControl})
Me.mnuTools.Text = "Tools"
'
'mnuToolsRandomCollectionControl
'
Me.mnuToolsRandomCollectionControl.Index = 0
Me.mnuToolsRandomCollectionControl.Text = "Control random collections and stress tests..."
'
'mnuHelp
'
Me.mnuHelp.Index = 2
Me.mnuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAbout})
Me.mnuHelp.Text = "Help"
'
'mnuAbout
'
Me.mnuAbout.Index = 0
Me.mnuAbout.Text = "About..."
'
'cmdCollectionClear
'
Me.cmdCollectionClear.Enabled = False
Me.cmdCollectionClear.Location = New System.Drawing.Point(640, 36)
Me.cmdCollectionClear.Name = "cmdCollectionClear"
Me.cmdCollectionClear.Size = New System.Drawing.Size(120, 20)
Me.cmdCollectionClear.TabIndex = 20
Me.cmdCollectionClear.Text = "Collection Clear"
'
'cmdCollectionTypes
'
Me.cmdCollectionTypes.Enabled = False
Me.cmdCollectionTypes.Location = New System.Drawing.Point(768, 36)
Me.cmdCollectionTypes.Name = "cmdCollectionTypes"
Me.cmdCollectionTypes.Size = New System.Drawing.Size(120, 20)
Me.cmdCollectionTypes.TabIndex = 21
Me.cmdCollectionTypes.Text = "Collection Types"
'
'cmdClearRegistry
'
Me.cmdClearRegistry.Location = New System.Drawing.Point(232, 312)
Me.cmdClearRegistry.Name = "cmdClearRegistry"
Me.cmdClearRegistry.Size = New System.Drawing.Size(104, 24)
Me.cmdClearRegistry.TabIndex = 22
Me.cmdClearRegistry.Text = "Clear Settings"
'
'cmdMkRandomCollections
'
Me.cmdMkRandomCollections.Location = New System.Drawing.Point(640, 312)
Me.cmdMkRandomCollections.Name = "cmdMkRandomCollections"
Me.cmdMkRandomCollections.Size = New System.Drawing.Size(144, 24)
Me.cmdMkRandomCollections.TabIndex = 23
Me.cmdMkRandomCollections.Text = "Make Random Collections"
'
'gbxCollection2String
'
Me.gbxCollection2String.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(192, Byte))
Me.gbxCollection2String.Controls.Add(Me.chkCollection2StringInspect)
Me.gbxCollection2String.Controls.Add(Me.chkCollection2StringReadable)
Me.gbxCollection2String.Controls.Add(Me.cmdCollection2String)
Me.gbxCollection2String.Location = New System.Drawing.Point(328, 12)
Me.gbxCollection2String.Name = "gbxCollection2String"
Me.gbxCollection2String.Size = New System.Drawing.Size(296, 48)
Me.gbxCollection2String.TabIndex = 24
Me.gbxCollection2String.TabStop = False
'
'chkCollection2StringInspect
'
Me.chkCollection2StringInspect.Enabled = False
Me.chkCollection2StringInspect.Location = New System.Drawing.Point(224, 19)
Me.chkCollection2StringInspect.Name = "chkCollection2StringInspect"
Me.chkCollection2StringInspect.Size = New System.Drawing.Size(64, 16)
Me.chkCollection2StringInspect.TabIndex = 9
Me.chkCollection2StringInspect.Text = "Inspect"
'
'chkCollection2StringReadable
'
Me.chkCollection2StringReadable.Checked = True
Me.chkCollection2StringReadable.CheckState = System.Windows.Forms.CheckState.Checked
Me.chkCollection2StringReadable.Location = New System.Drawing.Point(136, 19)
Me.chkCollection2StringReadable.Name = "chkCollection2StringReadable"
Me.chkCollection2StringReadable.Size = New System.Drawing.Size(80, 16)
Me.chkCollection2StringReadable.TabIndex = 8
Me.chkCollection2StringReadable.Text = "Readable"
'
'cmdString2Collection
'
Me.cmdString2Collection.Location = New System.Drawing.Point(775, 180)
Me.cmdString2Collection.Name = "cmdString2Collection"
Me.cmdString2Collection.Size = New System.Drawing.Size(120, 20)
Me.cmdString2Collection.TabIndex = 28
Me.cmdString2Collection.Text = "String to Collection"
'
'lstSerializedCollection
'
Me.lstSerializedCollection.BackColor = System.Drawing.SystemColors.Control
Me.lstSerializedCollection.Location = New System.Drawing.Point(7, 220)
Me.lstSerializedCollection.Name = "lstSerializedCollection"
Me.lstSerializedCollection.Size = New System.Drawing.Size(888, 82)
Me.lstSerializedCollection.TabIndex = 27
'
'txtSerializedCollection
'
Me.txtSerializedCollection.Location = New System.Drawing.Point(7, 200)
Me.txtSerializedCollection.Name = "txtSerializedCollection"
Me.txtSerializedCollection.Size = New System.Drawing.Size(888, 20)
Me.txtSerializedCollection.TabIndex = 26
Me.txtSerializedCollection.Text = ""
'
'lblSerializedCollection
'
Me.lblSerializedCollection.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(192, Byte))
Me.lblSerializedCollection.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblSerializedCollection.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblSerializedCollection.Location = New System.Drawing.Point(7, 180)
Me.lblSerializedCollection.Name = "lblSerializedCollection"
Me.lblSerializedCollection.Size = New System.Drawing.Size(888, 20)
Me.lblSerializedCollection.TabIndex = 25
Me.lblSerializedCollection.Text = "Serialized Collection: serialized collections available"
Me.lblSerializedCollection.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
'
'collectionTest
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(902, 483)
Me.ControlBox = False
Me.Controls.Add(Me.cmdString2Collection)
Me.Controls.Add(Me.lstSerializedCollection)
Me.Controls.Add(Me.txtSerializedCollection)
Me.Controls.Add(Me.lblSerializedCollection)
Me.Controls.Add(Me.gbxCollection2String)
Me.Controls.Add(Me.cmdMkRandomCollections)
Me.Controls.Add(Me.cmdClearRegistry)
Me.Controls.Add(Me.cmdCollectionTypes)
Me.Controls.Add(Me.cmdCollectionClear)
Me.Controls.Add(Me.cmdStatusZoom)
Me.Controls.Add(Me.lblStatus)
Me.Controls.Add(Me.lblProgress)
Me.Controls.Add(Me.lstStatus)
Me.Controls.Add(Me.cmdStressTest)
Me.Controls.Add(Me.cmdCollectionCompare)
Me.Controls.Add(Me.cmdCopyCollection)
Me.Controls.Add(Me.cmdClose)
Me.Controls.Add(Me.cmdForm2Registry)
Me.Controls.Add(Me.cmdRegistry2Form)
Me.Controls.Add(Me.lstCollectionsAvailable)
Me.Controls.Add(Me.lblCollectionsAvailable)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
Me.Menu = Me.MainMenu1
Me.Name = "collectionTest"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
Me.Text = "collectionTest"
Me.gbxCollection2String.ResumeLayout(False)
Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Form data "

    Private _OBJutilities As utilities.utilities
    Private _OBJwindowsUtilities As windowsUtilities.windowsUtilities

    Private COLcollections() As Collection
    
    Private Const DEFAULT_SERIALIZED_COLLECTIONS As String = _
    ": null string" & vbNewline & _
    "Member 1 of 1: contains one member" & vbNewline & _
    "Member 1 of 2,Member 2 of 2: contains two members" & vbNewline & _
    "Member 1., Member 2., " & _
    "(Member 2.1&#00000, (Member 2.1.1, (Member 2.1.1.1, 45)), " & _
    "2.3): " & _
    "nests members and contains a variety of data"
    
    Private Const DEFAULT_COLLECTIONS As String = _
    "(Empty): an allocated collection with a Count of zero" & vbNewline & _
    "(Example): example in documentation" & vbNewline & _    
    "(Nothing): an unallocated collection" & vbNewline & _
    "(Random): a random collection which randomly includes subcollections" & vbNewline & _
    "(Something): an allocated collection with one entry that is a null string" & vbNewline & _ 
    "(The letter A): an allocated collection with one entry that is the letter A" 
    Private Const RANDOM_COLLECTION_INDEX As Integer = 3
        
    Private Const ABOUT_INFO As String = _
    "This form allows you to test the collection utilities including " & _
    "collection2String, string2Collection, collectionClear, " & _
    "collectionCompare, collectionCopy, collectionDepth and collectionTypes"
    
    Private FRMrandomCollectionControl As randomCollectionControl

#End Region ' Form data

#Region " Form events "

    Private Sub chkCollection2StringReadable_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCollection2StringReadable.CheckedChanged
        chkCollection2StringInspect.Enabled = Not chkCollection2StringReadable.Checked
    End Sub

    Private Sub cmdClearRegistry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearRegistry.Click
        clearRegistry
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        closer(True)            
    End Sub

    Private Sub cmdCollection2String_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCollection2String.Click
        collection2StringInterface(chkCollection2StringReadable.Checked, _
                                   chkCollection2StringInspect.Checked)
    End Sub

    Private Sub cmdCollectionClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCollectionClear.Click
        collectionClearInterface
    End Sub

    Private Sub cmdCollectionCompare_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCollectionCompare.Click
        With lstCollectionsAvailable
            collectionCompareInterface(getCollection(.SelectedIndices(0)), _
                                       getCollection(.SelectedIndices(1)))
        End With        
    End Sub

    Private Sub cmdCollectionTypes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCollectionTypes.Click
        collectionTypesInterface
    End Sub

    Private Sub cmdCopyCollection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopyCollection.Click
        collectionCopyInterface 
    End Sub

    Private Sub cmdForm2Registry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdForm2Registry.Click
        form2Registry
    End Sub

    Private Sub cmdMkRandomCollections_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMkRandomCollections.Click
        mkRandomCollections
    End Sub

    Private Sub cmdRegistry2Form_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRegistry2Form.Click
        registry2Form(False)
    End Sub

    Private Sub cmdStatusZoom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStatusZoom.Click
        zoomInterface(lstStatus)
    End Sub

    Private Sub cmdStressTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStressTest.Click
        stressTest
    End Sub

    Private Sub cmdString2Collection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdString2Collection.Click
        string2CollectionInterface(txtSerializedCollection.Text)
    End Sub

    Private Sub collectionTest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not initializeCollections Then closer(False)
        Dim booFirstTime As Boolean = True
        Try
            booFirstTime = CBool(GetSetting(Application.ProductName, _
                                            Me.Name, _
                                            "firstTime", _
                                            "True"))
        Catch: End Try       
        If booFirstTime Then
            showAbout
            _OBJwindowsUtilities.string2Listbox(DEFAULT_SERIALIZED_COLLECTIONS, _
                                               lstSerializedCollection)
            With lstSerializedCollection                                               
                .SelectedIndex = CInt(Iif(.Items.Count = 0, -1, 0))  
            End With                                             
            SaveSetting(Application.ProductName, _
                        Me.Name, _
                        "firstTime", _
                        "False")
        End If        
        _OBJwindowsUtilities.string2Listbox(DEFAULT_COLLECTIONS, _
                                            lstCollectionsAvailable)
        registry2Form(booFirstTime)
        Try
            FRMrandomCollectionControl = New randomCollectionControl
        Catch  
            errorHandler("Cannot create random collection control form: " & _
                                       Err.Number & " " & Err.Description, _
                                       "collectionTest_Load", _
                                       "Continuing")
        End Try        
    End Sub

    Private Sub lstCollectionsAvailable_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstCollectionsAvailable.SelectedIndexChanged
        With lstCollectionsAvailable
            cmdCollectionClear.Enabled = .SelectedIndex > -1
            cmdCollection2String.Enabled = cmdCollectionClear.Enabled
            cmdCopyCollection.Enabled =  cmdCollection2String.Enabled
            cmdCollectionCompare.Enabled =  .SelectedItems.Count = 2
        End With
    End Sub

    Private Sub lstSerializedCollection_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        With lstSerializedCollection
            cmdString2Collection.Enabled = .SelectedIndex > -1
            txtSerializedCollection.Text = _OBJutilities.item(CStr(.Items.Item(.SelectedIndex)), _
                                                              1, ":", False) 
        End With
    End Sub

    Private Sub mnuFileExitNoSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileExitNoSave.Click
        closer(False)
    End Sub

    Private Sub mnuFileExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileExit.Click
        closer(True)
    End Sub

    Private Sub mnuAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAbout.Click
        showAbout
    End Sub

    Private Sub mnuToolsRandomCollectionControl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsRandomCollectionControl.Click
        FRMrandomCollectionControl.Show
    End Sub

#End Region ' Form events

#Region " General procedures "

    ' -----------------------------------------------------------------
    ' Add collection, to display and our list
    '
    '
    Private Function addCollection(ByVal colCollection As Collection) As Boolean
        Try
            Redim COLcollections(UBound(COLcollections) + 1) 
            COLcollections(UBound(COLcollections)) = colCollection
            With lstCollectionsAvailable
                Dim strCollectionName As String = collection2Name(colCollection)
                .Items.Add(strCollectionName)
                .SelectedIndex = _OBJwindowsUtilities.searchListBox(lstCollectionsAvailable, _
                                                                    strCollectionName)
            End With                
        Catch  
            errorHandler("Cannot add collection properly to table and/or display: " & _
                                       Err.Number & " " & Err.Description, _
                                       "addCollection", _
                                       "Returning False")
            Return(False)                                       
        End Try        
        Return(True)
    End Function    
    
    ' -----------------------------------------------------------------
    ' Clear the Registry and set defaults
    '
    '
    Private Function clearRegistry As Boolean
        _OBJwindowsUtilities.clearRegistryKey(Application.ProductName, _
                                              Me.Name)
        registry2Form(False)                                              
    End Function    

    ' -----------------------------------------------------------------
    ' Close the application
    '
    '
    Private Sub closer(ByVal booSave As Boolean)
        If booSave Then form2Registry
        If Not (FRMrandomCollectionControl Is Nothing) Then
            FRMrandomCollectionControl.Dispose
        End If            
        Hide
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Collection to name
    '
    '
    Private Function collection2Name(ByVal colCollection As Collection) As String
        Dim strName As String 
        If colCollection.Count = 0 Then 
            strName = "emptyCollection"
        Else            
            strName = _OBJutilities.ellipsis _
                      (_OBJutilities.collection2String(colCollection), 32)
        End If        
        Return(mkCollectionName(strName, nameSequence(strName)))
    End Function    
    
    ' -----------------------------------------------------------------
    ' Test string2Collection
    '
    '
    Private Sub collection2StringInterface(ByVal booReadable As Boolean, _
                                           ByVal booInspect As Boolean)  
        With lstCollectionsAvailable
            If .SelectedIndices.Count > 1 Then
                Select Case MsgBox("Convert each selected collection to a string?", _
                                   MsgBoxStyle.YesNo)
                    Case MsgBoxResult.Yes:
                    Case MsgBoxResult.No: Return
                    Case Else:
                        errorHandler("Unexpected case", _
                                     "collection2StringInterface", _
                                     "Closing application")
                        closer(False)                                                   
                End Select                                   
            End If            
            Dim strCollection As String  
            Dim intIndex1 As Integer
            Dim strConversions As String = "No collections have been converted"
            updateStatusListBox("Converting one or more collections to strings", 1)
            For intIndex1 = 0 To .SelectedIndices.Count - 1
                updateProgress("Converting one or more collections to strings", _
                               "collection", _
                               intIndex1 + 1, _
                               .SelectedIndices.Count)
                Try
                    strCollection = _
                        _OBJutilities.collection2String _
                        (getCollection(.SelectedIndices.Item(intIndex1)), _
                         booReadable:=booReadable, _
                         booInspect:=booInspect AndAlso Not booReadable, _
                         intRecursionDepth:=FRMrandomCollectionControl.MaxDepth)
                Catch
                    errorHandler("Error in collection2String: " & _
                                 Err.Number & " " & Err.Description, _
                                 "collection2StringInterface", _
                                 "Returning to caller")
                    Return                                       
                End Try                                                 
                _OBJwindowsUtilities.addUniqueListItem(lstSerializedCollection, _
                                                       strCollection)      
                If intIndex1 = 0 Then
                    strConversions = "The following collections have been converted:" & _
                                     vbNewline & vbNewline 
                End If                          
                _OBJutilities.append(strConversions, _
                                     vbNewline & vbNewline, _
                                     CStr(lstCollectionsAvailable.Items(.SelectedIndices(intIndex1))) & _
                                     vbNewline & _
                                     "converts to " & _
                                     vbNewline & _
                                     _OBJutilities.enquote(_OBJutilities.ellipsis(strCollection, 256)))                                                                                  
            Next intIndex1            
            updateStatusListBox(-1, "Conversions complete")
            lblProgress.Visible = False
            MsgBox(strConversions)
        End With                                                                                                          
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Test collectionCopy
    '
    '
    Private Sub collectionClearInterface 
        With lstCollectionsAvailable
            If .SelectedIndices.Count > 1 Then
                Select Case MsgBox("Copy each selected collection?", _
                                   MsgBoxStyle.YesNo)
                    Case MsgBoxResult.Yes:
                    Case MsgBoxResult.No: Return
                    Case Else:
                        errorHandler("Unexpected case", _
                                     "collection2StringInterface", _
                                     "Closing application")
                        closer(False)                                                   
                End Select                                   
            End If          
            Dim colNext As Collection  
            Dim intIndex1 As Integer
            Dim strCollection As String  
            updateStatusListBox("Converting one or more collections to strings", 1)
            For intIndex1 = 0 To .SelectedIndices.Count - 1
                updateProgress("Converting collections to strings", _
                               "collection", _
                               intIndex1 + 1, _
                               .SelectedIndices.Count)
                colNext = getCollection(.SelectedIndices.Item(intIndex1))
                If Not (colNext Is Nothing) Then
                    Try
                        Dim colCopy As Collection = _OBJutilities.collectionCopy(colNext)
                        If (colCopy Is Nothing) Then 
                            addCollection(colCopy)
                        End If                                                     
                    Catch  
                        errorHandler("collectionCopy error: " & _
                                     Err.Number & " " & Err.Description, _
                                     "collectionCopyInterface", _
                                     "Returning to caller")
                    End Try        
                End If                    
            Next intIndex1            
            updateStatusListBox(-1, "Complete")
            lblProgress.Visible = False
        End With        
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Interface to collectionCompare
    '
    '
    Private Sub collectionCompareInterface(ByVal colCollection1 As Collection, _
                                           ByVal colCollection2 As Collection) 
        With _OBJutilities
            MsgBox("The collection " & vbNewline & vbNewline & _
                   .ellipsis(.collection2String(colCollection1), 64) & " " & _
                   vbNewline & vbNewline & _
                   "is " & _
                   CStr(Iif(_OBJutilities.collectionCompare(colCollection1, colCollection2), _
                            "", "not")) & " " & _
                   "identifical to " & _                         
                   "The collection " & vbNewline & vbNewline & _
                   .ellipsis(.collection2String(colCollection1), 64))
        End With                       
    End Sub                                            
    
    ' -----------------------------------------------------------------
    ' Test collectionCopy
    '
    '
    Private Sub collectionCopyInterface 
        With lstCollectionsAvailable
            If .SelectedIndices.Count > 1 Then
                Select Case MsgBox("Copy each selected collection?", _
                                   MsgBoxStyle.YesNo)
                    Case MsgBoxResult.Yes:
                    Case MsgBoxResult.No: Return
                    Case Else:
                        errorHandler("Unexpected case", _
                                     "collection2StringInterface", _
                                     "Closing application")
                        closer(False)                                                   
                End Select                                   
            End If          
            Dim colNext As Collection  
            Dim intIndex1 As Integer
            Dim strCollection As String  
            updateStatusListBox("Copying one or more collections", 1)
            For intIndex1 = 0 To .SelectedIndices.Count - 1
                updateProgress("Copying collections", _
                               "collection", _
                               intIndex1 + 1, _
                               .SelectedIndices.Count)
                colNext = getCollection(.SelectedIndices.Item(intIndex1))
                If Not (colNext Is Nothing) Then
                    Try
                        Dim colCopy As Collection = _OBJutilities.collectionCopy(colNext)
                        If (colCopy Is Nothing) Then 
                            addCollection(colCopy)
                        End If                                                     
                    Catch  
                        errorHandler("collectionCopy error: " & _
                                     Err.Number & " " & Err.Description, _
                                     "collectionCopyInterface", _
                                     "Returning to caller")
                    End Try        
                End If                    
            Next intIndex1      
            updateStatusListBox(-1, "Complete")
            lblProgress.Visible = False      
        End With        
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Collection types interface
    '
    '
    Private Sub collectionTypesInterface
        With lstCollectionsAvailable
            If .SelectedIndices.Count > 1 Then
                Select Case MsgBox("Create type collections for each selected collection?", _
                                   MsgBoxStyle.YesNo)
                    Case MsgBoxResult.Yes:
                    Case MsgBoxResult.No: Return
                    Case Else:
                        errorHandler("Unexpected case", _
                                     "collection2StringInterface", _
                                     "Closing application")
                        closer(False)                                                   
                End Select                                   
            End If     
            Dim colNext As Collection
            Dim colTypes As Collection     
            Dim intIndex1 As Integer
            Dim strDisplay As String
            updateStatusListBox("Converting one or more collections to type collections", 1)
            For intIndex1 = 0 To .SelectedIndices.Count - 1
                updateProgress("Converting one or more collections to type collections", _
                               "Collection", _
                               intIndex1 + 1, _
                               .SelectedIndices.Count)
                colNext = getCollection(.SelectedIndices.Item(intIndex1))
                colTypes = _OBJutilities.collectionTypes(colNext)
                addCollection(colTypes)
                With _OBJutilities
                    .append(strDisplay, _
                            vbNewline, _
                            "Collection " & collection2Name(colNext) & _
                            "converts to the type collection " & _
                            .ellipsis(.collection2String(colTypes), 32))
                End With                                     
            Next intIndex1    
            updateStatusListBox(-1, "Complete")
            lblProgress.Visible = False   
            MsgBox(strDisplay)     
        End With            
    End Sub    
        
    ' -----------------------------------------------------------------
    ' Error handler interface
    '
    '
    Private Sub errorHandler(ByVal strMessage As String, _
                             ByVal strProcedure As String, _
                             ByVal strHelp As String)
        _OBJutilities.errorHandler(strMessage, Me.Name, strProcedure, strHelp)
        updateStatusListBox(strMessage)
    End Sub                             

    ' -----------------------------------------------------------------
    ' Form settings to the Registry
    '
    '
    Private Function form2Registry As Boolean
        If Not _OBJwindowsUtilities.listbox2Registry(lstSerializedCollection, _
                                                    strApplication:=Application.ProductName, _
                                                    strSection:=Me.Name) Then
            errorHandler("Cannot save serialized collection list", _
                         "form2Registry", _
                         "Continuing")
            Return(False)                                      
        End If          
        Try
            With chkCollection2StringReadable
                SaveSetting(Application.ProductName, _
                            Me.Name, _
                            .Name, _
                            CStr(.Checked)) 
            End With            
            With chkCollection2StringInspect
                SaveSetting(Application.ProductName, _
                            Me.Name, _
                            .Name, _
                            CStr(.Checked)) 
            End With            
        Catch  
            errorHandler("Cannot save control value in Registry: " & _
                         Err.Number & " " & Err.Description, _
                         "form2Registry", _
                         "Continuing")
            Return(False)                                      
        End Try                                                  
        Return(True)
    End Function    
    
    ' -----------------------------------------------------------------
    ' Get collection given an index to the collection list box
    '
    '
    Private Function getCollection(ByVal intListBoxIndex As Integer) As Collection
        If intListBoxIndex = RANDOM_COLLECTION_INDEX Then
            COLcollections(intListBoxIndex) = mkRandomCollection
        End If
        Return(COLcollections(intListBoxIndex))
    End Function    
    
    ' -----------------------------------------------------------------
    ' Initialize collections
    '
    '
    ' Collection(0) is always an empty collection. Collection(1) is
    ' the documentation example. Collection(2) is Nothing. Collection(3)
    ' is a random collection; note that this collection isn't set up here;
    ' instead, it is set up in the getCollection method. Collection(4) is 
    ' "something", specifically a one-item collection containing a null string.
    '
    ' Here is the documentation example.
    '
    '
    '      1.  Item 1 is the string "Member 1 of 4"
    '      2.  Item 2 is the integer 2
    '      3.  Item 3 is a collection and it contains these members:
    '          3.1 Item 1 is Nothing
    '          3.2 Item 2 is a collection containing one member set to the
    '              Boolean value True
    '      4.  Item 4 is a stringBuilder object
    '
    ' 
    Private Function initializeCollections As Boolean
        Try
            Redim COLcollections(_OBJutilities.items(DEFAULT_COLLECTIONS, vbNewline, False) - 1)
            COLcollections(0) = New Collection
            COLcollections(1) = New Collection
            With COLcollections(1)
                .Add("Member 1 of 4")
                .Add(CInt(2))
                Dim colItem3 As New Collection
                With colItem3
                    .Add(Nothing)
                    Dim colItem32 As New Collection
                    colItem32.Add(True)
                    .Add(colItem32)
                End With           
                .Add(colItem3)
                Dim objStringBuilder As New System.Text.StringBuilder     
                .Add(objStringBuilder)
            End With            
            COLcollections(4) = New Collection
            COLcollections(4).Add("")
            COLcollections(5) = New Collection
            COLcollections(5).Add("A")
        Catch  
            errorHandler("Cannot allocate base test collections: " & _
                         Err.Number & " " & Err.Description, _
                         "initializeCollections", _
                         "Returning False")
            Return(False)                                       
        End Try    
        Return(True)    
    End Function    
    
    ' -----------------------------------------------------------------
    ' Make collection name in the format name [seq]
    '
    '
    Private Function mkCollectionName(ByVal strNameRoot As String, _
                                      ByVal intSequence As Integer) _
            As String
        Return(strNameRoot & _
               CStr(Iif(intSequence = 1, "", " [" & intSequence & "]")))    
    End Function            

    ' -----------------------------------------------------------------
    ' Make one random collection
    '
    '
    Private Overloads Function mkRandomCollection As Collection
        Return(mkRandomCollection(0))
    End Function    
    Private Overloads Function mkRandomCollection(ByVal intRecursionDepth As Integer) _
            As Collection
        Dim colRandom As Collection
        Try
            colRandom = New Collection
        Catch  
            errorHandler("Can't create random collection: " & _
                         Err.Number & " " & Err.Description, _
                         "mkRandomCollection", _
                         "Returning Nothing")
            Return(Nothing)                                       
        End Try
        Dim strEntryTypes() As String
        Try
            strEntryTypes = split(FRMrandomCollectionControl.EntryTypes, " ")
        Catch  
            errorHandler("Can't split entry types: " & _
                         Err.Number & " " & Err.Description, _
                         "mkRandomCollection", _
                         "Returning Nothing")
            Return(Nothing)                                       
        End Try
        Dim intCount As Integer 
        Dim intEntries As Integer = CInt(Rnd * FRMrandomCollectionControl.MaxLength)
        Dim objEntry As Object
        updateStatusListBox("Creating a random collection or sub-collection", 1)
        For intCount = 1 To intEntries 
            updateProgress("Creating a random collection or sub-collection", _
                           "item", _
                           intCount, _
                           intEntries)
            Select Case UCase(strEntryTypes(CInt(Rnd * Ubound(strEntryTypes))))
                Case "BOOLEAN": objEntry = (Rnd > .5)
                Case "BYTE": objEntry = CByte(Rnd * 255)
                Case "SHORT": objEntry = CShort(Math.Max(Math.Max(Rnd * 65536 - 32768, 32767), -32768))  
                Case "INTEGER": objEntry = CInt(Math.Max(Math.Max(Rnd * 2^32 - 2^31, 2^31 - 1), -(2^31)))  
                Case "LONG": objEntry = CLng(Math.Min(Math.Max(Rnd * 2^63 - 2^62, 2^63 - 1), -(2^62)))  
                Case "SINGLE": objEntry = CSng(Math.Round(Rnd, CInt(Rnd * 10)))
                Case "DOUBLE": objEntry = CDbl(Math.Round(Rnd, CInt(Rnd * 10)))
                Case "STRING": objEntry = randomString
                Case "COLLECTION":
                    If intRecursionDepth >= FRMrandomCollectionControl.MaxDepth Then
                        objEntry = "Can't provide collection, recursion depth exceeded"
                    Else     
                        objEntry = mkRandomCollection(intRecursionDepth + 1)                   
                    End If                        
                Case "STRINGBUILDER":
                    Try
                        objEntry = New System.Text.StringBuilder
                    Catch  
                        Dim strMessage As String = "Could not create stringBuilder: " & _
                                                   Err.Number & " " & Err.Description
                        errorHandler(strMessage, "mkRandomCollection", _
                                     "String replaces object")                                                   
                        objEntry = strMessage
                    End Try   
                Case Else:
                    errorHandler("Unexpected case", "mkRandomCollection", "Closing")
                    closer(False)                                                                                                        
            End Select            
            Try
                colRandom.Add(objEntry)
            Catch  
                errorHandler("Cannot extend random collection: " & _
                             Err.Number & " " & Err.Description, _
                             "mkRandomCollection", _
                             "Will return an incomplete collection")
                Exit For                             
            End Try            
        Next intCount        
        updateStatusListBox(-1, "Complete")
        lblProgress.Visible = False
        Return(colRandom)
    End Function    
    
    ' -----------------------------------------------------------------
    ' Create 0..32 random collections
    Private Sub mkRandomCollections  
        Dim intCount As Integer = CInt(Rnd * 32)
        Dim intIndex1 As Integer
        updateStatusListBox("Creating some random collections", 1)
        For intIndex1 = 0 To intCount
            updateProgress("Creating random collection", _
                           "Collection", _
                           intIndex1 + 1, _
                           intCount)
            addCollection(mkRandomCollection)
        Next intIndex1        
        lblProgress.Visible = False
        updateStatusListBox(-1, "Complete")
    End Sub     

    ' -----------------------------------------------------------------
    ' Determine name sequence number
    '
    '
    Private Function nameSequence(ByVal strName As String) As Integer
        Dim intIndex1 As Integer
        Dim intMaxSequence As Integer
        Dim intSequence As Integer
        Dim strNextName As String
        With lstCollectionsAvailable
            For intIndex1 = 0 To .Items.Count - 1
                parseCollectionName(CStr(.Items(intIndex1)), strNextName, intSequence)  
                If intSequence = 0 Then intSequence = 1
                If strName = strNextName AndAlso intSequence > intMaxSequence Then 
                    intMaxSequence = intSequence                                    
                End If                    
            Next intIndex1        
        End With
        Return(intMaxSequence + 1)
    End Function    
        
    ' -----------------------------------------------------------------
    ' Parse collection name format stringCollection [seq]
    '
    '
    Private Sub parseCollectionName(ByVal strName As String, _
                                    ByRef strNameRoot As String, _
                                    ByRef intSequence As Integer) 
        strNameRoot = strName
        intSequence = 0
        Dim intIndex1 As Integer = Instr(strName, "[")  
        If intIndex1 = 0 Then Return
        If Mid(strName, Len(strName)) <> "]" Then Return
        intIndex1 += 1
        Try
            intSequence = CInt(Mid(strName, intIndex1, Len(strName) - intIndex1))
        Catch: End Try    
        strNameRoot = Trim(Mid(strName, 1, intIndex1 - 2))    
    End Sub                            
    
    ' -----------------------------------------------------------------
    ' Return a random string
    '
    '
    Private Function randomString As String
        Dim intLength As Integer
        Dim strRandom As String
        For intLength = 0 To CInt(Rnd * 32)
            strRandom &= Mid("abcdef", CInt(Rnd * 5) + 1, 1)
        Next intLength        
        Return(strRandom)
    End Function    

    ' -----------------------------------------------------------------
    ' Registry settings to the form
    '
    '
    Private Function registry2Form(ByVal booFirstTime As Boolean) As Boolean
        If Not booFirstTime Then
            If Not _OBJwindowsUtilities.registry2ListBox(lstSerializedCollection, _
                                                        strApplication:=Application.ProductName, _
                                                        strSection:=Me.Name) Then
                errorHandler("Cannot obtain serialized collection list", _
                             "registry2Form", _
                             "Continuing")
                Return(False)                                      
            End If          
            With lstSerializedCollection                                               
                .SelectedIndex = CInt(Iif(.Items.Count = 0, -1, 0))  
            End With                                             
        End If
        Try
            With chkCollection2StringReadable
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            Me.Name, _
                                            .Name, _
                                            CStr(.Checked)))
            End With            
            With chkCollection2StringInspect
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            Me.Name, _
                                            .Name, _
                                            CStr(.Checked)))
            End With            
        Catch  
            errorHandler("Cannot obtain control value from Registry: " & _
                         Err.Number & " " & Err.Description, _
                         "registry2Form", _
                         "Continuing")
            Return(False)                                      
        End Try                                                  
        Return(True)
    End Function    
    
    ' -----------------------------------------------------------------
    ' Show about information
    '
    '
    Private Sub showAbout
        MsgBox(Me.Name & vbNewline & vbNewline & ABOUT_INFO)
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Stress test
    '
    '
    Private Sub stressTest
        With FRMrandomCollectionControl
            Select Case MsgBox("Conduct " & .StressTestCount & " stress test(s)?", _
                               MsgBoxStyle.YesNo)
                Case MsgBoxResult.Yes:
                Case MsgBoxResult.No: Return
                Case Else:
                    errorHandler("Unexpected case", "stressTest", "Closing")
                    closer(False)                                               
            End Select        
            updateStatusListBox("Conducting the stress test", 1)   
            mkRandomCollections
            Dim intIndex1 As Integer 
            Dim intSelectionSave() As Integer
            With lstCollectionsAvailable
                Try
                    Redim intSelectionSave(.SelectedIndices.Count)
                Catch  
                    errorHandler("Cannot allocate selection save array", _
                                 "stressTest", _
                                 "Selections will be lost")
                End Try            
                For intIndex1 = 0 To .SelectedIndices.Count
                    intSelectionSave(intIndex1) = .SelectedIndices(intIndex1)
                Next intIndex1                 
                lstCollectionsAvailable.ClearSelected  
            End With       
            Dim booReadable As Boolean
            Dim colCompared(1) As Collection
            Dim intIndex2 As Integer      
            Dim strCollectionName As String    
            For intIndex1 = 1 To .StressTestCount
                updateProgress("Conducting the stress test", "test", intIndex1, .StressTestCount)
                Try
                    With lstCollectionsAvailable
                        .SelectedIndex = CInt(Rnd * .Items.Count) - 1
                        strCollectionName = collection2Name(getCollection(.SelectedIndex))
                        Select Case CByte(Rnd * 4)
                            Case 0: 
                                updateStatusListBox("Clearing the collection " & _
                                                    strCollectionName)
                                collectionClearInterface 
                            Case 1: 
                                booReadable = (Rnd < FRMrandomCollectionControl.ReadabilityProbability)
                                updateStatusListBox("Converting the collection " & _
                                                    strCollectionName & _
                                                    "to a " & _
                                                    CStr(Iif(booReadable, "readable", "serialized")) & " " & _
                                                    "string")
                                collection2StringInterface(booReadable:=booReadable, _
                                                           booInspect:=Not booReadable) 
                            Case 2: 
                                collectionCopyInterface
                                updateStatusListBox("Duplicating the collection " & _
                                                    strCollectionName)
                            Case 3: 
                                collectionTypesInterface 
                                updateStatusListBox("Creating the type collection using " & _
                                                    strCollectionName)
                            Case 4: 
                                If .SelectedIndices.Count < 2 Then
                                    .ClearSelected
                                    .SelectedIndex = CInt(Rnd * (.Items.Count - 1))
                                    colCompared(0) = getCollection(.SelectedIndex)
                                    strCollectionName = collection2Name(colCompared(0))
                                    intIndex2 = CInt(Rnd * (.Items.Count - 2))
                                    .SelectedIndex = CInt(Iif(intIndex2 >= .SelectedIndex, _
                                                            intIndex2 + 1, intIndex2))
                                    colCompared(1) = getCollection(.SelectedIndex)                                                                                                                                        
                                End If       
                                With .SelectedIndices                     
                                    updateStatusListBox("Comparing " & _
                                                        strCollectionName & " with " & _
                                                        collection2Name(colCompared(1)))
                                    collectionCompareInterface(colCompared(0), _
                                                               colCompared(1))
                                End With
                            Case Else:    
                                errorHandler("Unexpected case", _
                                             "stressTest", _
                                             "Terminating the stress test")
                        End Select     
                    End With      
                Catch
                    errorHandler("Error in test: " & Err.Number & " " & Err.Description, _
                                 "stressTest", _
                                 "Terminating the stress test")
                End try                      
            Next intIndex1            
            updateStatusListBox("Completed", 1)   
            lblProgress.Visible = False
            With lstCollectionsAvailable
                .Focus
                For intIndex1 = 0 To UBound(intSelectionSave)
                    .SelectedIndex = intSelectionSave(intIndex1)
                    .Refresh
                Next intIndex1                
            End With                 
        End With        
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Test string2Collection
    '
    '
    Private Sub string2CollectionInterface(ByVal strSerialized As String) 
        Dim colNew As Collection  
        Try
            colNew = _OBJutilities.string2Collection(strSerialized, _
                                                     booInspect:=chkCollection2StringInspect.Checked)
        Catch
            errorHandler("Error in string2Collection: " & _
                         Err.Number & " " & Err.Description, _
                         "string2CollectionInterface", _
                         "Returning to caller")
            Return                                       
        End Try                                                 
        addCollection(colNew)                                                                   
    End Sub    
    
    ' ------------------------------------------------------------------
    ' Update progress
    '
    '
    Private Sub updateProgress(ByVal strActivity As String, _
                               ByVal strEntity As String, _
                               ByVal intEntityNumber As Integer, _
                               ByVal intEntityCount As Integer)
        If intEntityCount = 0 Then Return                               
        With lblProgress
            Dim strMessage As String = strActivity & " " & _
                                       "at " & strEntity & " " & intEntityNumber &  " " & _
                                       "of " & intEntityCount & ": " & _
                                       Math.Round(intEntityNumber/intEntityCount * 100, 2) & _
                                       "% complete" 
            updateStatusListBox(strMessage)                                       
            .Width = CInt(_OBJutilities.histogram(CDbl(intEntityNumber), _
                                                  dblRangeMax:=CDbl(lstStatus.Width), _
                                                  dblValueMax:=CDbl(intEntityCount)))                                       
            .Text = strMessage
            .Visible = True
            .Refresh
        End With        
    End Sub                                   
    ' ------------------------------------------------------------------
    ' Update nested status reporting
    '
    '
    ' --- Don't change the level at all
    Private Overloads Sub updateStatusListBox(ByVal strMessage As String)
        updateStatusListBox(0, strMessage, 0)
    End Sub    
    ' --- Change the level before appending the report
    Private Overloads Sub updateStatusListBox(ByVal intLevelChangeBefore As Integer, _
                                              ByVal strMessage As String)
        updateStatusListBox(intLevelChangeBefore, strMessage, 0)
    End Sub    
    ' --- Change the level after appending the report
    Private Overloads Sub updateStatusListBox(ByVal strMessage As String, _
                                              ByVal intLevelChangeAfter As Integer)
        updateStatusListBox(0, strMessage, intLevelChangeAfter)
    End Sub    
    ' --- Change the level before and after appending the report
    Private Overloads Sub updateStatusListBox(ByVal intLevelChangeBefore As Integer, _
                                              ByVal strMessage As String, _
                                              ByVal intLevelChangeAfter As Integer)
        Static intLevel As Integer
        intLevel += intLevelChangeBefore
        With lstStatus
            .Focus
            Try
                .Items.Add(_OBJutilities.copies(" ", intLevel * 5) & _
                           Now & " " & _
                           strMessage)
            Catch  
                errorHandler("Cannot append to status list box: " & _
                             Err.Number & " " & Err.Description, _
                             "updateStatusListBox", _
                             "Continuing")
            End Try            
            .SelectedIndex = .Items.Count - 1
            .Refresh
        End With        
        intLevel += intLevelChangeAfter
    End Sub    

    ' -----------------------------------------------------------------
    ' Zoom the control
    '
    '
    Private Sub zoomInterface(ByVal ctlZoomed As Control)
        Dim ctlZoom As zoom.zoom 
        Try
            ctlZoom = New zoom.zoom(ctlZoomed, CDbl(3), CDbl(4))
        Catch  
            errorHandler("Cannot create zoom control: " & _
                         Err.Number & " " & Err.Description, _
                         "zoomInterface", _
                         "Will not be able to show zoomed control info")
            Return                                       
        End Try   
        With ctlZoom
            .showZoom
            .dispose
        End With             
    End Sub    
    
#End Region ' General procedures

End Class
