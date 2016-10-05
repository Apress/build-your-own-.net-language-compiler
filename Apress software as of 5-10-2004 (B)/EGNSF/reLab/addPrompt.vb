Option Strict

' *********************************************************************
' *                                                                   *
' * addPrompt                                                         *
' *                                                                   *
' *                                                                   *
' * This form handles the way in which new regular expressions and new*
' * test strings are added, and it allows the user to create a persis-*
' * tent policy.  It exposes the following properties and methods.    *
' *                                                                   *
' *                                                                   *
' *      *  Add: this read-only property returns True when the user   *
' *         has selected, through a radio button, the choice to add   *
' *         the item, False otherwise.                                *
' *                                                                   *
' *      *  Cancel: this read-only property returns True when the     *
' *         form is closed by any means other than the cmdOK button,  *
' *         False otherwise.                                          *
' *                                                                   *
' *      *  showForCustomization: this method displays this form with *
' *         the controls to add a regular expression or string and to *
' *         change its description greyed-out.                        *
' *                                                                   *
' *      *  Description: this read-write property returns and may be  *
' *         set to the description of the specific item's purpose.    *
' *                                                                   *
' *      *  ItemDescription: this read-write property returns and may *
' *         be set to the item description which in practice will be  *
' *         "regular expression" or "test string".                    *
' *                                                                   *
' *      *  Policy: this read-only property returns the persistent    *
' *         policy as one of the following enumerators.               *
' *                                                                   *
' *         + ENUaddPromptPolicy.alwaysShow: the prompt always appears*
' *                                                                   *
' *         + ENUaddPromptPolicy.neverAdd: the prompt does not appear,*
' *           and items are not added                                 *
' *                                                                   *
' *         + ENUaddPromptPolicy.alwaysAdd: the prompt does not       *
' *           appear, but items are always added                      *
' *                                                                   *
' *      *  Value: this read-write property returns and may be set to *
' *         the item value                                            *
' *                                                                   *
' *                                                                   *
' *********************************************************************

Public Class addPrompt
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
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents gbxPromptRule As System.Windows.Forms.GroupBox
    Friend WithEvents radPromptRuleAlwaysShow As System.Windows.Forms.RadioButton
    Friend WithEvents radPromptRuleNeverAdd As System.Windows.Forms.RadioButton
    Friend WithEvents radPromptRuleAlwaysAdd As System.Windows.Forms.RadioButton
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents radAdd As System.Windows.Forms.RadioButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.radAdd = New System.Windows.Forms.RadioButton
Me.txtDescription = New System.Windows.Forms.TextBox
Me.cmdOK = New System.Windows.Forms.Button
Me.gbxPromptRule = New System.Windows.Forms.GroupBox
Me.radPromptRuleAlwaysAdd = New System.Windows.Forms.RadioButton
Me.radPromptRuleNeverAdd = New System.Windows.Forms.RadioButton
Me.radPromptRuleAlwaysShow = New System.Windows.Forms.RadioButton
Me.cmdCancel = New System.Windows.Forms.Button
Me.gbxPromptRule.SuspendLayout()
Me.SuspendLayout()
'
'radAdd
'
Me.radAdd.Checked = True
Me.radAdd.Location = New System.Drawing.Point(8, 8)
Me.radAdd.Name = "radAdd"
Me.radAdd.Size = New System.Drawing.Size(352, 40)
Me.radAdd.TabIndex = 0
Me.radAdd.TabStop = True
Me.radAdd.Text = "Add the %DESCRIPTION %VALUE with the following description"
'
'txtDescription
'
Me.txtDescription.Location = New System.Drawing.Point(8, 48)
Me.txtDescription.Name = "txtDescription"
Me.txtDescription.Size = New System.Drawing.Size(352, 20)
Me.txtDescription.TabIndex = 1
Me.txtDescription.Text = ""
'
'cmdOK
'
Me.cmdOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdOK.Location = New System.Drawing.Point(264, 76)
Me.cmdOK.Name = "cmdOK"
Me.cmdOK.Size = New System.Drawing.Size(96, 84)
Me.cmdOK.TabIndex = 2
Me.cmdOK.Text = "OK"
'
'gbxPromptRule
'
Me.gbxPromptRule.Controls.Add(Me.radPromptRuleAlwaysAdd)
Me.gbxPromptRule.Controls.Add(Me.radPromptRuleNeverAdd)
Me.gbxPromptRule.Controls.Add(Me.radPromptRuleAlwaysShow)
Me.gbxPromptRule.Location = New System.Drawing.Point(8, 80)
Me.gbxPromptRule.Name = "gbxPromptRule"
Me.gbxPromptRule.Size = New System.Drawing.Size(248, 120)
Me.gbxPromptRule.TabIndex = 3
Me.gbxPromptRule.TabStop = False
'
'radPromptRuleAlwaysAdd
'
Me.radPromptRuleAlwaysAdd.Location = New System.Drawing.Point(8, 80)
Me.radPromptRuleAlwaysAdd.Name = "radPromptRuleAlwaysAdd"
Me.radPromptRuleAlwaysAdd.Size = New System.Drawing.Size(232, 32)
Me.radPromptRuleAlwaysAdd.TabIndex = 2
Me.radPromptRuleAlwaysAdd.Text = "Never show this prompt: always add with default description"
'
'radPromptRuleNeverAdd
'
Me.radPromptRuleNeverAdd.Location = New System.Drawing.Point(8, 48)
Me.radPromptRuleNeverAdd.Name = "radPromptRuleNeverAdd"
Me.radPromptRuleNeverAdd.Size = New System.Drawing.Size(232, 16)
Me.radPromptRuleNeverAdd.TabIndex = 1
Me.radPromptRuleNeverAdd.Text = "Never show this prompt: never add"
'
'radPromptRuleAlwaysShow
'
Me.radPromptRuleAlwaysShow.Checked = True
Me.radPromptRuleAlwaysShow.Location = New System.Drawing.Point(8, 16)
Me.radPromptRuleAlwaysShow.Name = "radPromptRuleAlwaysShow"
Me.radPromptRuleAlwaysShow.Size = New System.Drawing.Size(232, 16)
Me.radPromptRuleAlwaysShow.TabIndex = 0
Me.radPromptRuleAlwaysShow.TabStop = True
Me.radPromptRuleAlwaysShow.Text = "Always show this prompt"
'
'cmdCancel
'
Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdCancel.Location = New System.Drawing.Point(264, 168)
Me.cmdCancel.Name = "cmdCancel"
Me.cmdCancel.Size = New System.Drawing.Size(96, 32)
Me.cmdCancel.TabIndex = 4
Me.cmdCancel.Text = "Cancel"
'
'addPrompt
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(366, 211)
Me.ControlBox = False
Me.Controls.Add(Me.cmdCancel)
Me.Controls.Add(Me.gbxPromptRule)
Me.Controls.Add(Me.cmdOK)
Me.Controls.Add(Me.txtDescription)
Me.Controls.Add(Me.radAdd)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
Me.Name = "addPrompt"
Me.Text = "addPrompt"
Me.gbxPromptRule.ResumeLayout(False)
Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Form data "

    ' ***** Policies *****
    Friend Enum ENUaddPromptPolicy
        alwaysShow
        neverAdd
        alwaysAdd
    End Enum    

    ' ***** Shared utilities *****
    Private _OBJutilities As utilities.utilities

    ' ***** Object state *****
    Private booCancel As Boolean
    Private STRitemDescription As String
    Private STRvalue As String
    
#End Region

#Region " Friend properties and methods "

    Friend ReadOnly Property Add As Boolean
        Get
            Return radAdd.Checked
        End Get        
    End Property    
    
    Friend ReadOnly Property Cancel As Boolean
        Get
            Return booCancel
        End Get        
    End Property    
    
    Friend Property Description As String
        Get
            Return txtDescription.Text
        End Get        
        Set(ByVal strNewValue As String)
            txtDescription.Text = strNewValue
        End Set        
    End Property    
    
    Friend Property ItemDescription As String
        Get
            Return(STRitemDescription)
        End Get        
        Set(ByVal strNewValue As String)
            STRitemDescription = strNewValue
            registry2Form
            editAddText(STRitemDescription, STRvalue)
            Me.Text = "Add " & STRitemDescription
        End Set        
    End Property    
    
    Friend ReadOnly Property Policy As ENUaddPromptPolicy
        Get
            If radPromptRuleAlwaysAdd.Checked Then Return(ENUaddPromptPolicy.alwaysAdd)
            If radPromptRuleAlwaysShow.Checked Then Return(ENUaddPromptPolicy.alwaysShow)
            If radPromptRuleNeverAdd.Checked Then Return(ENUaddPromptPolicy.neverAdd)
            radPromptRuleAlwaysAdd.Checked = True
            Return(Me.Policy)
        End Get        
    End Property    
    
    Friend Sub showForCustomization
        Dim booEnabled(1) As Boolean
        booEnabled(0) = radAdd.Enabled: booEnabled(1) = txtDescription.Enabled
        radAdd.Enabled = False: txtDescription.Enabled = False
        ShowDialog
        radAdd.Enabled = booEnabled(0): txtDescription.Enabled = booEnabled(1)
    End Sub    
    
    Friend Property Value As String
        Get
            Return(STRvalue)
        End Get        
        Set(ByVal strNewValue As String)
            STRvalue = strNewValue
            editAddText(STRitemDescription, STRvalue)
        End Set        
    End Property    
        
#End Region

#Region " Form events "

    Private Sub addPrompt_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        booCancel = True
    End Sub

    Private Sub addPrompt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        registry2Form
    End Sub
 
    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        registry2Form
        Hide
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        booCancel = False
        form2Registry
        Hide
    End Sub
    
#End Region ' " Form events "

#Region " General procedures "

    ' -----------------------------------------------------------------
    ' Replace a keyword in the add text
    '
    '
    Private Sub editAddText(ByVal strDescription As String, _
                            ByVal strValue As String)
        With radAdd
            If (.Tag Is Nothing) Then .Tag = .Text
            .Text = Replace(Replace(CStr(.Tag), "%VALUE", strValue), "%DESCRIPTION", strDescription)
            .Refresh
        End With        
    End Sub
    
    ' ------------------------------------------------------------------
    ' Save form settings in the Registry
    '
    '
    Private Function form2Registry As Boolean
        Try
            Dim strSectionName As String = mkSectionName            
            With radAdd
                SaveSetting(Application.ProductName, _
                            strSectionName, _
                            .Name, _
                            CStr(.Checked)) 
            End With   
            With radPromptRuleAlwaysAdd
                SaveSetting(Application.ProductName, _
                            strSectionName, _
                            .Name, _
                            CStr(.Checked)) 
            End With   
            With radPromptRuleAlwaysShow
                SaveSetting(Application.ProductName, _
                            strSectionName, _
                            .Name, _
                            CStr(.Checked)) 
            End With   
            With radPromptRuleNeverAdd
                SaveSetting(Application.ProductName, _
                            strSectionName, _
                            .Name, _
                            CStr(.Checked)) 
            End With   
        Catch ex As Exception
            _OBJutilities.errorHandler("Can't save Registry settings", _
                                       Me.Name, _
                                       "registry2Form", _
                                       "Continuing, without a complete save")
            Return(False)                                       
        End Try        
        Return(True)
    End Function    
    
    ' ------------------------------------------------------------------
    ' Return the section name
    '
    '
    Private Function mkSectionName As String
        Dim strSectionName As String = Me.Name
        If Me.ItemDescription <> "" Then
            strSectionName &= "_" & Replace(Me.ItemDescription, " ", "_")
        End If 
        Return(strSectionName)           
    End Function    
    
    ' ------------------------------------------------------------------
    ' Get form settings from the Registry
    '
    '
    Private Function registry2Form As Boolean
        Try
            Dim strSectionName As String = mkSectionName            
            With radAdd
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            strSectionName, _
                                            .Name, _
                                            CStr(.Checked)))
            End With   
            With radPromptRuleAlwaysAdd
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            strSectionName, _
                                            .Name, _
                                            CStr(.Checked)))
            End With   
            With radPromptRuleAlwaysShow
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            strSectionName, _
                                            .Name, _
                                            CStr(.Checked)))
            End With   
            With radPromptRuleNeverAdd
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            strSectionName, _
                                            .Name, _
                                            CStr(.Checked)))
            End With   
        Catch ex As Exception
            _OBJutilities.errorHandler("Can't get Registry settings", _
                                       Me.Name, _
                                       "registry2Form", _
                                       "Continuing, with incomplete settings")
            Return(False)                                       
        End Try        
        Return(True)
    End Function    
    
#End Region ' " General procedures "

End Class
