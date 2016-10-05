VERSION 5.00
Begin VB.Form frmShowFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "showFile"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox txtShowFile 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
   End
End
Attribute VB_Name = "frmShowFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************************
' *                                                                   *
' * showFile                                                          *
' *                                                                   *
' *                                                                   *
' * This is a very simple text editor so that straightforward changes *
' * to BNF can be made without leaving bnfAnalyzer. This form as an   *
' * object supports the following properties.                         *
' *                                                                   *
' *                                                                   *
' *      *  The read-write FileText property returns and may be set to*
' *         the contents of the file to be changed.                   *
' *                                                                   *
' *      *  The read-only WasClosed property returns True if the Close*
' *         button is used to close the file, False otherwise.        *
' *                                                                   *
' *                                                                   *
' *********************************************************************

Private BOOwasClosed As Boolean
Friend Property Get FileText() As String
    FileText = txtShowFile.Text
End Property
Friend Property Let FileText(ByVal strNewValue As String)
    With txtShowFile
        .Text = FileText
        .Refresh
    End With
End Property
Friend Property Get WasClosed() As Boolean
    WasClosed = BOOwasClosed
End Property
Private Sub cmdClose_Click()
    BOOwasClosed = True
End Sub

Private Sub Form_Activate()
    BOOwasClosed = False
End Sub
