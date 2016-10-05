VERSION 5.00
Begin VB.Form frmRRdiagram 
   BorderStyle     =   1  'Fixed Single
   Caption         =   """Railroad"" Diagram"
   ClientHeight    =   6060
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   8820
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblBuildingBlock 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1332
   End
End
Attribute VB_Name = "frmRRdiagram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************************
' *                                                                   *
' * frmRRdiagram     Create railroad diagram                          *
' *                                                                   *
' *                                                                   *
' * A "railroad" diagram shows Backus-Naur Form syntax as a graph such*
' * that a well-formed formula will have a path through the diagram.  *
' * This form shows the railroad diagram for the current parse when it*
' * is activated. This form exposes the write-only ParseTree property *
' * which must be set to the parse tree collection before the form is *
' * activated.                                                        *
' *                                                                   *
' *********************************************************************

' ***** Utilities *****
Private OBJutilities As New clsUtilities

' ***** Must be set to the parse tree *****
Private COLparseTree As Collection

' ***** Working rectangle *****
Private Type TYPrectangle
    lngLeft As Long
    lngTop As Long
    lngWidth As Long
    lngHeight As Long
End Type

' ***** Constants *****
Private Const CONNECTORSPACE = 0.15     ' Percent of width occupied by leading or
                                        ' ending connector
Private Const GRIDSIZE = 120            ' Form's grid
Friend Property Set ParseTree(ByVal colNewValue As Collection)
    Set COLparseTree = colNewValue
End Property
Private Sub Form_Activate()
    If (COLparseTree Is Nothing) Then
        OBJutilities.errorHandler 0, _
                                  "Parse tree not set", _
                                  "Form_Activate", _
                                  Me.Name
        Exit Sub
    End If
    parseTree2RRdiagram COLparseTree
End Sub
Private Sub parseTree2RRdiagram(ByRef colTree As Collection)
    '
    ' Converts the parse collection to a railroad diagram
    '
    '
    ' Unfortunately this code rather parallels parseTree2RRDiagram as well as
    ' parseTree2Rules in the flagship form. This is a misfortune since
    ' if the structure of the parse tree changes all of these routines must change.
    '
    '
    Dim colHandle(1 To 2)
    Dim lngIndex1 As Long
    Dim lngIndex2 As Long
    Dim usrRectangle As TYPrectangle
    mkRectangle usrRectangle, _
                GRIDSIZE, _
                GRIDSIZE, _
                Width - GRIDSIZE * 2, _
                ScaleHeight - GRIDSIZE * 2
    With COLparseTree
        Set colHandle(1) = .item(1)
        With colHandle(1)
            lngIndex2 = 1
            For lngIndex1 = 1 To .Count
                Set colHandle(2) = .item(lngIndex1)
                With colHandle(2)
                    If frmBNFanalyzer.nonTerminal2NodeIndex(colTree, .item(1)) <> 0 Then
                        node2diagram colTree, _
                                     .item(1), _
                                     frmBNFanalyzer.nonTerminal2NodeIndex _
                                     (colTree, .item(1)), _
                                     usrRectangle
                        lngIndex2 = lngIndex2 + 1
                    End If
                End With
            Next lngIndex1
        End With
    End With
End Sub
Private Sub node2diagram(ByRef colTree As Collection, _
                         ByVal lngNodeIndex As Long, _
                         ByVal strNonterminal As String, _
                         ByRef usrRectangle As TYPrectangle)
    '
    ' Node to diagram
    '
    '
    Dim colNodeHandle As Collection
    Dim colSubnodeHandle As Collection
    Dim lblEntryConnector As Label
    Dim lblExitConnector As Label
    Dim lngConnectorLength As Long
    Dim lngIndex1 As Long
    Dim usrInsideRectangle As TYPrectangle
    With colTree
        If lngNodeIndex = 0 Then Exit Sub
        Set colNodeHandle = .item(lngNodeIndex)
        With colNodeHandle
            ' --- Draw lead and end line
            lblEntryConnector = drawLeadLine(usrRectangle)
            lblExitConnector = drawEndLine(usrRectangle)
            ' --- Calculate the inside rectangle containing the rule
            With usrRectangle
                lngConnectorLength = .lngWidth * CONNECTORSPACE
                mkRectangle usrInsideRectangle, _
                            .lngLeft + lngConnectorLength, _
                            .lngTop, _
                            .lngWidth - lngConnectorLength * 2, _
                            .lngHeight \ 1
            End With
            If (TypeOf .item(3) Is Collection) Then
                ' Operator and one or two operands found
                Set colSubnodeHandle = .item(3)
                Select Case colSubnodeHandle.item(1)
                    Case OPERATOR_ALTERNATION:
                         drawAlternation colTree, _
                                         colSubnodeHandle, _
                                         usrInsideRectangle, _
                                         lblEntryConnector, _
                                         lblExitConnector
                    Case OPERATOR_MRE_ONETRIP:
                         drawOneTripRepeat colTree, _
                                           colSubnodeHandle, _
                                           usrInsideRectangle, _
                                           lblEntryConnector, _
                                           lblExitConnector
                    Case OPERATOR_OPTIONALSEQUENCE:
                         drawOptionalSequence colTree, _
                                              colSubnodeHandle, _
                                              usrInsideRectangle, _
                                              lblEntryConnector, _
                                              lblExitConnector
                    Case OPERATOR_PRODUCTION:
                         drawProduction colTree, colSubnodeHandle, usrRectangle
                    Case OPERATOR_MRE_ZEROTRIP:
                         drawZeroTripRepeat colTree, _
                                            colSubnodeHandle, _
                                            usrInsideRectangle, _
                                            lblEntryConnector, _
                                            lblExitConnector
                    Case OPERATOR_SEQUENCE:
                         drawSequence colTree, colSubnodeHandle, usrInsideRectangle
                    Case Else:
                        OBJutilities.errorHandler 0, _
                                                  "Internal programming error: " & _
                                                  "unsupported operator", _
                                                  "parseTree2RRDiagram_node2Rules", _
                                                  Me.Name
                          Exit Sub
                End Select
            Else
                ' This is a grammar symbol
                lngIndex1 = .item(3)
                If lngIndex1 > 0 Then
                    drawGS frmBNFanalyzer.index2Name(colTree, lngIndex1, True), _
                           usrInsideRectangle, _
                           True
                ElseIf lngIndex1 < 0 Then
                    drawGS frmBNFanalyzer.index2Name(colTree, lngIndex1, False), _
                           usrInsideRectangle, _
                           True
                Else
                    drawError "undefined syntax", _
                              usrInsideRectangle
                End If
            End If
        End With
    End With
End Sub
Private Sub mkRectangle(ByRef usrRectangle As TYPrectangle, _
                        ByVal lngLeft As Long, _
                        ByVal lngTop As Long, _
                        ByVal lngWidth As Long, _
                        ByVal lngHeight As Long)
    '
    ' Creates the rectangle data structure
    '
    '
    With usrRectangle
        .lngLeft = lngLeft
        .lngTop = lngTop
        .lngWidth = lngWidth
        .lngHeight = lngHeight
    End With
End Sub
Private Sub drawAlternation(ByRef colTree As Collection, _
                            ByRef colSubnodeHandle As Collection, _
                            ByRef usrRectangle As TYPrectangle, _
                            ByRef lblEntryConnector As Label, _
                            ByRef lblExitConnector As Label)
    '
    ' Draw the alternation of two possible rules
    '
    '
    With colSubnodeHandle
        ' --- Draw alternative one above alternative two
        node2diagram colTree, CLng(.item(2)), "", usrRectangle
        ' --- Move below alternative one
        moveRectangle usrRectangle, _
                      usrRectangle.lngLeft, _
                      rectangleBottom(usrRectangle)
        ' --- Draw alternative two below alternative one
        node2diagram colTree, CLng(.item(2)), "", usrRectangle
        ' --- Connect the entering line with alternative two
        drawConnectorFromLbl lblEntryConnector, _
                             usrRectangle.lngLeft, _
                             usrRectangle.lngTop
        drawConnectorFromLbl lblExitConnector, _
                             rectangleRight(usrRectangle), _
                             usrRectangle.lngTop
    End With
End Sub
Private Sub drawError(ByVal strMessage As String, _
                      ByRef usrRectangle As TYPrectangle)
                      
End Sub
Private Sub drawGS(ByVal strGS As String, _
                   ByRef usrRectangle As TYPrectangle, _
                   ByVal booNT As Boolean)
    '
    ' Draw the grammar symbol
    '
    '
    Dim lngBackColor As Long
    Dim lngForeColor As Long
    Dim usrLblRectangle As TYPrectangle
    If booNT Then
        lngBackColor = vbGreen: lngForeColor = vbBlack
    Else
        lngBackColor = vbRed: lngForeColor = vbWhite
    End If
    drawLabel usrLblRectangle, _
              strGS, _
              lngBackColor, _
              lngForeColor
End Sub
Private Sub drawOneTripRepeat(ByRef colTree As Collection, _
                              ByRef colSubnodeHandle As Collection, _
                              ByRef usrRectangle As TYPrectangle, _
                              ByRef lblEntryConnector As Label, _
                              ByRef lblExitConnector As Label)
    '
    ' Draw one-trip repetition of a rule
    '
    '
    Dim lblSmallConnector As Label
    Dim lngWidth As Long
    Dim usrRectangle2 As TYPrectangle
    With usrRectangle
        ' --- Draw the rule in 40% of the horizontal space available
        lngWidth = .lngWidth
        .lngWidth = lngWidth * 0.4
        node2diagram colTree, CLng(colSubnodeHandle.item(2)), "", usrRectangle
        ' --- Move to the right of the rule as drawn on the form
        mkRectangle usrRectangle2, _
                    rectangleRight(usrRectangle) + lngWidth * 0.2, _
                    .lngTop, _
                    .lngWidth, _
                    .lngHeight
        ' --- Draw and remember a small connector
        lblSmallConnector = drawLine(rectangleRight(usrRectangle), _
                                     .lngTop, _
                                     usrRectangle2.lngLeft, _
                                     .lngTop)
        ' --- Draw a zero trip rule in 40% of space
        drawZeroTripRepeat colTree, _
                           colSubnodeHandle, _
                           usrRectangle2, _
                           lblSmallConnector, _
                           lblExitConnector
    End With
End Sub
Private Sub drawOptionalSequence(ByRef colTree As Collection, _
                                 ByRef colSubnodeHandle As Collection, _
                                 ByRef usrRectangle As TYPrectangle, _
                                 ByRef lblEntryConnector As Label, _
                                 ByRef lblExitConnector As Label)
    '
    ' Draw optional sequence
    '
    '
    ' --- Make some space for the bypass
    usrRectangle.lngHeight = usrRectangle.lngHeight * 0.9
    ' --- Draw the rule
    node2diagram colTree, CLng(colSubnodeHandle.item(2)), "", usrRectangle
    ' --- Draw the optional bypass
    drawByPass usrRectangle, lblEntryConnector, lblExitConnector
End Sub
Private Sub drawByPass(ByRef usrAround As TYPrectangle, _
                       ByRef lblFromConnector As Label, _
                       ByRef lblToConnector As Label)
    '
    ' Draw sequence bypass
    '
    '
End Sub
Private Sub drawProduction(ByRef colTree As Collection, _
                           ByRef colSubnodeHandle As Collection, _
                           ByRef usrRectangle As TYPrectangle)
    '
    ' Draw production
    '
    '
    Dim colNT As Collection
    Dim lngWidth As Long
    Dim lngX As Long
    With usrRectangle
        Set colNT = colTree.item(1)
        lngWidth = .lngWidth
        ' --- Draw the nonterminal that is defined in 20% of the space
        .lngWidth = CLng(.lngWidth * 0.2)
        lngX = rectangleRight(usrRectangle)
        drawGS colNT.item(colSubnodeHandle.item(2)), usrRectangle, True
        ' --- Draw the defining rule in 70% of the space
        .lngWidth = CLng(lngWidth * 0.7)
        node2diagram colTree, CLng(colSubnodeHandle.item(3)), "", usrRectangle
        ' --- Connect the nt and the rule in the remaining 10% of the space
        drawLine lngX, .lngTop, lngX + lngWidth * 0.1, .lngTop
    End With
End Sub
Private Sub drawSequence(ByRef colTree As Collection, _
                         ByRef colSubnodeHandle As Collection, _
                         ByRef usrRectangle As TYPrectangle)
    '
    ' Draw sequence
    '
    '
    Dim lngConnectorWidth As Long
    Dim lngWidth As Long
    Dim lngX As Long
    With usrRectangle
        ' --- Draw the first rule in 40% of the space
        lngWidth = .lngWidth
        .lngWidth = .lngWidth * 0.4
        node2diagram colTree, CLng(colSubnodeHandle.item(2)), "", usrRectangle
        ' --- Draw the second rule in 40% of the space
        lngConnectorWidth = lngWidth * 0.2
        .lngLeft = rectangleRight(usrRectangle) + lngConnectorWidth
        node2diagram colTree, CLng(colSubnodeHandle.item(3)), "", usrRectangle
        ' --- Connect the nt and the rule in the remaining 20% of the space
        drawLine lngX, .lngTop, lngX + lngConnectorWidth, .lngTop
    End With
End Sub
Private Sub drawZeroTripRepeat(ByRef colTree As Collection, _
                               ByRef colSubnodeHandle As Collection, _
                               ByRef usrRectangle As TYPrectangle, _
                               ByRef lblEntryConnector As Label, _
                               ByRef lblExitConnector As Label)
    '
    ' Draw zero-trip repetition of a rule
    '
    '
    ' --- Draw the repeated rule
    node2diagram colTree, CLng(colSubnodeHandle.item(2)), "", usrRectangle
    ' --- Draw the repetition connector
    drawRepetitionConnector lblEntryConnector, lblExitConnector
End Sub
Private Function drawRepetitionConnector(ByRef lblEntryConnector As Label, _
                                         ByRef lblExitConnector As Label) _
        As Label
    '
    ' Draw the cycle connector
    '
    '
End Function
Private Function rectangleRight(ByRef usrRectangle As TYPrectangle) As Long
    '
    ' Returns the rectangle's right side
    '
    '
    With usrRectangle
        rectangleRight = .lngLeft + .lngWidth
    End With
End Function
Private Function rectangleBottom(ByRef usrRectangle As TYPrectangle) As Long
    '
    ' Returns the rectangle's bum
    '
    '
    With usrRectangle
        rectangleBottom = .lngTop + .lngHeight
    End With
End Function
Private Sub expandForm(ByVal lngNewWidth As Long, _
                       ByVal lngNewHeight As Long)
    '
    ' Expand the form
    '
    '
    With Me
        .Width = lngNewWidth + GRIDSIZE
        .ScaleHeight = lngNewHeight + GRIDSIZE
        centerForm
        .Refresh
    End With
End Sub
Private Sub centerForm()
    '
    ' Recenter the form
    '
    '
    With Me
        .Left = Screen.Width \ 2
        .Top = Screen.Height \ 2
        .Refresh
    End With
End Sub
Private Function drawLeadLine(ByRef usrRectangle As TYPrectangle) As Label
    '
    ' Draw the line leading into a rule and return it as the label
    '
    '
    With usrRectangle
        drawLeadLine = drawConnectorFromXY(.lngLeft, _
                                            .lngTop, _
                                            .lngWidth - .lngWidth * CONNECTORSPACE \ 2, _
                                            .lngTop)
    End With
End Function
Private Function drawEndLine(ByRef usrRectangle As TYPrectangle) As Label
    '
    ' Draw the line leading out of rule and return it as the label
    '
    '
    With usrRectangle
        drawEndLine = drawConnectorFromXY(rectangleRight(usrRectangle) _
                                            - _
                                            .lngWidth * CONNECTORSPACE \ 2, _
                                            .lngTop, _
                                            .lngWidth, _
                                            .lngTop)
    End With
End Function
Private Sub moveRectangle(ByRef usrRectangle As TYPrectangle, _
                          ByVal lngX As Long, _
                          ByVal lngY As Long)

End Sub
Private Function drawConnectorFromLbl(ByRef lblFrom As Label, _
                                      ByVal lngXTo As Long, _
                                      ByVal lngYTo As Long) As Label
                               
End Function
Private Function drawConnectorFromXY(ByVal lngXFrom As Long, _
                                      ByVal lngYFrom As Long, _
                                      ByVal lngXTo As Long, _
                                      ByVal lngYTo As Long) As Label
                               
End Function
Private Function drawLine(ByVal lngXFrom As Long, _
                          ByVal lngYFrom As Long, _
                          ByVal lngXTo As Long, _
                          ByVal lngYTo As Long) As Label
                               
End Function
Private Function drawLabel(ByRef usrRectangle As TYPrectangle, _
                           ByVal strCaption As String, _
                           ByVal lngBackColor As Long, _
                           ByVal lngForeColor As Long) As Label

End Function
