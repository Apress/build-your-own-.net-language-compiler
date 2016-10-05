Option Strict On

Public Class zoomInterface

    Private Shared _OBJutilities As utilities.utilities

    ' -----------------------------------------------------------------
    ' Zoom control interface
    '
    '
    Friend Shared Sub zoomInterface(ByVal frmParent As Form, _
                                    ByVal objZoomed As Object, _
                                    Optional ByVal dblWidth As Double = 1.5, _
                                    Optional ByVal dblHeight As Double = 3)
        Dim ctlZoomedHandle As Control
        If (TypeOf objZoomed Is System.String) Then
            Dim txtBox As TextBox
            Try
                txtBox = New TextBox
                With txtBox
                    .Height = frmParent.Height
                    .Width = frmParent.Width
                    .Multiline = True
                    .Text = CStr(objZoomed)
                End With
            Catch ex As Exception
                Dim _OBJutilities As utilities.utilities
                _OBJutilities.errorHandler("Can't create textbox: " & _
                                            Err.Number & " " & Err.Description, _
                                            "displayInfoForm", _
                                            "Returning to caller" & vbNewLine & vbNewLine & _
                                            ex.ToString)
                Return
            End Try
            ctlZoomedHandle = txtBox
        ElseIf (TypeOf objZoomed Is Control) Then
            ctlZoomedHandle = CType(objZoomed, Control)
        End If
        Dim ctlZoom As zoom.zoom
        Try
            ctlZoom = New zoom.zoom(ctlZoomedHandle)
        Catch ex As Exception
            _OBJutilities.errorHandler("Cannot create zoom control: " & _
                         Err.Number & " " & Err.Description, _
                         "zoomInterface", _
                         ex.ToString)
            Return
        End Try
        With ctlZoom
            .setSize(dblWidth, dblHeight)
            .showZoom()
            .dispose()
        End With
    End Sub
End Class
