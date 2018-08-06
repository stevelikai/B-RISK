Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Collections.Generic

Public Class frmSensitivity

    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    'Public Sub cmdPlot_Click(sender As Object, e As EventArgs) Handles cmdPlot2.Click

    '    '    sensitivity plot for max upper layer temperature - output


    '    '    Dim DataShift As Double
    '    '    Dim DataMultiplier, XMultiplier As Double
    '    '    Dim k, counter, OutputInterval, NumberIterations, IterationsCompleted As Integer
    '    '    Dim j As Integer
    '    '    Dim room As Integer
    '    '    Dim GraphStyle As Integer
    '    '    Dim Title As String = ""
    '    '    Dim xTitle As String = ""
    '    '    Dim yTitle As String = ""

    '    '    If nO Then data exists
    '    '    If NumberTimeSteps < 1 Then
    '    '            MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
    '    '            Exit Sub
    '    '        End If

    '    '        Try
    '    '            Chart1.Visible = False 'histograms
    '    '            Chart2.Visible = False 'cdf plots
    '    '            Chart3.Visible = True 'time-series plots
    '    '            Label4.Visible = False
    '    '            Label5.Visible = False
    '    '            Panel1.BringToFront()

    '    '            ToolStrip1.Visible = False

    '    '            xTitle = " Time (s)"
    '    '            yTitle = "Max upper layer temp (C)"

    '    '            counter = 1
    '    '            room = 1
    '    '            DataMultiplier = 1
    '    '            DataShift = 0
    '    '            GraphStyle = 4 '2=user-defined
    '    '            MaxYValue = 900
    '    '            XMultiplier = 1 'time multipier

    '    '            Title = "Output Data - Sensitivity Plot: " & "Room " & NumericUpDownRoom.Value & ". " & IterationsCompleted & " iterations"
    '    '            Chart3.Titles("Title1").Text = Title

    '    '            Chart3.Legends("Legend1").Enabled = False
    '    '            Chart3.Dock = DockStyle.Fill
    '    '            Chart3.ChartAreas("ChartArea1").BorderWidth = 1
    '    '            Chart3.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
    '    '            Chart3.ChartAreas("ChartArea1").AxisY.Title = yTitle
    '    '            Chart3.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0"
    '    '            Chart3.ChartAreas("ChartArea1").AxisX.LabelStyle.Format = "0.00"
    '    '            Chart3.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
    '    '            Chart3.ChartAreas("ChartArea1").AxisY.Maximum = [Double].NaN
    '    '            Chart3.ChartAreas("ChartArea1").AxisX.Title = "Vent Width (m)"
    '    '            Chart3.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False


    '    '            max Number of time periods for which the data Is available
    '    '        If IsNumeric(OutputInterval) And OutputInterval > 0 Then counter = Math.Ceiling((SimTime / OutputInterval))

    '    '            NumericUpDown1.Visible = False
    '    '            PercentileUpDown.Visible = False
    '    '            Label2.Visible = False
    '    '            Label18.Visible = False
    '    '            Label4.Visible = False
    '    '            Label5.Visible = False
    '    '            Label19.Visible = True
    '    '            NumericUpDownRoom.Minimum = 1
    '    '            NumericUpDownRoom.Maximum = NumberRooms
    '    '            NumericUpDownRoom.Increment = 1
    '    '            NumericUpDownRoom.Value = 1
    '    '            NumericUpDownRoom.Show()
    '    '            IterationUpDown.Show()
    '    '            IterationUpDown.Value = 0
    '    '            IterationUpDown.Maximum = NumberIterations
    '    '            Label40.Show()

    '    '        -------------
    '    '        *====================================================================
    '    '        *  This function takes data for a variable from a two-dimensional array
    '    '        *  And displays it in a graph
    '    '        *====================================================================

    '    '        Dim ydata(0 To counter + 1) As Double
    '    '            NumericUpDown_Bins.Hide()
    '    '            Label45.Hide()
    '    '            Chart3.Series.Clear()
    '    '            room = 1

    '    '            plot multiple series
    '    '        Chart3.Series.Add(CStr(1))
    '    '            Chart3.Series(CStr(1)).ChartType = SeriesChartType.Point
    '    '            Dim Yval As Double = 0
    '    '            Dim Xval As Double = 0
    '    '            Dim ventid As Integer = 1

    '    '            For k = 1 To IterationsCompleted

    '    '                Xval = mc_vent_width(ventid - 1, k - 1)

    '    '                find when gas temp 
    '    '            Yval = 0
    '    '                For j = 0 To counter
    '    '                    If mc_ULTemp(room, j, k - 1) > Yval Then
    '    '                        Yval = mc_ULTemp(room, j, k - 1)
    '    '                    End If
    '    '                Next
    '    '                Chart3.Series(CStr(1)).Points.AddXY(Xval, Yval)

    '    '            Next k

    '    '            Chart3.Visible = True

    '    '            PageToolStripMenuItem.Visible = True
    '    '            PrintPreviewToolStripMenuItem.Visible = True
    '    '            PrintToolStripMenuItem.Visible = True

    '    '        Catch ex As Exception
    '    '            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in frmSensitivity.vb cmdPlot_Click")
    '    '        End Try
    'End Sub

    Private Sub frmSensitivity_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub LB_input_Click(sender As Object, e As EventArgs) Handles LB_input.Click
        Try
            Dim var As Integer = LB_input.SelectedIndex

            Dim ovents As New List(Of oVent)
            ovents = VentDB.GetVents
            Dim number_vents As Integer = ovents.Count

            Dim ocvents As New List(Of oCVent)
            ocvents = VentDB.GetCVents
            Dim number_cvents As Integer = ocvents.Count

            Dim oitems As New List(Of oItem)
            oitems = ItemDB.GetItemsv2
            Dim number_items As Integer = oitems.Count

            LB_input_ID.Items.Clear()

            Select Case var
                Case 0 'Vent width
                    Label2.Text = "Vent ID"
                    For i = 1 To number_vents
                        LB_input_ID.Items.Add(i.ToString)
                    Next
                Case 1 'Vent height
                    Label2.Text = "Vent ID"
                    For i = 1 To number_vents
                        LB_input_ID.Items.Add(i.ToString)
                    Next
                Case 2 'Room length
                    Label2.Text = "Room ID"
                    For i = 1 To NumberRooms
                        LB_input_ID.Items.Add(i.ToString)
                    Next
                Case 3 'Room width
                    Label2.Text = "Room ID"
                    For i = 1 To NumberRooms
                        LB_input_ID.Items.Add(i.ToString)
                    Next
                Case 4 'Heat of combustion by item
                    Label2.Text = "Item ID"
                    For i = 1 To number_items
                        LB_input_ID.Items.Add(i.ToString)
                    Next
                Case 5, 6, 7, 8 'soot yield by item
                    Label2.Text = "Item ID"
                    For i = 1 To number_items
                        LB_input_ID.Items.Add(i.ToString)
                    Next
                Case 9 'ceiling vent area
                    Label2.Text = "CVent ID"
                    For i = 1 To number_cvents
                        LB_input_ID.Items.Add(i.ToString)
                    Next
            End Select
            LB_input_ID.SelectedIndex = 0

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in LB_input_click")
        End Try
    End Sub


    Private Sub cmdPlotSENS_Click(sender As Object, e As EventArgs) Handles cmdPlotSENS.Click

        Dim var1 As Integer = LB_input.SelectedIndex
        Dim var2 As Integer = LB_input_ID.SelectedIndex
        Dim var3 As Integer = LB_output.SelectedIndex
        Dim xTitle As String = ""
        Dim yTitle As String = ""
        Dim yval As Double
        Dim k, counter As Integer
        Dim j As Integer
        Dim room As Integer = fireroom
       
        Dim Title As String = ""
        Dim xval(0 To IterationsCompleted) As Double

        If IsNumeric(OutputInterval) And OutputInterval > 0 Then counter = Math.Ceiling((SimTime / OutputInterval))

        Try
            If NumberTimeSteps < 1 Then
                MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
                Exit Sub
            End If

            frmInputs.NumericUpDown_Bins.Hide()
            frmInputs.Label45.Hide()
            frmInputs.Chart3.Series.Clear()
            frmInputs.Chart1.Visible = False 'histograms
            frmInputs.Chart2.Visible = False 'cdf plots

            frmInputs.ChkAutosavePdf.Visible = False
            frmInputs.chkAutosaveXL.Visible = False
            frmInputs.Chart3.Visible = True 'time-series plots
            frmInputs.Label4.Visible = False
            frmInputs.Label5.Visible = False
            frmInputs.Panel1.BringToFront()
            frmInputs.ToolStrip1.Visible = False
            frmInputs.Chart3.Series.Add(CStr(1))
            frmInputs.Chart3.Series(CStr(1)).ChartType = SeriesChartType.Point

            Title = "Output Data - Sensitivity Plot: "
            frmInputs.Chart3.Titles("Title1").Text = Title

            Select Case var1
                Case 0 'Vent width
                    xTitle = "Vent width (m)"

                    For k = 1 To IterationsCompleted
                        xval(k) = mc_vent_width(var2, k - 1)
                    Next k
                Case 1 'Vent height
                    xTitle = "Vent height (m)"

                    For k = 1 To IterationsCompleted
                        xval(k) = mc_vent_height(var2, k - 1)
                    Next k
                Case 2 'Room length
                    xTitle = "Room length (m)"

                    For k = 1 To IterationsCompleted
                        xval(k) = mc_room_length(var2, k - 1)
                    Next k
                Case 3 'Room width
                    xTitle = "Room width (m)"

                    For k = 1 To IterationsCompleted
                        xval(k) = mc_room_width(var2, k - 1)
                    Next k
                Case 4 'Heat of combustion by item
                    xTitle = "Heat of combustion (kJ/g)"

                    For k = 1 To IterationsCompleted
                        xval(k) = mc_item_hoc(var2, k - 1)
                    Next k
                Case 5 'Soot yield by item
                    xTitle = "Soot yield (g/g)"

                    For k = 1 To IterationsCompleted
                        xval(k) = mc_item_soot(var2, k - 1)
                    Next k
                Case 6 'RLF by item
                    xTitle = "Radiant loss fraction (-)"

                    For k = 1 To IterationsCompleted
                        xval(k) = mc_item_RLF(var2, k - 1)
                    Next k
                Case 7 'HRRPUA by item
                    xTitle = "HRRPUA (kW/m2)"

                    For k = 1 To IterationsCompleted
                        xval(k) = mc_item_hrrua(var2, k - 1)
                    Next k
                Case 8 'Latent heat of gasification by item
                    xTitle = "Latent heat of gasification (kJ/g)"

                    For k = 1 To IterationsCompleted
                        xval(k) = mc_item_lhog(var2, k - 1)
                    Next k
                Case 9 'ceiling vent area
                    xTitle = "Ceiling vent area (m2)"

                    For k = 1 To IterationsCompleted
                        xval(k) = mc_vent_area(var2, k - 1)
                    Next k
            End Select

            'output criteria
            Select Case var3
                Case 0 'Maximum upper layer temperature
                    yTitle = "Max upper layer temp (C)"

                    For k = 1 To IterationsCompleted
                        yval = 0
                        For j = 0 To counter
                            If mc_ULTemp(room, j, k - 1) > yval Then
                                yval = mc_ULTemp(room, j, k - 1)
                            End If
                        Next

                        frmInputs.Chart3.Series(CStr(1)).Points.AddXY(xval(k), yval)

                    Next k

                Case 1 'Time for layer height to drop below 2 m
                    yTitle = "Time for layer height to drop below " & MonitorHeight.ToString & " m (sec)"

                    For k = 1 To IterationsCompleted
                        yval = SimTime
                        For j = 0 To counter
                            If mc_LayerHeight(room, j, k - 1) < MonitorHeight Then
                                yval = j * OutputInterval
                                Exit For
                            End If
                        Next

                        frmInputs.Chart3.Series(CStr(1)).Points.AddXY(xval(k), yval)

                    Next k
                Case 2 'Time for visibility to drop below 10 m
                    yTitle = "Time for visibility to drop below 10 m (sec)"

                    For k = 1 To IterationsCompleted
                        yval = SimTime
                        For j = 0 To counter
                            If mc_visi(room, j, k - 1) < 10 Then
                                yval = j * OutputInterval
                                Exit For
                            End If
                        Next

                        frmInputs.Chart3.Series(CStr(1)).Points.AddXY(xval(k), yval)

                    Next k
                Case 3 'Time for visibility to drop below 5 m
                    yTitle = "Time for visibility to drop below 5 m (sec)"

                    For k = 1 To IterationsCompleted
                        yval = SimTime
                        For j = 0 To counter
                            If mc_visi(room, j, k - 1) < 5 Then
                                yval = j * OutputInterval
                                Exit For
                            End If
                        Next
                        frmInputs.Chart3.Series(CStr(1)).Points.AddXY(xval(k), yval)

                    Next k
                Case 4 'Time for FEDCO to exceed 0.3
                    yTitle = "Time for FEDCO to exceed 0.3 (sec)"

                    For k = 1 To IterationsCompleted
                        yval = SimTime
                        For j = 0 To counter
                            If mc_FEDgas(room, j, k - 1) > 0.3 Then
                                yval = j * OutputInterval
                                Exit For
                            End If
                        Next
                        frmInputs.Chart3.Series(CStr(1)).Points.AddXY(xval(k), yval)

                    Next k
                Case 5 'Time for FED thermal to exceed 0.3
                    yTitle = "Time for FED thermal to exceed 0.3 (sec)"

                    For k = 1 To IterationsCompleted
                        yval = SimTime
                        For j = 0 To counter
                            If mc_FEDheat(room, j, k - 1) > 0.3 Then
                                yval = j * OutputInterval
                                Exit For
                            End If
                        Next
                        frmInputs.Chart3.Series(CStr(1)).Points.AddXY(xval(k), yval)

                    Next k
                Case 6 'Time for upper layer temperature to exceed 500 C
                    yTitle = "Time for upper layer temperature to exceed 500 C (sec)"

                    For k = 1 To IterationsCompleted
                        yval = SimTime
                        For j = 0 To counter
                            If mc_ULTemp(room, j, k - 1) > 500 Then
                                yval = j * OutputInterval
                                Exit For
                            End If
                        Next
                        frmInputs.Chart3.Series(CStr(1)).Points.AddXY(xval(k), yval)

                    Next k
                Case Else
            End Select

            frmInputs.Chart3.Legends("Legend1").Enabled = False
            frmInputs.Chart3.Dock = DockStyle.Fill
            frmInputs.Chart3.ChartAreas("ChartArea1").BorderWidth = 1
            frmInputs.Chart3.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmInputs.Chart3.ChartAreas("ChartArea1").AxisY.Title = yTitle
            frmInputs.Chart3.ChartAreas("ChartArea1").AxisY.IsLabelAutoFit = True

            frmInputs.Chart3.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0"
            frmInputs.Chart3.ChartAreas("ChartArea1").AxisX.LabelStyle.Format = "0.00"
            frmInputs.Chart3.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            frmInputs.Chart3.ChartAreas("ChartArea1").AxisY.Maximum = [Double].NaN
            frmInputs.Chart3.ChartAreas("ChartArea1").AxisX.Title = xTitle
            frmInputs.Chart3.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmInputs.NumericUpDown1.Visible = False
            frmInputs.PercentileUpDown.Visible = False
            frmInputs.Label2.Visible = False
            frmInputs.Label18.Visible = False
            frmInputs.Label4.Visible = False
            frmInputs.Label5.Visible = False
            frmInputs.Label19.Visible = False
            frmInputs.PageToolStripMenuItem.Visible = True
            frmInputs.PrintPreviewToolStripMenuItem.Visible = True
            frmInputs.PrintToolStripMenuItem.Visible = True

            frmInputs.ChkAutosavePdf.Visible = False
            frmInputs.chkAutosaveXL.Visible = False
            frmInputs.Chart3.Visible = True
            Me.Close()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in frmSensitivity_click")
        End Try

    End Sub
End Class