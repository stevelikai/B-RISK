Option Strict Off
Option Explicit On

Imports System.Windows.Forms.DataVisualization.Charting

'Imports AxMSChart20Lib

Public Class frmFire
    Inherits System.Windows.Forms.Form

    Private Sub cmdAddObject_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAddObject.Click
        '*  ===================================================
        '*      This event adds a burning object when the user
        '*      clicks on the add object command button.
        '*  ===================================================

        'show the fire database
        Me.Hide()
        frmFireObjectDB.ToolStripButton2.Visible = True
        frmFireObjectDB.ToolStripLabel1.Visible = False
        frmFireObjectDB.Show()

    End Sub

    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        '*  ====================================================
        '*      This event makes the fire specification form
        '*      invisible.
        '*  ====================================================
        'disable menu item run if no heat release data entered
        Dim i As Integer

        Try

            'Resize_Objects()

            For i = 0 To NumberObjects
                If NumberDataPoints(i) <> 0 And NumberObjects > 0 Then
                    'MDIFrmMain.mnuRun.Enabled = True
                    'MDIFrmMain.mnuRun.Visible = True
                    Exit For
                Else
                    'MDIFrmMain.mnuRun.Enabled = False
                    'MDIFrmMain.mnuRun.Visible = False
                End If
            Next i
            Me.Close()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in frmFire.vb cmdClose_Click")
        End Try

    End Sub

    Private Sub cmdDeleteObject_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDeleteObject.Click
        '*  ====================================================
        '*      This event deletes a burning object from the
        '*      analysis.
        '*  ====================================================

        Dim id As Short
        Dim j, i, k As Integer
        Dim msg As String
        Dim response As Short

        'if there are no objects or if there is only one object then proceed no further
        If NumberObjects < 1 Then Exit Sub

        'identify the currently selected object
        id = lstObjectID.SelectedIndex + 1

        'check to sure if user really wants to delete the object
        msg = "Are You Sure You Want To Delete Object " & CStr(id) & "?"
        response = MsgBox(msg, MB_YESNO + MB_ICONQUESTION, ProgramTitle)

        'if they answer no then proceed no further
        If response = IDNO Then Exit Sub

        'call procedure to resize the object arrays
        Resize_Objects()

        'remove any empty spaces in the array
        If id < NumberObjects Then
            For i = id To NumberObjects - 1
                For j = 1 To NumberDataPoints(id + 1)
                    For k = 1 To 2
                        HeatReleaseData(k, j, i) = HeatReleaseData(k, j, i + 1)
                        MLRData(k, j, i) = MLRData(k, j, i + 1)
                        NumberDataPoints(i) = NumberDataPoints(i + 1)
                        EnergyYield(i) = EnergyYield(i + 1)
                        'COYield(i) = COYield(i + 1)
                        CO2Yield(i) = CO2Yield(i + 1)
                        SootYield(i) = SootYield(i + 1)
                        FireHeight(i) = FireHeight(i + 1)
                        FireLocation(i) = FireLocation(i + 1)
                        ObjectDescription(i) = ObjectDescription(i + 1)
                    Next k
                Next j
            Next i
        End If

        'erase the contents of the last empty object
        For j = 1 To NumberDataPoints(NumberObjects)
            For k = 1 To 2
                HeatReleaseData(k, j, NumberObjects) = 0
                MLRData(k, j, NumberObjects) = 0
                NumberDataPoints(NumberObjects) = 0
                EnergyYield(NumberObjects) = 0
                'COYield(NumberObjects) = 0
                HCNuserYield(NumberObjects) = 0
                ObjectDescription(NumberObjects) = ""
                CO2Yield(NumberObjects) = 0
                SootYield(NumberObjects) = 0
                FireHeight(NumberObjects) = 0
                FireLocation(NumberObjects) = 0
            Next k
        Next j

        'update the textboxes
        txtEnergyYield.Text = VB6.Format(EnergyYield(NumberObjects - 1), "0.0")
        txtCO2Yield.Text = VB6.Format(CO2Yield(NumberObjects - 1), "0.000")
        txtSootYield.Text = VB6.Format(SootYield(NumberObjects - 1), "0.000")
        txtHCNYield.Text = VB6.Format(HCNuserYield(NumberObjects - 1), "0.000")
        lblObjectDescription.Text = ObjectDescription(NumberObjects - 1)
        txtFireHeight.Text = VB6.Format(FireHeight(NumberObjects - 1), "0.000")
        lstObjectLocation.SelectedIndex = FireLocation(NumberObjects - 1)

        'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        'Graph2.CtlRefresh()

        'decrement the object counter by 1
        NumberObjects = NumberObjects - 1

        'remove current item from listbox
        'If NumberObjects > 1 Then
        lstObjectID.Items.RemoveAt((lstObjectID.SelectedIndex))
        'End If

        'update list box
        lstObjectID.Items.Clear()
        If NumberObjects > 0 Then
            For i = 1 To NumberObjects
                lstObjectID.Items.Add("Fire " & CStr(i) & " of " & CStr(NumberObjects))
            Next i
            lstObjectID.SelectedIndex = NumberObjects - 1
            GroupBox1.Visible = True
        Else
            lstObjectID.Items.Add("None")
            lstObjectID.SelectedIndex = 0
            GroupBox1.Visible = False
        End If

        'Call cmdPlot_click
        'Me.Frame1.SendToBack()

    End Sub
    Private Sub plot_graph()
        '*  ====================================================
        '*      Show a graph of heat release rate versus time for
        '*      the current object.
        '*  ====================================================

        Dim i As Long, id As Integer
        Dim NumberPoints As Integer
        Dim title As String
        Try


            id = Me.lstObjectID.SelectedIndex + 1
            title = "Rate of heat release (kW)"
            NumberPoints = NumberDataPoints(id)

            Dim chdata(0 To NumberPoints - 1, 0 To 1) As Object
            Dim ydata(0 To NumberPoints - 1) As Double



            frmPlot.Chart1.Series.Clear()
            'room = 1
            'For room = 1 To NumberRooms

            frmPlot.Chart1.Series.Add("Fire " & id)

            frmPlot.Chart1.Series("Fire " & id).ChartType = SeriesChartType.FastLine

            'For j = 1 To NumberTimeSteps
            'ydata(j) = (1 * datatobeplotted(room, j, layer) + 0) 'data to be plotted
            'frmPlot.Chart1.Series("Room " & room).Points.AddXY(tim(j, 1), ydata(j))
            'Next

            For i = 1 To NumberPoints
                chdata(i - 1, 0) = HeatReleaseData(1, i, id) 'time
                chdata(i - 1, 1) = HeatReleaseData(2, i, id) 'hrr
                ydata(i - 1) = HeatReleaseData(2, i, id) 'hrr
                frmPlot.Chart1.Series(("Fire " & id)).Points.AddXY(chdata(i - 1, 0), ydata(i - 1))

            Next

            ' Next room
            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = title
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            'frmPlot.Chart1.Titles("Title1").Text = Title
            ' Disable X axis margin

            frmPlot.Chart1.Visible = True
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in frmFire.vb plot_graph")
        End Try
    End Sub
    Private Sub cmdViewData_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdViewData.Click
        '*  ===================================================
        '*      This event views the heat release rate data
        '*      entered for the currently selected burning object
        '*      when the user clicks on the View Data command button.
        '*  ===================================================

        Dim id As Short

        'identify which fire object is selected
        id = lstObjectID.SelectedIndex + 1
        'Me.Graph2.Visible = False

    End Sub

    Private Sub frmFire_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '*  =============================================
        '*      This event centres the form on the screen
        '*      when the form first loads.
        '*  =============================================

        Dim i, room As Integer

        'call procedure to centre form
        Centre_Form(Me)

        'add objects to list box
        If NumberObjects > 0 Then
            For i = 1 To NumberObjects
                lstObjectID.Items.Add("Fire " & CStr(i) & " of " & CStr(NumberObjects))
            Next i
            GroupBox1.Visible = True
        Else
            lstObjectID.Items.Add("None")
            GroupBox1.Visible = False
        End If

        lstObjectLocation.Items.Clear()
        lstObjectLocation.Items.Add("Centre")
        lstObjectLocation.Items.Add("Wall")
        lstObjectLocation.Items.Add("Corner")
        lstObjectLocation.SelectedIndex = 0
        lstObjectLocation.Refresh()

        On Error Resume Next
        Me.lstObjectID.SelectedIndex = 0
        lstObjectID.Refresh()

        For room = 1 To NumberRooms
            Me.lstFireRoom2.Items.Add(CStr(room))
        Next room
        Me.lstFireRoom2.SelectedIndex = 0
        Me.lstFireRoom2.Refresh()
        'Me.Graph2.Visible = True

    End Sub

    Private Sub frmFire_Paint(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        'cmdPlot_Click(cmdPlot, New System.EventArgs())
    End Sub


    Private Sub lstFireRoom2_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstFireRoom2.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*  ===================================================

        Dim id As Short

        id = lstFireRoom2.SelectedIndex + 1
        fireroom = id
    End Sub

    Sub lstFireRoom2_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstFireRoom2.SelectedIndexChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================

        Dim id As Short
        On Error Resume Next

        If lstFireRoom2.SelectedIndex = -1 Then lstFireRoom2.SelectedIndex = 0
        id = lstFireRoom2.SelectedIndex + 1
        fireroom = id
    End Sub

    Private Sub lstObjectID_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstObjectID.SelectedIndexChanged
        '*  ====================================================
        '*      This event updates the contents in the
        '*      text boxes when the user clicks on the list box.
        '*  ====================================================

        Dim id As Short
        If NumberObjects = 0 Then Exit Sub
        If NumberObjects + 1 <> NumberDataPoints.Count Then Resize_Objects()

        'identify current object selected
        id = lstObjectID.SelectedIndex + 1

        'update text boxes
        txtEnergyYield.Text = VB6.Format(EnergyYield(id), "0.0")
        txtCO2Yield.Text = VB6.Format(CO2Yield(id), "0.000")
        txtSootYield.Text = VB6.Format(SootYield(id), "0.000")
        txtHCNYield.Text = VB6.Format(HCNuserYield(id), "0.000")
        lblObjectDescription.Text = ObjectDescription(id)
        txtFireHeight.Text = VB6.Format(FireHeight(id), "0.000")

        If lstObjectLocation.Items.Count = 0 Then Exit Sub

        'update object location list box
        On Error Resume Next
        lstObjectLocation.SelectedIndex = FireLocation(id)


    End Sub

    Private Sub lstObjectLocation_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstObjectLocation.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================

        Dim id As Short
        id = lstObjectID.SelectedIndex + 1
        FireLocation(id) = lstObjectLocation.SelectedIndex
    End Sub

    Private Sub lstObjectLocation_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstObjectLocation.SelectedIndexChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================

        Dim id As Short
        On Error Resume Next
        If lstObjectID.Items.Count > 0 Then
            If lstObjectID.SelectedIndex = -1 Then lstObjectID.SelectedIndex = 0
            id = lstObjectID.SelectedIndex + 1
            FireLocation(id) = lstObjectLocation.SelectedIndex
        End If

    End Sub

    Private Sub txtCO2Yield_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCO2Yield.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================

        Dim id As Short

        'identify which fire object is selected
        id = lstObjectID.SelectedIndex + 1
        If id >= 0 Then
            If IsNumeric(txtCO2Yield.Text) = True Then
                'store the carbon dioxide yield as a variable
                CO2Yield(id) = CSng(txtCO2Yield.Text)
            End If
        End If
    End Sub

    Private Sub txtCO2Yield_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCO2Yield.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*  ===================================================
        '*      This event checks for valid input when the user
        '*      pressing the enter key and if valid stores the
        '*      data as a variable.
        '*  ===================================================

        Dim A As Single
        Dim id As Short
        If IsNumeric(txtCO2Yield.Text) = True Then
            A = CSng(txtCO2Yield.Text)
            'identify which fire object is selected
            id = lstObjectID.SelectedIndex + 1

            If KeyAscii = 13 Then
                If A >= 0 Then
                    CO2Yield(id) = A
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCO2Yield_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCO2Yield.Leave
        '*  =====================================================
        '*      This event checks for valid data when the control
        '*      loses the focus.
        '*  =====================================================

        Dim A As Single
        Dim id As Short
        If IsNumeric(txtCO2Yield.Text) = True Then
            A = CSng(txtCO2Yield.Text)

            'identify which fire object is selected
            id = lstObjectID.SelectedIndex + 1

            If A >= 0 Then
            Else
                MsgBox("Invalid Entry!")
                txtCO2Yield.Focus()
            End If
        Else
            MsgBox("Invalid Entry!")
            txtCO2Yield.Focus()
        End If
    End Sub

    Private Sub txtEnergyYield_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEnergyYield.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================

        Dim id As Short

        'identify which fire object is selected
        id = lstObjectID.SelectedIndex + 1
        If id >= 0 Then
            If IsNumeric(txtEnergyYield.Text) = True Then
                'store the energy yield as a variable
                EnergyYield(id) = CSng(txtEnergyYield.Text)
            End If
        End If
    End Sub

    Private Sub txtEnergyYield_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtEnergyYield.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*  ===================================================
        '*      This event checks for valid input when the user
        '*      pressing the enter key and if valid stores the
        '*      data as a variable.
        '*  ===================================================

        Dim id As Short
        Dim A As Single
        If IsNumeric(txtEnergyYield.Text) = True Then
            A = CSng(txtEnergyYield.Text)

            'identify which fire object is selected
            id = lstObjectID.SelectedIndex + 1
            If KeyAscii = 13 Then
                If A > 0 Then
                    EnergyYield(id) = A
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtEnergyYield_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEnergyYield.Leave
        '*  =====================================================
        '*      This event checks for valid data when the control
        '*      loses the focus.
        '*  =====================================================

        Dim id As Short
        Dim A As Single
        If IsNumeric(txtEnergyYield.Text) = True Then
            A = CSng(txtEnergyYield.Text)

            'identify which fire object is selected
            id = lstObjectID.SelectedIndex

            If A > 0 Then
            ElseIf id > 0 Then

                MsgBox("Invalid Entry. Must be > 0!")
                txtEnergyYield.Focus()
            Else
            End If
        Else
            MsgBox("Invalid Entry. Must be > 0!")
            txtEnergyYield.Focus()
        End If
    End Sub

    Private Sub txtFireHeight_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFireHeight.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================

        Dim id As Short

        'identify which fire object is selected

        id = lstObjectID.SelectedIndex + 1
        If id >= 0 Then
            If IsNumeric(txtFireHeight.Text) = True Then
                'store the height of the base of the fire as a variable
                FireHeight(id) = CSng(txtFireHeight.Text)
            End If
        End If
    End Sub

    Private Sub txtFireHeight_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFireHeight.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*  ===================================================
        '*      This event checks for valid input when the user
        '*      pressing the enter key and if valid stores the
        '*      data as a variable.
        '*  ===================================================

        Dim id As Short
        Dim A As Single
        If IsNumeric(txtFireHeight.Text) = True Then
            A = CSng(txtFireHeight.Text)

            'identify which fire object is selected
            id = lstObjectID.SelectedIndex + 1

            If KeyAscii = 13 Then
                If A >= 0 And A < RoomHeight(fireroom) Then
                    FireHeight(id) = A
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFireHeight_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFireHeight.Leave
        '*  =====================================================
        '*      This event checks for valid data when the control
        '*      loses the focus.
        '*   20/3/08 CW - fixed bug in error checking height of fire in the room
        '*  =====================================================

        Dim id As Short
        Dim A As Single
        If IsNumeric(txtFireHeight.Text) = True Then
            A = CSng(txtFireHeight.Text)

            'identify which fire object is selected
            id = lstObjectID.SelectedIndex + 1

            If A >= 0 And A < RoomHeight(fireroom) Then
            Else
                MsgBox("Invalid Entry!")
                txtFireHeight.Focus()
            End If
        Else
            MsgBox("Invalid Entry!")
            txtFireHeight.Focus()
        End If
    End Sub

    Private Sub txtSootYield_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSootYield.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================

        Dim id As Short

        'identify which fire object is selected
        id = lstObjectID.SelectedIndex + 1
        If id >= 0 Then
            If IsNumeric(txtSootYield.Text) = True Then
                'store the soot yield as a variable
                SootYield(id) = CSng(txtSootYield.Text)
            End If
        End If
    End Sub

    Private Sub txtSootYield_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSootYield.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*  ===================================================
        '*      This event checks for valid input when the user
        '*      pressing the enter key and if valid stores the
        '*      data as a variable.
        '*  ===================================================

        Dim A As Single
        Dim id As Short
        If IsNumeric(txtSootYield.Text) = True Then
            A = CSng(txtSootYield.Text)

            'identify which fire object is selected
            id = lstObjectID.SelectedIndex + 1

            If KeyAscii = 13 Then
                If A >= 0 Then
                    SootYield(id) = A
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtSootYield_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSootYield.Leave
        '*  =====================================================
        '*      This event checks for valid data when the control
        '*      loses the focus.
        '*  =====================================================

        Dim id As Short
        Dim A As Single
        If IsNumeric(txtSootYield.Text) = True Then
            A = CSng(txtSootYield.Text)

            'identify which fire object is selected
            id = lstObjectID.SelectedIndex + 1

            If A >= 0 Then
            Else
                MsgBox("Invalid Entry!")
                txtSootYield.Focus()
            End If
        Else
            MsgBox("Invalid Entry!")
            txtSootYield.Focus()
        End If
    End Sub

    Private Sub txtHCNYield_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtHCNYield.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================

        Dim id As Short

        'identify which fire object is selected
        id = lstObjectID.SelectedIndex + 1
        If id >= 0 Then
            If IsNumeric(txtHCNYield.Text) = True Then
                'store the hcn yield as a variable
                HCNuserYield(id) = CSng(txtHCNYield.Text)
            End If
        End If
    End Sub

    Private Sub txtHCNYield_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtHCNYield.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*  ===================================================
        '*      This event checks for valid input when the user
        '*      pressing the enter key and if valid stores the
        '*      data as a variable.
        '*  ===================================================

        Dim A, id As Short
        If IsNumeric(txtHCNYield.Text) = True Then
            A = CSng(txtHCNYield.Text)

            'identify which fire object is selected
            id = lstObjectID.SelectedIndex + 1

            If KeyAscii = 13 Then
                If A >= 0 Then
                    HCNuserYield(id) = A
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txthcnYield_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtHCNYield.Leave
        '*  =====================================================
        '*      This event checks for valid data when the control
        '*      loses the focus.
        '*  =====================================================

        Dim id As Short
        Dim A As Single
        If IsNumeric(txtHCNYield.Text) = True Then
            A = CSng(txtHCNYield.Text)

            'identify which fire object is selected
            id = lstObjectID.SelectedIndex + 1

            If A >= 0 Then
            Else
                MsgBox("Invalid Entry!")
                txtHCNYield.Focus()
            End If
        Else
            MsgBox("Invalid Entry!")
            txtHCNYield.Focus()
        End If
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub cmdPlot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPlot.Click
        Call plot_graph()
    End Sub

End Class