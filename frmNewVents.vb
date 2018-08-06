Imports System.Collections.Generic
Public Class frmNewVents
    Public ventid As Integer
    Public oVent As oVent
    Public oVents As List(Of oVent)
    Public oventdistributions As List(Of oDistribution)


    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        Try
            If VM2 = True Then
                'do not change these inputs if using VM2 mode
                txtVentProb.Text = 0
                txtHOreliability.Text = 1
            End If


            oVents = VentDB.GetVents
            oventdistributions = VentDB.GetVentDistributions

            For Each Me.oVent In oVents

                If Me.oVent.id = Me.Tag Then

                    oVent.description = txtVentDescription.Text
                    oVent.fromroom = lstRoom1.SelectedItem
                    If IsNumeric(lstRoom2.SelectedItem) Then
                        oVent.toroom = lstRoom2.SelectedItem
                    Else
                        oVent.toroom = NumberRooms + 1
                    End If
                    oVent.face = lstRoomFace.SelectedIndex
                    oVent.offset = txtOffset.Text

                    If lstRoomFace.SelectedIndex = 0 Then 'front
                        oVent.walllength1 = RoomLength(lstRoom1.SelectedItem)
                        If IsNumeric(lstRoom2.SelectedItem) Then oVent.walllength2 = RoomLength(lstRoom2.SelectedItem)
                    ElseIf lstRoomFace.SelectedIndex = 1 Then 'right
                        oVent.walllength1 = RoomWidth(lstRoom1.SelectedItem)
                        If IsNumeric(lstRoom2.SelectedItem) Then
                            oVent.walllength2 = RoomWidth(lstRoom2.SelectedItem)
                        End If
                    ElseIf lstRoomFace.SelectedIndex = 2 Then 'rear
                        oVent.walllength1 = RoomLength(lstRoom1.SelectedItem)
                        If IsNumeric(lstRoom2.SelectedItem) Then oVent.walllength2 = RoomLength(lstRoom2.SelectedItem)
                    ElseIf lstRoomFace.SelectedIndex = 3 Then 'left
                        oVent.walllength1 = RoomWidth(lstRoom1.SelectedItem)
                        If IsNumeric(lstRoom2.SelectedItem) Then oVent.walllength2 = RoomWidth(lstRoom2.SelectedItem)
                    End If


                    oVent.height = txtVentHeight.Text
                    oVent.width = txtVentWidth.Text
                    oVent.sillheight = txtVentSillHeight.Text
                    oVent.opentime = txtVentOpenTime.Text
                    oVent.closetime = txtVentCloseTime.Text
                    If optGlassAutoBreak.Checked = True Then
                        oVent.autobreakglass = True
                    Else
                        oVent.autobreakglass = False
                    End If


                    If optStandard.Checked = True Then
                        oVent.spillplumemodel = 0 'standard door entrainment 
                    ElseIf opt2Dadhered.Checked = True Then
                        oVent.spillplumemodel = 1 '2D adhered
                    ElseIf opt2Dbalcony.Checked = True Then
                        oVent.spillplumemodel = 2 '2D balcony
                    ElseIf opt3Dadhered.Checked = True Then
                        oVent.spillplumemodel = 3 '3D adhered
                    ElseIf opt3Dchanbalcony.Checked = True Then
                        oVent.spillplumemodel = 4 '3D channelled balcony
                    ElseIf opt3Dunchanbalcony.Checked = True Then
                        oVent.spillplumemodel = 5 '3D unchannelled balcony

                    Else
                    End If

                    oVent.cd = txtVentCD.Text
                    oVent.probventclosed = txtVentProb.Text
                    oVent.horeliability = txtHOreliability.Text
                    oVent.spillbalconyprojection = txtBalconyWidth.Text

                    'oVent.downstand = CSng(txtDownstand.Text)
                    oVent.downstand = RoomHeight(oVent.fromroom) - oVent.height - oVent.sillheight
                    If oVent.downstand <= 0 Then
                        If oVent.cd < 1 Then
                            MsgBox("Please check the flow coefficient. If the top of the vent is flush with the compartment ceiling, then typically a flow coefficient of 1.0 would apply.", MsgBoxStyle.Information)
                        End If
                    End If

                    For Each x In oventdistributions
                        If x.id = oVent.id Then
                            If x.varname = "height" Then
                                x.varvalue = txtVentHeight.Text
                            End If
                            If x.varname = "width" Then
                                x.varvalue = txtVentWidth.Text
                            End If
                            If x.varname = "prob" Then
                                x.varvalue = txtVentProb.Text
                            End If
                            If x.varname = "HOreliability" Then
                                x.varvalue = txtHOreliability.Text
                            End If
                        End If
                    Next

                End If
            Next

            If oVent IsNot Nothing Then
                VentDB.SaveVents(oVents, oventdistributions)
                frmVentList.FillVentList()

                Call frmInputs.wallventarrays(oVents, oventdistributions, 0)
            End If

            Me.Close()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in frmNewVents.vb cmdSave_Click")

        End Try
    End Sub

    Private Sub cmdDist_vwidth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_vwidth.Click
       
        oVents = VentDB.GetVents
        oventdistributions = VentDB.GetVentDistributions()

        ventid = Me.Tag - 1
        Dim ovent As oVent = oVents(ventid)
        Dim paramdist As String
        Dim param As String = "width"
        Dim units As String = "m"
        Dim instruction As String = "vent width"

        Call frmDistParam.ShowVentDistributionForm(param, units, instruction, ovent, oventdistributions, ventid, paramdist)

        If paramdist <> "None" Then
            txtVentWidth.BackColor = Color.LightSalmon
        Else
            txtVentWidth.BackColor = Color.White
        End If

        txtVentWidth.Text = ovent.width

    End Sub

    Private Sub cmdDist_vheight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_vheight.Click

        ovents = VentDB.GetVents
        oventdistributions = VentDB.GetVentDistributions()

        ventid = Me.Tag - 1
        Dim ovent As oVent = ovents(ventid)
        Dim paramdist As String
        Dim param As String = "height"
        Dim units As String = "m"
        Dim instruction As String = "vent height"

        Call frmDistParam.ShowVentDistributionForm(param, units, instruction, ovent, oventdistributions, ventid, paramdist)
        If paramdist <> "None" Then
            txtVentHeight.BackColor = Color.LightSalmon
        Else
            txtVentHeight.BackColor = Color.White
        End If

        txtVentHeight.Text = ovent.height
    End Sub

    Private Sub cmdOpenOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOpenOptions.Click

        oVents = VentDB.GetVents
        oventdistributions = VentDB.GetVentDistributions
        frmVentOpenOptions.Text = "Wall Vent Opening Options"
        oVent = oVents(Me.Tag - 1)
        Call frmVentOpenOptions.ventopenoptions(oVent, oVents, oventdistributions)

        If frmVentOpenOptions.chkAutoVentOpen.Checked = True Then
            frmVentOpenOptions.GroupBox1.Visible = True
            frmVentOpenOptions.GroupBox2.Visible = True
            frmVentOpenOptions.GroupBox3.Visible = True
            frmVentOpenOptions.GroupBox4.Visible = True
            frmVentOpenOptions.GroupBox5.Visible = True
            frmVentOpenOptions.GroupBox6.Visible = True
            frmVentOpenOptions.GroupBox7.Visible = True
            frmVentOpenOptions.GroupBox8.Visible = True
            frmVentOpenOptions.GroupBox9.Visible = True
        Else
            frmVentOpenOptions.GroupBox1.Visible = False
            frmVentOpenOptions.GroupBox2.Visible = False
            frmVentOpenOptions.GroupBox3.Visible = False
            frmVentOpenOptions.GroupBox4.Visible = False
            frmVentOpenOptions.GroupBox5.Visible = False
            frmVentOpenOptions.GroupBox6.Visible = False
            frmVentOpenOptions.GroupBox7.Visible = False
            frmVentOpenOptions.GroupBox8.Visible = False
            frmVentOpenOptions.GroupBox9.Visible = False
        End If

    End Sub

 
    Private Sub cmdGlassProperties_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGlassProperties.Click

        oVents = VentDB.GetVents
        oventdistributions = VentDB.GetVentDistributions
        oVent = oVents(Me.Tag - 1)
        Call frmGlassBreak.glass_properties(oVent, oVents, oventdistributions)

    End Sub


    Private Sub txtVentWidth_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVentWidth.Validated
        ErrorProvider1.Clear()
        oVent.width = CDbl(txtVentWidth.Text)

    End Sub

    Private Sub txtVentWidth_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtVentWidth.Validating
        If IsNothing(oVents) Then
            oVents = VentDB.GetVents
        End If
        If IsNumeric(txtVentWidth.Text) Then
            If (CDbl(txtVentWidth.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtVentWidth.Select(0, txtVentWidth.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtVentWidth, "Invalid Entry. Vent width must be greater than 0 m.")

    End Sub

    Private Sub txtVentHeight_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVentHeight.Validated
        ErrorProvider1.Clear()
        oVent.height = CDbl(txtVentHeight.Text)
    End Sub

    Private Sub txtVentHeight_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtVentHeight.Validating
        If IsNothing(oVents) Then
            oVents = VentDB.GetVents
        End If
        If IsNumeric(txtVentHeight.Text) Then
            If (CDbl(txtVentHeight.Text) > 0) Then
                'If (CDbl(txtVentHeight.Text) <= RoomHeight(oVents(Me.Tag - 1).fromroom)) Then
                'okay
                Exit Sub
                'End If
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtVentHeight.Select(0, txtVentHeight.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtVentHeight, "Invalid Entry. Height width must be greater than 0 m and less than the room height.")

    End Sub

    Private Sub txtVentSillHeight_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVentSillHeight.Validated
        ErrorProvider1.Clear()
        oVent.sillheight = CDbl(txtVentSillHeight.Text)
    End Sub

    Private Sub txtVentSillHeight_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtVentSillHeight.Validating
        If IsNothing(oVents) Then Exit Sub
        If IsNumeric(txtVentSillHeight.Text) Then
            If (CDbl(txtVentSillHeight.Text) >= 0) Then
                If (CDbl(txtVentSillHeight.Text) <= RoomHeight(lstRoom1.SelectedItem)) Then
                    'okay
                    Exit Sub
                End If
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtVentSillHeight.Select(0, txtVentSillHeight.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtVentSillHeight, "Invalid Entry. Sill height must be greater or equal to 0 m and less than the room height.")

    End Sub

    Private Sub txtVentOpenTime_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVentOpenTime.Validated
        ErrorProvider1.Clear()
        oVent.opentime = CDbl(txtVentOpenTime.Text)
    End Sub

    Private Sub txtVentOpenTime_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtVentOpenTime.Validating
        If IsNumeric(txtVentOpenTime.Text) Then
            If (CDbl(txtVentOpenTime.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtVentOpenTime.Select(0, txtVentOpenTime.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtVentOpenTime, "Invalid Entry. Time must be greater or equal to 0 sec.")

    End Sub

    Private Sub txtVentCloseTime_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVentCloseTime.Validated
        ErrorProvider1.Clear()
        oVent.closetime = CDbl(txtVentCloseTime.Text)
    End Sub

    Private Sub txtVentCloseTime_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtVentCloseTime.Validating
        If IsNumeric(txtVentCloseTime.Text) Then
            If (CDbl(txtVentCloseTime.Text) >= oVent.opentime) Then
                'okay
                Exit Sub
            ElseIf CDbl(txtVentCloseTime.Text) = 0 Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtVentCloseTime.Select(0, txtVentCloseTime.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtVentCloseTime, "Invalid Entry. Time must be greater or equal to 0 sec.")

    End Sub

    Private Sub txtVentCD_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVentCD.Validated
        ErrorProvider1.Clear()
        oVent.cd = CDbl(txtVentCD.Text)
    End Sub

    Private Sub txtVentCD_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtVentCD.Validating
        If IsNumeric(txtVentCD.Text) Then
            If (CDbl(txtVentCD.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtVentCD.Select(0, txtVentCD.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtVentCD, "Invalid Entry. Discharge coeeficient must be greater or equal to 0.")

    End Sub

    Private Sub txtVentProb_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVentProb.Validated
        ErrorProvider1.Clear()
        oVent.probventclosed = CSng(txtVentProb.Text)
    End Sub

    Private Sub txtVentProb_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtVentProb.Validating
        If IsNumeric(txtVentProb.Text) Then
            If (CSng(txtVentProb.Text) >= 0) And (CSng(txtVentProb.Text) <= 1) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtVentProb.Select(0, txtVentProb.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtVentProb, "Invalid Entry. Probability that the vent is initially closed must be in the range 0 to 1.")

    End Sub

    Private Sub cmdDist_vprob_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_vprob.Click
        oVents = VentDB.GetVents
        oventdistributions = VentDB.GetVentDistributions()

        ventid = Me.Tag - 1
        Dim ovent As oVent = oVents(ventid)
        Dim paramdist As String
        Dim param As String = "prob"
        Dim units As String = "-"
        Dim instruction As String = "probability vent is initially closed"

        Call frmDistParam.ShowVentDistributionForm(param, units, instruction, ovent, oventdistributions, ventid, paramdist)
        If paramdist <> "None" Then
            txtVentProb.BackColor = Color.LightSalmon
        Else
            txtVentProb.BackColor = Color.White
        End If

        txtVentProb.Text = ovent.probventclosed
    End Sub


    Private Sub cmdVentGeometry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdVentGeometry.Click
        '=========================================================
        '   show a new form for describing the room-vent geometry
        '=========================================================

        Dim r1, R2 As Integer
        Dim vid As Integer

        If NumberRooms > 1 Then
            'If lstVentID.SelectedIndex < 0 Then
            '    MsgBox("You have not added a vent between these two rooms yet!", MsgBoxStyle.OkOnly)
            '    Exit Sub
            'End If

            'r1 = lstRoom1.SelectedIndex + 1
            'R2 = lstRoom2.SelectedIndex + 1 + r1
            'vid = lstVentID.SelectedIndex + 1

            If R2 <= NumberRooms Then
                'show blank form
                'frmBlank.Show()

                'set control properties
                'frmVentGeom.lblRoom1.Text = "Room " + CStr(lstRoom1.SelectedIndex + 1) + ": wall length = " + CStr(RoomWidth(lstRoom1.SelectedIndex + 1)) + " m"
                'frmVentGeom.lblRoom2.Text = "Room " + CStr(lstRoom2.SelectedIndex + 2 + lstRoom1.SelectedIndex) + ": wall length = " + CStr(RoomWidth(lstRoom2.SelectedIndex + 2 + lstRoom1.SelectedIndex)) + " m"
                'frmVentGeom.lblRoom1.Text = "Room " & CStr(r1) & ": wall length = " & CStr(WallLength1(r1, R2, vid)) & " m"
                'frmVentGeom.lblRoom2.Text = "Room " & CStr(R2) & ": wall length = " & CStr(WallLength2(r1, R2, vid)) & " m"

                Dim Wall1 As New Object
                'Wall1 = frmVentGeom.shpWall(0)

                'set width of wall in first room
                'Wall1.Height = RoomWidth(lstRoom1.SelectedIndex + 1) * 1000
                'frmVentGeom.shpWall(1).Height = RoomWidth(lstRoom2.SelectedIndex + 2 + lstRoom1.SelectedIndex) * 1000
                'frmVentGeom.shpWall(0).Height = VB6.TwipsToPixelsY(WallLength1(r1, R2, vid) * 1000)
                'frmVentGeom.shpWall(1).Height = VB6.TwipsToPixelsY(WallLength2(r1, R2, vid) * 1000)

                'frmVentGeom.shpVent.Height = VB6.TwipsToPixelsY(VentWidth(r1, R2, vid) * 1000)
                'frmVentGeom.shpVent.FillColor = System.Drawing.Color.White
                'Wall1.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmVentGeom.Width) / 2), VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(frmVentGeom.Height) - VB6.PixelsToTwipsY(frmVentGeom.shpWall(0).Height)) / 2), frmVentGeom.shpWall(0).Width, frmVentGeom.shpWall(0).Height)
                'frmVentGeom.shpWall(1).SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmVentGeom.Width) / 2 + VB6.PixelsToTwipsX(frmVentGeom.shpWall(0).Width)), VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(frmVentGeom.Height) - VB6.PixelsToTwipsY(frmVentGeom.shpWall(1).Height)) / 2), frmVentGeom.shpWall(1).Width, frmVentGeom.shpWall(1).Height)
                'frmVentGeom.shpVent.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmVentGeom.Width) / 2), VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(frmVentGeom.Height) - VB6.PixelsToTwipsY(frmVentGeom.shpVent.Height)) / 2), frmVentGeom.shpVent.Width, frmVentGeom.shpVent.Height)

                'frmVentGeom.Show()

            Else
                Exit Sub
            End If
        End If
    End Sub

    Private Sub txtBalconyWidth_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBalconyWidth.Validated
        ErrorProvider1.Clear()
        oVent.spillbalconyprojection = CSng(txtBalconyWidth.Text)
    End Sub

    Private Sub txtBalconyWidth_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtBalconyWidth.Validating
        Dim bmax As Double = 0
        If IsNumeric(txtBalconyWidth.Text) Then
            If (CDbl(txtBalconyWidth.Text) >= 0) Then
                If CDbl(txtVentWidth.Text) >= 2 * CDbl(txtBalconyWidth.Text) Then
                    'okay
                    Exit Sub

                Else
                    If opt3Dunchanbalcony.Checked = True Then
                        bmax = CDbl(txtVentWidth.Text) / 2

                    Else
                        'okay
                        Exit Sub
                    End If
                End If

            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtBalconyWidth.Select(0, txtBalconyWidth.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtBalconyWidth, "Invalid Entry. Balcony projection depth must be not more than " & bmax.ToString & " m for the specified vent width.")

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        MDIFrmMain.ViewToolStripMenuItem.PerformClick()

    End Sub


    Private Sub txtDownstand_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDownstand.Validated
        ErrorProvider1.Clear()
        oVent.downstand = CSng(txtDownstand.Text)
    End Sub

    Private Sub txtDownstand_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDownstand.Validating
        If IsNumeric(txtDownstand.Text) Then
            If (CSng(txtDownstand.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtDownstand.Select(0, txtDownstand.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtDownstand, "Invalid Entry. Downstand height must be > 0m.")

    End Sub


    Private Sub opt3Dchanbalcony_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles opt3Dchanbalcony.CheckedChanged
        If Me.opt3Dchanbalcony.Checked = True Then
            txtDownstand.Enabled = True
            txtBalconyWidth.Enabled = True
        End If
    End Sub

    Private Sub opt3Dunchanbalcony_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles opt3Dunchanbalcony.CheckedChanged
        If Me.opt3Dunchanbalcony.Checked = True Then
            txtDownstand.Enabled = True
            txtBalconyWidth.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        oVents = VentDB.GetVents
        oventdistributions = VentDB.GetVentDistributions()

        ventid = Me.Tag - 1
        Dim ovent As oVent = oVents(ventid)
        Dim paramdist As String = ""
        Dim param As String = "HOreliability"
        Dim units As String = "-"
        Dim instruction As String = "probability the hold open device is effective"

        Call frmDistParam.ShowVentDistributionForm(param, units, instruction, ovent, oventdistributions, ventid, paramdist)
        If paramdist <> "None" Then
            txtHOreliability.BackColor = Color.LightSalmon
        Else
            txtHOreliability.BackColor = Color.White
        End If

        txtHOreliability.Text = ovent.horeliability
    End Sub



    Private Sub txtHOreliability_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHOreliability.Validated
        ErrorProvider1.Clear()
        oVent.horeliability = CSng(txtHOreliability.Text)
    End Sub

    Private Sub txtHOreliability_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtHOreliability.Validating
        If IsNumeric(txtHOreliability.Text) Then
            If (CSng(txtHOreliability.Text) >= 0) And (CSng(txtHOreliability.Text) <= 1) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtHOreliability.Select(0, txtHOreliability.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtHOreliability, "Invalid Entry. Reliability of the hold open device must be in the range 0 to 1.")

    End Sub

    Private Sub lstRoom1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstRoom1.SelectedIndexChanged

        lstRoom2.Items.Clear()
        For i = lstRoom1.SelectedItem To NumberRooms
            If i = NumberRooms Then
                lstRoom2.Items.Add("outside")
            Else
                lstRoom2.Items.Add(i + 1)
            End If

        Next
        If lstRoom1.SelectedItem < NumberRooms Then
            lstRoom2.SelectedItem = lstRoom1.SelectedItem + 1
        Else
            lstRoom2.SelectedItem = "outside"
        End If


    End Sub



    Private Sub optStandard_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optStandard.CheckedChanged
        If Me.optStandard.Checked = True Then
            txtDownstand.Enabled = False
            txtBalconyWidth.Enabled = False
        End If
    End Sub

    Private Sub opt2Dadhered_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles opt2Dadhered.CheckedChanged
        If Me.opt2Dadhered.Checked = True Then
            txtDownstand.Enabled = True
            txtBalconyWidth.Enabled = True
        End If
    End Sub

    Private Sub opt3Dadhered_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles opt3Dadhered.CheckedChanged
        If Me.opt3Dadhered.Checked = True Then
            txtDownstand.Enabled = True
            txtBalconyWidth.Enabled = True
        End If
    End Sub

    Private Sub opt2Dbalcony_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles opt2Dbalcony.CheckedChanged
        If Me.opt2Dbalcony.Checked = True Then
            txtDownstand.Enabled = True
            txtBalconyWidth.Enabled = True
        End If
    End Sub

    Private Sub frmNewVents_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AddOwnedForm(frmVentOpenOptions)
        AddOwnedForm(frmGlassBreak)
        AddOwnedForm(frmFireResistance)

        If VM2 = True Then
            'do not change these inputs if using VM2 mode
            txtVentProb.Enabled = False
            txtHOreliability.Enabled = False
        Else
            txtVentProb.Enabled = True
            txtHOreliability.Enabled = True
        End If
    End Sub

    Private Sub cmdFireResistance_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFireResistance.Click

        oVents = VentDB.GetVents
        oventdistributions = VentDB.GetVentDistributions
        oVent = oVents(Me.Tag - 1)
        frmFireResistance.FRoptions(oVent, oVents, oventdistributions)
        If oVent.FRcriteria = -1 Then
            frmFireResistance.cboFRcriterion.SelectedIndex = 0
        Else
            frmFireResistance.cboFRcriterion.SelectedIndex = oVent.FRcriteria
        End If
    End Sub

End Class