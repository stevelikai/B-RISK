Imports System.Collections.Generic
Public Class frmVentOpenOptions
    Dim id, idr, idc, id1, id2 As Integer
    Dim oVents As List(Of oVent)
    Dim oventdistributions As List(Of oDistribution)
    Dim oVent As New oVent
    Dim ocVents As List(Of oCVent)
    Dim ocventdistributions As List(Of oDistribution)
    Dim ocVent As New oCVent

    Private Sub cboSDroom_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSDroom.SelectedIndexChanged
        Try

       

            If Me.Text = "Ceiling Vent Opening Options" Then
                SDtriggerroom2(idr, idc, id) = Me.cboSDroom.SelectedIndex + 1
                ocVent.sdtriggerroom = Me.cboSDroom.SelectedIndex + 1
                txtHODlabel.Text = "Device closes the vent " & ocVent.triggerventopendelay & " sec after when any smoke detector responds in room " & ocVent.sdtriggerroom

            Else
                SDtriggerroom(idr, idc, id) = Me.cboSDroom.SelectedIndex + 1
                oVent.sdtriggerroom = Me.cboSDroom.SelectedIndex + 1
                txtHODlabel.Text = "Device closes the vent " & oVent.triggerventopendelay & " sec after when any smoke detector responds in room " & oVent.sdtriggerroom

            End If

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in frmVentOpenOptions.vb cboSDroom_SelectedIndexChanged")
        End Try

    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Try

   
            If id < 1 Then Exit Sub

            If Me.Text = "Ceiling Vent Opening Options" Then

                ocVents = VentDB.GetCVents
                ocventdistributions = VentDB.GetCVentDistributions

                'ReDim Preserve trigger_device2(0 To 6, MaxNumberRooms + 1, MaxNumberRooms + 1, ocVents.Count)
                'ReDim Preserve AutoOpenCVent(MaxNumberRooms + 1, MaxNumberRooms + 1, ocVents.Count)

                If chkAutoVentOpen.Checked = True Then
                    AutoOpenCVent(idr, idc, id) = True
                    ocVents(id1 - 1).autoopenvent = True
                Else
                    AutoOpenCVent(idr, idc, id) = False
                    ocVents(id1 - 1).autoopenvent = False
                End If

                If chkHRR.Checked = True Then
                    trigger_device2(2, idr, idc, id) = True
                    ocVents(id1 - 1).triggerHRR = True
                Else
                    trigger_device2(2, idr, idc, id) = False
                    ocVents(id1 - 1).triggerHRR = False
                End If

                If chkVL.Checked = True Then
                    trigger_device2(4, idr, idc, id) = True
                    ocVents(id1 - 1).triggerVL = True
                Else
                    trigger_device2(4, idr, idc, id) = False
                    ocVents(id1 - 1).triggerVL = False
                End If

                If chkHoldOpen.Checked = True Then
                    trigger_device2(5, idr, idc, id) = True
                    ocVents(id1 - 1).triggerHO = True
                    chkSD.Checked = False
                Else
                    trigger_device2(5, idr, idc, id) = False
                    ocVents(id1 - 1).triggerHO = False
                End If

                If chkFO.Checked = True Then
                    trigger_device2(3, idr, idc, id) = True
                    ocVents(id1 - 1).triggerFO = True
                Else
                    trigger_device2(3, idr, idc, id) = False
                    ocVents(id1 - 1).triggerFO = False
                End If

                If chkHD.Checked = True Then
                    trigger_device2(1, idr, idc, id) = True
                    ocVents(id1 - 1).triggerHD = True
                Else
                    trigger_device2(1, idr, idc, id) = False
                    ocVents(id1 - 1).triggerHD = False
                End If

                If chkSD.Checked = True Then
                    trigger_device2(0, idr, idc, id) = True
                    ocVents(id1 - 1).triggerSD = True
                Else
                    trigger_device2(0, idr, idc, id) = False
                    ocVents(id1 - 1).triggerSD = False
                End If

                SDtriggerroom2(idr, idc, id) = cboSDroom.SelectedIndex + 1
                ocVents(id1 - 1).sdtriggerroom = SDtriggerroom2(idr, idc, id)

                ocVents(id1 - 1).HRRthreshold = txtFireSize.Text
                ocVents(id1 - 1).HRRventopendelay = txtHRROpenDelay.Text
                ocVents(id1 - 1).HRRventopenduration = txtHRRTimeVentOpen.Text

                ocVents(id1 - 1).triggerventopendelay = txtTriggerOpenDelay.Text
                ocVents(id1 - 1).triggerventopenduration = txtTriggerTimeVentOpen.Text

                If ocVents IsNot Nothing Then
                    VentDB.SaveCVents(ocVents, ocventdistributions)
                End If

            Else
                'wall vents

                oVents = VentDB.GetVents
                oventdistributions = VentDB.GetVentDistributions

                If chkAutoVentOpen.Checked = True Then
                    AutoOpenVent(idr, idc, id) = True
                    oVents(id1 - 1).autoopenvent = True
                Else
                    AutoOpenVent(idr, idc, id) = False
                    oVents(id1 - 1).autoopenvent = False
                End If

                If chkHRR.Checked = True Then
                    trigger_device(2, idr, idc, id) = True
                    oVents(id1 - 1).triggerHRR = True
                Else
                    trigger_device(2, idr, idc, id) = False
                    oVents(id1 - 1).triggerHRR = False
                End If

                If chkVL.Checked = True Then
                    trigger_device(4, idr, idc, id) = True
                    oVents(id1 - 1).triggerVL = True
                Else
                    trigger_device(4, idr, idc, id) = False
                    oVents(id1 - 1).triggerVL = False
                End If

                If chkHoldOpen.Checked = True Then
                    trigger_device(5, idr, idc, id) = True
                    oVents(id1 - 1).triggerHO = True
                    chkSD.Checked = False
                Else
                    trigger_device(5, idr, idc, id) = False
                    oVents(id1 - 1).triggerHO = False
                End If

                If chkFO.Checked = True Then
                    trigger_device(3, idr, idc, id) = True
                    oVents(id1 - 1).triggerFO = True
                Else
                    trigger_device(3, idr, idc, id) = False
                    oVents(id1 - 1).triggerFO = False
                End If

                If chkHD.Checked = True Then
                    trigger_device(1, idr, idc, id) = True
                    oVents(id1 - 1).triggerHD = True
                Else
                    trigger_device(1, idr, idc, id) = False
                    oVents(id1 - 1).triggerHD = False
                End If

                If chkSD.Checked = True Then
                    trigger_device(0, idr, idc, id) = True
                    oVents(id1 - 1).triggerSD = True
                Else
                    trigger_device(0, idr, idc, id) = False
                    oVents(id1 - 1).triggerSD = False
                End If

                SDtriggerroom(idr, idc, id) = cboSDroom.SelectedIndex + 1
                oVents(id1 - 1).sdtriggerroom = SDtriggerroom(idr, idc, id)

                oVents(id1 - 1).HRRthreshold = txtFireSize.Text
                oVents(id1 - 1).HRRventopendelay = txtHRROpenDelay.Text
                oVents(id1 - 1).HRRventopenduration = txtHRRTimeVentOpen.Text

                oVents(id1 - 1).triggerventopendelay = txtTriggerOpenDelay.Text
                oVents(id1 - 1).triggerventopenduration = txtTriggerTimeVentOpen.Text

                If oVents IsNot Nothing Then
                    VentDB.SaveVents(oVents, oventdistributions)
                End If
            End If

            Me.Hide()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in frmVentOpenOptions.vb_cmdclose")
            Me.Close()
        End Try
    End Sub
    Public Sub ventopenoptions2(ByRef ocVent, ByRef OcVents, ByRef OcVentdistributions)
        Try

            'ReDim Preserve trigger_device2(0 To 6, MaxNumberRooms + 1, MaxNumberRooms + 1, OcVents.Count)

            id1 = ocVent.id
            idr = ocVent.upperroom
            idc = ocVent.lowerroom

            id2 = 0
            For Each ocVent2 In OcVents
                If ocVent2.upperroom = idr And ocVent2.lowerroom = idc Then
                    id2 = id2 + 1
                    If ocVent2.id = id1 Then Exit For
                End If
            Next

            id = id2 'this is the id to use in old ventarrays

            If NumberRooms > 0 Then
                cboSDroom.Items.Clear()
                For i = 1 To NumberRooms
                    cboSDroom.Items.Add(CStr(i))
                Next i

                cboSDroom.SelectedIndex = ocVent.sdtriggerroom - 1
                txtFireSize.Text = ocVent.HRRthreshold
                txtHRROpenDelay.Text = ocVent.HRRventopendelay
                txtHRRTimeVentOpen.Text = ocVent.HRRventopenduration
                txtTriggerOpenDelay.Text = ocVent.triggerventopendelay
                txtTriggerTimeVentOpen.Text = ocVent.triggerventopenduration

                If ocVent.AutoOpenVent = True Then
                    chkAutoVentOpen.Checked = True
                Else
                    chkAutoVentOpen.Checked = False
                End If

                'SD
                If trigger_device2(0, idr, idc, id) = True Then
                    chkSD.Checked = True
                Else
                    chkSD.Checked = False
                End If

                'HD
                If trigger_device2(1, idr, idc, id) = True Then
                    chkHD.Checked = True
                Else
                    chkHD.Checked = False
                End If

                'fire size
                If trigger_device2(2, idr, idc, id) = True Then
                    chkHRR.Checked = True
                Else
                    chkHRR.Checked = False
                End If

                'flashover
                If trigger_device2(3, idr, idc, id) = True Then
                    chkFO.Checked = True
                Else
                    chkFO.Checked = False
                End If

                'ventilation limit
                If trigger_device2(4, idr, idc, id) = True Then
                    chkVL.Checked = True
                Else
                    chkVL.CheckState = False
                End If

                'hold open fitted
                If trigger_device2(5, idr, idc, id) = True Then
                    chkHoldOpen.Checked = True
                Else
                    chkHoldOpen.Checked = False
                End If

            End If

            Me.Show()
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in frmVentOpenOptions.vb_ventopenoptions2")

        End Try
    End Sub
    Public Sub ventopenoptions(ByRef oVent, ByRef OVents, ByRef OVentdistributions)

        id1 = oVent.id
        idr = oVent.fromroom
        idc = oVent.toroom

        id2 = 0
        For Each oVent2 In OVents
            If oVent2.fromroom = idr And oVent2.toroom = idc Then
                id2 = id2 + 1
                If oVent2.id = id1 Then Exit For
            End If
        Next

        id = id2 'this is the id to use in old ventarrays

        If NumberRooms > 0 Then
            cboSDroom.Items.Clear()
            For i = 1 To NumberRooms
                cboSDroom.Items.Add(CStr(i))
            Next i

            cboSDroom.SelectedIndex = oVent.sdtriggerroom - 1
            txtFireSize.Text = oVent.HRRthreshold
            txtHRROpenDelay.Text = oVent.HRRventopendelay
            txtHRRTimeVentOpen.Text = oVent.HRRventopenduration
            txtTriggerOpenDelay.Text = oVent.triggerventopendelay
            txtTriggerTimeVentOpen.Text = oVent.triggerventopenduration

            If oVent.AutoOpenVent = True Then
                chkAutoVentOpen.Checked = True
            Else
                chkAutoVentOpen.Checked = False
            End If

            'SD
            If trigger_device(0, idr, idc, id) = True Then
                chkSD.Checked = True
            Else
                chkSD.Checked = False
            End If

            'HD
            If trigger_device(1, idr, idc, id) = True Then
                chkHD.Checked = True
            Else
                chkHD.Checked = False
            End If

            'fire size
            If trigger_device(2, idr, idc, id) = True Then
                chkHRR.Checked = True
            Else
                chkHRR.Checked = False
            End If

            'flashover
            If trigger_device(3, idr, idc, id) = True Then
                chkFO.Checked = True
            Else
                chkFO.Checked = False
            End If

            'ventilation limit
            If trigger_device(4, idr, idc, id) = True Then
                chkVL.Checked = True
            Else
                chkVL.CheckState = False
            End If

            'hold open fitted
            If trigger_device(5, idr, idc, id) = True Then
                chkHoldOpen.Checked = True
            Else
                chkHoldOpen.Checked = False
            End If

        End If

        Me.Show()

    End Sub

    Private Sub txtFireSize_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFireSize.Validated
        ErrorProvider1.Clear()
        If Me.Text = "Ceiling Vent Opening Options" Then
            HRR_threshold2(idr, idc, id) = CDbl(Me.txtFireSize.Text)
            HRR_threshold2(idc, idr, id) = CDbl(Me.txtFireSize.Text)
            ocVent.HRRthreshold = CDbl(Me.txtFireSize.Text)
        Else
            HRR_threshold(idr, idc, id) = CDbl(Me.txtFireSize.Text)
            HRR_threshold(idc, idr, id) = CDbl(Me.txtFireSize.Text)
            oVent.HRRthreshold = CDbl(Me.txtFireSize.Text)
        End If
    End Sub

    Private Sub txtFireSize_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtFireSize.Validating
        If IsNumeric(txtFireSize.Text) Then
            If (CDbl(txtFireSize.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtFireSize.Select(0, txtFireSize.Text.Length)

        ' Give the ErrorProvider the error message to display.
        ErrorProvider1.SetError(txtFireSize, "Invalid Entry. Fire size must be greater or equal to 0 kW.")
    End Sub

    Private Sub txtTriggerTimeVentOpen_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTriggerTimeVentOpen.Validated
        ErrorProvider1.Clear()
        If Me.Text = "Ceiling Vent Opening Options" Then
            trigger_ventopenduration2(idr, idc, id) = CDbl(Me.txtTriggerTimeVentOpen.Text)
            trigger_ventopenduration2(idc, idr, id) = CDbl(Me.txtTriggerTimeVentOpen.Text)
            ocVent.triggerventopenduration = CDbl(Me.txtTriggerTimeVentOpen.Text)
        Else
            trigger_ventopenduration(idr, idc, id) = CDbl(Me.txtTriggerTimeVentOpen.Text)
            trigger_ventopenduration(idc, idr, id) = CDbl(Me.txtTriggerTimeVentOpen.Text)
            oVent.triggerventopenduration = CDbl(Me.txtTriggerTimeVentOpen.Text)
        End If
    End Sub

    Private Sub txtTriggerTimeVentOpen_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtTriggerTimeVentOpen.Validating
        If IsNumeric(txtTriggerTimeVentOpen.Text) Then
            If (CDbl(txtTriggerTimeVentOpen.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtTriggerTimeVentOpen.Select(0, txtTriggerTimeVentOpen.Text.Length)

        ' Give the ErrorProvider the error message to display.
        ErrorProvider1.SetError(txtTriggerTimeVentOpen, "Invalid Entry. Time period must be greater or equal to 0 sec.")

    End Sub

    Private Sub txtHRROpenDelay_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHRROpenDelay.Validated
        ErrorProvider1.Clear()
        If Me.Text = "Ceiling Vent Opening Options" Then
            HRR_ventopendelay2(idr, idc, id) = CSng(Me.txtHRROpenDelay.Text)
            HRR_ventopendelay2(idc, idr, id) = CSng(Me.txtHRROpenDelay.Text)
            ocVent.HRRventopendelay = CSng(Me.txtHRROpenDelay.Text)
        Else
            HRR_ventopendelay(idr, idc, id) = CSng(Me.txtHRROpenDelay.Text)
            HRR_ventopendelay(idc, idr, id) = CSng(Me.txtHRROpenDelay.Text)
            oVent.HRRventopendelay = CSng(Me.txtHRROpenDelay.Text)
        End If
       

    End Sub

    Private Sub txtHRROpenDelay_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtHRROpenDelay.Validating
        If IsNumeric(txtHRROpenDelay.Text) Then
            If (CDbl(txtHRROpenDelay.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtHRROpenDelay.Select(0, txtHRROpenDelay.Text.Length)

        ' Give the ErrorProvider the error message to display.
        ErrorProvider1.SetError(txtHRROpenDelay, "Invalid Entry. Time period must be greater or equal to 0 sec.")

    End Sub

    Private Sub txtTriggerOpenDelay_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTriggerOpenDelay.Validated
        ErrorProvider1.Clear()
        If Me.Text = "Ceiling Vent Opening Options" Then
            trigger_ventopendelay2(idr, idc, id) = CDbl(Me.txtTriggerOpenDelay.Text)
            trigger_ventopendelay2(idc, idr, id) = CDbl(Me.txtTriggerOpenDelay.Text)
            ocVent.triggerventopendelay = CDbl(Me.txtTriggerOpenDelay.Text)
        Else
            trigger_ventopendelay(idr, idc, id) = CDbl(Me.txtTriggerOpenDelay.Text)
            trigger_ventopendelay(idc, idr, id) = CDbl(Me.txtTriggerOpenDelay.Text)
            oVent.triggerventopendelay = CDbl(Me.txtTriggerOpenDelay.Text)
        End If

       
    End Sub

    Private Sub txtTriggerOpenDelay_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtTriggerOpenDelay.Validating
        If IsNumeric(txtTriggerOpenDelay.Text) Then
            If (CDbl(txtTriggerOpenDelay.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtTriggerOpenDelay.Select(0, txtTriggerOpenDelay.Text.Length)

        ' Give the ErrorProvider the error message to display.
        ErrorProvider1.SetError(txtTriggerOpenDelay, "Invalid Entry. Time period must be greater or equal to 0 sec.")

    End Sub

    Private Sub txtHRRTimeVentOpen_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHRRTimeVentOpen.Validated
        ErrorProvider1.Clear()
        If Me.Text = "Ceiling Vent Opening Options" Then
            HRR_ventopenduration2(idr, idc, id) = CDbl(Me.txtHRRTimeVentOpen.Text)
            HRR_ventopenduration2(idc, idr, id) = CDbl(Me.txtHRRTimeVentOpen.Text)
            ocVent.HRRventopenduration = CDbl(Me.txtHRRTimeVentOpen.Text)
        Else
            HRR_ventopenduration(idr, idc, id) = CDbl(Me.txtHRRTimeVentOpen.Text)
            HRR_ventopenduration(idc, idr, id) = CDbl(Me.txtHRRTimeVentOpen.Text)
            oVent.HRRventopenduration = CDbl(Me.txtHRRTimeVentOpen.Text)
        End If
  
           
    End Sub

    Private Sub txtHRRTimeVentOpen_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtHRRTimeVentOpen.Validating
        If IsNumeric(txtHRRTimeVentOpen.Text) Then
            If (CDbl(txtHRRTimeVentOpen.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtHRRTimeVentOpen.Select(0, txtHRRTimeVentOpen.Text.Length)

        ' Give the ErrorProvider the error message to display.
        ErrorProvider1.SetError(txtHRRTimeVentOpen, "Invalid Entry. Time period must be greater or equal to 0 sec.")

    End Sub

    Private Sub chkAutoVentOpen_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAutoVentOpen.CheckedChanged
        If chkAutoVentOpen.Checked = True Then
            GroupBox1.Visible = True
            GroupBox2.Visible = True
            GroupBox3.Visible = True
            GroupBox4.Visible = True
            GroupBox5.Visible = True
            GroupBox6.Visible = True
            GroupBox7.Visible = True
            GroupBox8.Visible = True
            If Me.Text = "Ceiling Vent Opening Options" Then
                GroupBox9.Visible = False
            Else
                GroupBox9.Visible = True
            End If

        Else
            GroupBox1.Visible = False
            GroupBox2.Visible = False
            GroupBox3.Visible = False
            GroupBox4.Visible = False
            GroupBox5.Visible = False
            GroupBox6.Visible = False
            GroupBox7.Visible = False
            GroupBox8.Visible = False
            GroupBox9.Visible = False
        End If
    End Sub

End Class