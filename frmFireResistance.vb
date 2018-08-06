Imports System.Collections.Generic
Public Class frmFireResistance
    Public ventid As Integer

    Dim oVents As List(Of oVent)
    Dim oventdistributions As List(Of oDistribution)
    Dim oVent As New oVent

    Dim ocVents As List(Of oCVent)
    Dim ocventdistributions As List(Of oDistribution)
    Dim ocVent As New oCVent

    Private Sub cmdDist_integrity_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_integrity.Click

        Dim paramdist As String = ""
        Dim param As String = "integrity"
        Dim units As String = "min"
        Dim instruction As String = "fire resistance integrity value"

        If Me.Tag <> "ceiling" Then

            oVents = VentDB.GetVents
            oventdistributions = VentDB.GetVentDistributions()

            Dim ovent As oVent = oVents(ventid - 1)

            Call frmDistParam.ShowVentDistributionForm(param, units, instruction, ovent, oventdistributions, ventid - 1, paramdist)

            txtFRintegrity.Text = ovent.integrity

        Else

            'for ceiling vents
            ocVents = VentDB.GetCVents
            ocventdistributions = VentDB.GetCVentDistributions()

            Dim ocvent As oCVent = ocVents(ventid - 1)
           
            Call frmDistParam.ShowCVentDistributionForm(param, units, instruction, ocvent, ocventdistributions, ventid - 1, paramdist)

            txtFRintegrity.Text = ocvent.integrity

        End If

        If paramdist <> "None" Then
            txtFRintegrity.BackColor = Color.LightSalmon
        Else
            txtFRintegrity.BackColor = Color.White
        End If

    End Sub


    Private Sub txtFRintegrity_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFRintegrity.Validated
        ErrorProvider1.Clear()

        If Me.Tag <> "ceiling" Then
            oVent.integrity = CSng(txtFRintegrity.Text)
        Else
            ocVent.integrity = CDbl(txtFRintegrity.Text)
        End If

    End Sub

    Private Sub txtFRintegrity_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtFRintegrity.Validating

        If IsNumeric(txtFRintegrity.Text) Then
            If (CDbl(txtFRintegrity.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtFRintegrity.Select(0, txtFRintegrity.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtFRintegrity, "Invalid Entry. Integrity must be greater than or equal to 0 min.")

    End Sub
    Public Sub FRoptions(ByRef oVent, ByRef OVents, ByRef OVentdistributions)

        Me.Tag = "wall"

        ventid = oVent.id

        If oVent.triggerFR = True Then
            chkFireBarrierModel.Checked = True 'will open this vent based on FR criterion
        Else
            chkFireBarrierModel.Checked = False
        End If

        txtFRintegrity.Text = oVent.integrity
        txtMaxOpening.Text = oVent.maxopening
        txtMaxOpeningTime.Text = oVent.maxopeningtime
        txtFRgastemp.Text = oVent.gastemp
        cboFRcriterion.SelectedIndex = -1
        cboFRcriterion.SelectedIndex = oVent.FRcriteria

        For Each oDistribution In OVentdistributions
            If oDistribution.id = oVent.id Then
                If oDistribution.varname = "integrity" Then
                    txtFRintegrity.Text = oDistribution.varvalue
                    If oDistribution.distribution <> "None" Then
                        txtFRintegrity.BackColor = distbackcolor
                    Else
                        txtFRintegrity.BackColor = distnobackcolor
                    End If
                End If
                If oDistribution.varname = "maxopening" Then
                    txtMaxOpening.Text = oDistribution.varvalue
                    If oDistribution.distribution <> "None" Then
                        txtMaxOpening.BackColor = distbackcolor
                    Else
                        txtMaxOpening.BackColor = distnobackcolor
                    End If
                End If
                If oDistribution.varname = "maxopeningtime" Then
                    txtMaxOpeningTime.Text = oDistribution.varvalue
                    If oDistribution.distribution <> "None" Then
                        txtMaxOpeningTime.BackColor = distbackcolor
                    Else
                        txtMaxOpeningTime.BackColor = distnobackcolor
                    End If
                End If
                If oDistribution.varname = "gastemp" Then
                    txtFRgastemp.Text = oDistribution.varvalue
                    If oDistribution.distribution <> "None" Then
                        txtFRgastemp.BackColor = distbackcolor
                    Else
                        txtFRgastemp.BackColor = distnobackcolor
                    End If
                End If
            End If
        Next


        Me.Show()

    End Sub
    Public Sub FRoptions2(ByRef ocVent, ByRef OcVents, ByRef OcVentdistributions)

        Me.Tag = "ceiling"

        ventid = ocVent.id

        If ocVent.triggerFR = True Then
            chkFireBarrierModel.Checked = True
        Else
            chkFireBarrierModel.Checked = False
        End If

        txtFRintegrity.Text = ocVent.integrity
        txtMaxOpening.Text = ocVent.maxopening
        txtMaxOpeningTime.Text = ocVent.maxopeningtime
        txtFRgastemp.Text = ocVent.gastemp
        cboFRcriterion.SelectedIndex = ocVent.FRcriteria

        For Each oDistribution In OcVentdistributions
            If oDistribution.id = ocVent.id Then
                If oDistribution.varname = "integrity" Then
                    txtFRintegrity.Text = oDistribution.varvalue
                    If oDistribution.distribution <> "None" Then
                        txtFRintegrity.BackColor = distbackcolor
                    Else
                        txtFRintegrity.BackColor = distnobackcolor
                    End If
                End If
                If oDistribution.varname = "maxopening" Then
                    txtMaxOpening.Text = oDistribution.varvalue
                    If oDistribution.distribution <> "None" Then
                        txtMaxOpening.BackColor = distbackcolor
                    Else
                        txtMaxOpening.BackColor = distnobackcolor
                    End If
                End If
                If oDistribution.varname = "maxopeningtime" Then
                    txtMaxOpeningTime.Text = oDistribution.varvalue
                    If oDistribution.distribution <> "None" Then
                        txtMaxOpeningTime.BackColor = distbackcolor
                    Else
                        txtMaxOpeningTime.BackColor = distnobackcolor
                    End If
                End If
                If oDistribution.varname = "gastemp" Then
                    txtFRgastemp.Text = oDistribution.varvalue
                    If oDistribution.distribution <> "None" Then
                        txtFRgastemp.BackColor = distbackcolor
                    Else
                        txtFRgastemp.BackColor = distnobackcolor
                    End If
                End If
            End If
        Next

        Me.Show()

    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Try

            If ventid < 1 Then Exit Sub

            If Me.Tag <> "ceiling" Then
                'wall vents
                oVents = VentDB.GetVents
                oventdistributions = VentDB.GetVentDistributions

                If IsNothing(trigger_device) Then
                    Call frmInputs.wallventarrays(oVents, oventdistributions, 0)
                End If

                Dim idr As Integer = oVents(ventid - 1).fromroom
                Dim idc As Integer = oVents(ventid - 1).toroom
                Dim id1 As Integer = oVents(ventid - 1).id

                Dim id As Integer = 0
                Dim id2 As Integer = 0
                For Each oVent2 In oVents
                    If oVent2.fromroom = idr And oVent2.toroom = idc Then
                        id2 = id2 + 1
                        If oVent2.id = id1 Then Exit For
                    End If
                Next

                id = id2 'this is the id to use in old ventarray

                If chkFireBarrierModel.Checked = True Then
                    trigger_device(6, idr, idc, id) = True
                    oVents(ventid - 1).triggerFR = True
                Else
                    trigger_device(6, idr, idc, id) = False
                    oVents(ventid - 1).triggerFR = False
                End If

                oVents(ventid - 1).integrity = txtFRintegrity.Text
                oVents(ventid - 1).maxopening = txtMaxOpening.Text
                oVents(ventid - 1).maxopeningtime = txtMaxOpeningTime.Text
                oVents(ventid - 1).gastemp = txtFRgastemp.Text
                oVents(ventid - 1).FRcriteria = cboFRcriterion.SelectedIndex

                For Each x In oventdistributions
                    If x.id = ventid Then
                        If x.varname = "integrity" Then
                            x.varvalue = txtFRintegrity.Text
                        End If
                        If x.varname = "maxopening" Then
                            x.varvalue = txtMaxOpening.Text
                        End If
                        If x.varname = "maxopeningtime" Then
                            x.varvalue = txtMaxOpeningTime.Text
                        End If
                        If x.varname = "gastemp" Then
                            x.varvalue = txtFRgastemp.Text
                        End If
                    End If
                Next

                If oVents IsNot Nothing Then
                    VentDB.SaveVents(oVents, oventdistributions)
                End If
            Else
                ocVents = VentDB.GetCVents
                ocventdistributions = VentDB.GetCVentDistributions

                If IsNothing(trigger_device2) Then
                    Call frmInputs.ceilingventarrays(ocVents, ocventdistributions, 0)
                End If

                Dim idr As Integer = ocVents(ventid - 1).upperroom
                Dim idc As Integer = ocVents(ventid - 1).lowerroom
                Dim id1 As Integer = ocVents(ventid - 1).id

                Dim id As Integer = 0
                Dim id2 As Integer = 0
                For Each ocVent2 In ocVents
                    If ocVent2.upperroom = idr And ocVent2.lowerroom = idc Then
                        id2 = id2 + 1
                        If ocVent2.id = id1 Then Exit For
                    End If
                Next

                id = id2 'this is the id to use in old ventarray

                'If IsNothing(trigger_device2) Then
                '    Resize_CVents()
                '    ReDim trigger_device2(0 To 6, MaxNumberRooms + 1, MaxNumberRooms + 1, ocVents.Count)
                'End If

                If chkFireBarrierModel.Checked = True Then
                    trigger_device2(6, idr, idc, id) = True
                    ocVents(ventid - 1).triggerFR = True
                Else
                    trigger_device2(6, idr, idc, id) = False
                    ocVents(ventid - 1).triggerFR = False
                End If

                ocVents(ventid - 1).FRcriteria = cboFRcriterion.SelectedIndex
                ocVents(ventid - 1).integrity = txtFRintegrity.Text
                ocVents(ventid - 1).maxopening = txtMaxOpening.Text
                ocVents(ventid - 1).maxopeningtime = txtMaxOpeningTime.Text
                ocVents(ventid - 1).gastemp = txtFRgastemp.Text

                For Each x In ocventdistributions
                    If x.id = ventid Then
                        If x.varname = "integrity" Then
                            x.varvalue = txtFRintegrity.Text
                        End If
                        If x.varname = "maxopening" Then
                            x.varvalue = txtMaxOpening.Text
                        End If
                        If x.varname = "maxopeningtime" Then
                            x.varvalue = txtMaxOpeningTime.Text
                        End If
                        If x.varname = "gastemp" Then
                            x.varvalue = txtFRgastemp.Text
                        End If

                    End If
                Next

                If ocVents IsNot Nothing Then
                    VentDB.SaveCVents(ocVents, ocventdistributions)
                End If
            End If

            Me.Close()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in frmFireResistance.vb cmdClose_Click")

        End Try

    End Sub

    Private Sub txtMaxOpening_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMaxOpening.Validated
        ErrorProvider1.Clear()

        If Me.Tag <> "ceiling" Then
            oVent.maxopening = CSng(txtMaxOpening.Text)
        Else
            ocVent.maxopening = CDbl(txtMaxOpening.Text)
        End If

    End Sub

    Private Sub txtMaxOpening_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtMaxOpening.Validating

        If IsNumeric(txtMaxOpening.Text) Then
            If (CDbl(txtMaxOpening.Text) >= 0) And (CDbl(txtMaxOpening.Text) <= 100) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtMaxOpening.Select(0, txtMaxOpening.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtMaxOpening, "Invalid Entry. Max opening area must be in the range 0 to 100%.")

    End Sub


    Private Sub txtMaxOpeningTime_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMaxOpeningTime.Validated
        If Me.Tag <> "ceiling" Then
            oVent.maxopeningtime = CSng(txtMaxOpeningTime.Text)
        Else
            ocVent.maxopeningtime = CDbl(txtMaxOpeningTime.Text)
        End If
    End Sub

    Private Sub txtMaxOpeningTime_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtMaxOpeningTime.Validating

        If IsNumeric(txtMaxOpening.Text) Then
            If (CDbl(txtMaxOpeningTime.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtMaxOpeningTime.Select(0, txtMaxOpeningTime.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtMaxOpeningTime, "Invalid Entry. Max opening time must be greater or equal to 0 s.")

    End Sub

    Private Sub cmdDist_maxopening_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_maxopening.Click

        Dim paramdist As String
        Dim param As String = "maxopening"
        Dim units As String = "%"
        Dim instruction As String = "maximum opening area"

        If Me.Tag <> "ceiling" Then
       
            oVents = VentDB.GetVents
            oventdistributions = VentDB.GetVentDistributions()

            Dim ovent As oVent = oVents(ventid - 1)

            Call frmDistParam.ShowVentDistributionForm(param, units, instruction, ovent, oventdistributions, ventid - 1, paramdist)

            txtMaxOpening.Text = ovent.maxopening
        Else
            ocVents = VentDB.GetCVents
            ocventdistributions = VentDB.GetCVentDistributions()

            Dim ocvent As oCVent = ocVents(ventid - 1)

            Call frmDistParam.ShowCVentDistributionForm(param, units, instruction, ocvent, ocventdistributions, ventid - 1, paramdist)

            txtMaxOpening.Text = ocvent.maxopening
        End If

        If paramdist <> "None" Then
            txtMaxOpening.BackColor = Color.LightSalmon
        Else
            txtMaxOpening.BackColor = Color.White
        End If

    End Sub

    Private Sub cmdDist_maxopeningtime_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_maxopeningtime.Click

        Dim paramdist As String
        Dim param As String = "maxopeningtime"
        Dim units As String = "s"
        Dim instruction As String = "maximum opening time"

        If Me.Tag <> "ceiling" Then

            oVents = VentDB.GetVents
            oventdistributions = VentDB.GetVentDistributions()

            Dim ovent As oVent = oVents(ventid - 1)

            Call frmDistParam.ShowVentDistributionForm(param, units, instruction, ovent, oventdistributions, ventid - 1, paramdist)

            txtMaxOpeningTime.Text = ovent.maxopeningtime
        Else
            ocVents = VentDB.GetCVents
            ocventdistributions = VentDB.GetCVentDistributions()

            Dim ocvent As oCVent = ocVents(ventid - 1)

            Call frmDistParam.ShowCVentDistributionForm(param, units, instruction, ocvent, ocventdistributions, ventid - 1, paramdist)

            txtMaxOpeningTime.Text = ocvent.maxopeningtime
        End If

        If paramdist <> "None" Then
            txtMaxOpeningTime.BackColor = Color.LightSalmon
        Else
            txtMaxOpeningTime.BackColor = Color.White
        End If

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub txtFRgastemp_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFRgastemp.Validated
        If Me.Tag <> "ceiling" Then
            oVent.gastemp = CSng(txtFRgastemp.Text)
        Else
            ocVent.gastemp = CDbl(txtFRgastemp.Text)
        End If

        ErrorProvider1.Clear()

    End Sub

    Private Sub txtFRgastemp_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtFRgastemp.Validating
        
        If IsNumeric(txtFRgastemp.Text) Then
            If (CDbl(txtFRgastemp.Text) >= 40) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtFRgastemp.Select(0, txtFRgastemp.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtFRgastemp, "Invalid Entry. Gas temperature must be greater or equal to 40 C.")

    End Sub

    Private Sub cmdDist_gastemp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_gastemp.Click

        Dim paramdist As String
        Dim param As String = "gastemp"
        Dim units As String = "C"
        Dim instruction As String = "gas temperature"

        If Me.Tag <> "ceiling" Then

            oVents = VentDB.GetVents
            oventdistributions = VentDB.GetVentDistributions()

            Dim ovent As oVent = oVents(ventid - 1)

            Call frmDistParam.ShowVentDistributionForm(param, units, instruction, ovent, oventdistributions, ventid - 1, paramdist)

            txtFRgastemp.Text = ovent.gastemp

        Else
            ocVents = VentDB.GetCVents
            ocventdistributions = VentDB.GetCVentDistributions()

            Dim ocvent As oCVent = ocVents(ventid - 1)

            Call frmDistParam.ShowCVentDistributionForm(param, units, instruction, ocvent, ocventdistributions, ventid - 1, paramdist)

            txtFRgastemp.Text = ocvent.gastemp

        End If

        If paramdist <> "None" Then
            txtFRgastemp.BackColor = Color.LightSalmon
        Else
            txtFRgastemp.BackColor = Color.White
        End If

    End Sub

    Private Sub cboFRcriterion_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFRcriterion.SelectedIndexChanged

        If cboFRcriterion.SelectedItem = "Upper layer gas temperature" Then
            For Each Control In TableLayoutPanel1.Controls
                Control.visible = True
            Next
            Label1.Visible = False
            txtFRintegrity.Visible = False
            cmdDist_integrity.Visible = False
        Else
            For Each Control In TableLayoutPanel1.Controls
                Control.visible = True
            Next
            Label4.Visible = False
            txtFRgastemp.Visible = False
            cmdDist_gastemp.Visible = False
        End If

    End Sub

    Private Sub frmFireResistance_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        cboFRcriterion.Items.Clear()
        cboFRcriterion.Items.Add("Upper layer gas temperature")
        cboFRcriterion.Items.Add("Heat Load using emissive power of fire gases")
        cboFRcriterion.Items.Add("Normalised heat load using Harmathy's furnace eqn")
        cboFRcriterion.Items.Add("Ceiling normalised heat load")
        If Me.Tag = "wall" Then cboFRcriterion.Items.Add("Upper wall normalised heat load")

    End Sub
End Class