Imports System.Collections.Generic
Public Class frmNewCVents
    Public ventid As Integer
    Public ocVent As oCVent
    Public ocVents As List(Of oCVent)
    Public ocventdistributions As List(Of oDistribution)

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Try

            ocVents = VentDB.GetCVents
            ocventdistributions = VentDB.GetCVentDistributions

            For Each Me.ocVent In ocVents

                If Me.ocVent.id = Me.Tag Then

                    ocVent.description = txtVentDescription.Text

                    If IsNumeric(lstCVentIDUpper.SelectedItem) Then
                        ocVent.upperroom = lstCVentIDUpper.SelectedItem
                    Else
                        ocVent.upperroom = NumberRooms + 1
                    End If
                    If IsNumeric(lstCVentIDlower.SelectedItem) Then
                        ocVent.lowerroom = lstCVentIDlower.SelectedItem
                    Else
                        ocVent.lowerroom = NumberRooms + 1
                    End If

                    ocVent.area = txtCVentArea.Text
                    ocVent.opentime = txtCVentOpenTime.Text
                    ocVent.closetime = txtCVentCloseTime.Text
                    ocVent.dischargecoeff = txtDischargeCoeff.Text

                    For Each x In ocventdistributions
                        If x.id = ocVent.id Then
                            If x.varname = "area" Then
                                x.varvalue = txtCVentArea.Text
                            End If
                           
                        End If

                    Next

                End If
            Next

            If ocVent IsNot Nothing Then
                VentDB.SaveCVents(ocVents, ocventdistributions)
                frmCVentList.FillVentList()

                Call frmInputs.ceilingventarrays(ocVents, ocventdistributions, 0)
            End If

            Me.Close()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in frmNewCVents.vb cmdSave_Click")

        End Try
    End Sub

    Private Sub frmNewCVents_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        AddOwnedForm(frmVentOpenOptions)
        AddOwnedForm(frmFireResistance)
    End Sub

    Private Sub txtCVentArea_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCVentArea.Validated
        ErrorProvider1.Clear()
        ocVent.area = CDbl(txtCVentArea.Text)
    End Sub

    Private Sub txtCVentArea_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCVentArea.Validating
        If IsNothing(ocVents) Then
            ocVents = VentDB.GetCVents
        End If
        If IsNumeric(txtCVentArea.Text) Then
            If (CDbl(txtCVentArea.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtCVentArea.Select(0, txtCVentArea.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtCVentArea, "Invalid Entry. Vent area must be greater than 0 m.")

    End Sub

    Private Sub txtCVentOpenTime_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCVentOpenTime.Validated
        ErrorProvider1.Clear()
        ocVent.opentime = CDbl(txtCVentOpenTime.Text)
    End Sub

    Private Sub txtCVentOpenTime_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCVentOpenTime.Validating
        If IsNothing(ocVents) Then
            ocVents = VentDB.GetCVents
        End If
        If IsNumeric(txtCVentOpenTime.Text) Then
            If (CDbl(txtCVentOpenTime.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtCVentOpenTime.Select(0, txtCVentOpenTime.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtCVentOpenTime, "Invalid Entry. Opening time must be greater than or equal to 0 s.")
    End Sub


    Private Sub txtCVentCloseTime_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCVentCloseTime.Validated
        ErrorProvider1.Clear()
        ocVent.closetime = CDbl(txtCVentCloseTime.Text)
    End Sub

    Private Sub txtCVentCloseTime_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCVentCloseTime.Validating
        If IsNothing(ocVents) Then
            ocVents = VentDB.GetCVents
        End If
        If IsNumeric(txtCVentCloseTime.Text) Then
            If (CDbl(txtCVentCloseTime.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtCVentCloseTime.Select(0, txtCVentCloseTime.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtCVentCloseTime, "Invalid Entry. Closing time must be greater than or equal to 0 s.")

    End Sub

    Private Sub cmdDist_cvarea_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_cvarea.Click
        ocVents = VentDB.GetCVents
        ocventdistributions = VentDB.GetCVentDistributions()

        ventid = Me.Tag - 1
        Dim ocvent As oCVent = ocVents(ventid)
        Dim paramdist As String
        Dim param As String = "area"
        Dim units As String = "m2"
        Dim instruction As String = "vent area"

        Call frmDistParam.ShowCVentDistributionForm(param, units, instruction, ocvent, ocventdistributions, ventid, paramdist)

        If paramdist <> "None" Then
            txtCVentArea.BackColor = Color.LightSalmon
        Else
            txtCVentArea.BackColor = Color.White
        End If

        txtCVentArea.Text = ocvent.area

    End Sub

    
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFireResistance.Click

        ocVents = VentDB.GetCVents
        ocventdistributions = VentDB.GetCVentDistributions
        ocVent = ocVents(Me.Tag - 1)
        frmFireResistance.FRoptions2(ocVent, ocVents, ocventdistributions)
        If ocVent.FRcriteria = -1 Then
            frmFireResistance.cboFRcriterion.SelectedIndex = 0
        Else
            frmFireResistance.cboFRcriterion.SelectedIndex = ocVent.FRcriteria
        End If
    End Sub


    Private Sub cmdVentOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdVentOptions.Click
        Try

            ocVents = VentDB.GetCVents
            ocventdistributions = VentDB.GetCVentDistributions
            ocVent = ocVents(Me.Tag - 1)
            frmVentOpenOptions.Text = "Ceiling Vent Opening Options"
            Call frmVentOpenOptions.ventopenoptions2(ocVent, ocVents, ocventdistributions)

            If frmVentOpenOptions.chkAutoVentOpen.Checked = True Then
                frmVentOpenOptions.GroupBox1.Visible = True
                frmVentOpenOptions.GroupBox2.Visible = True
                frmVentOpenOptions.GroupBox3.Visible = True
                frmVentOpenOptions.GroupBox4.Visible = True
                frmVentOpenOptions.GroupBox5.Visible = True
                frmVentOpenOptions.GroupBox6.Visible = True
                frmVentOpenOptions.GroupBox7.Visible = True
                frmVentOpenOptions.GroupBox8.Visible = True
                frmVentOpenOptions.GroupBox9.Visible = False
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
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in frmNewCVent.vb cmdVentOptions_click - line " & Err.Erl)

        End Try
    End Sub

    Private Sub txtDischargeCoeff_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDischargeCoeff.Validated
        ErrorProvider1.Clear()
        ocVent.dischargecoeff = CDbl(txtDischargeCoeff.Text)
    End Sub

    Private Sub txtDischargeCoeff_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDischargeCoeff.Validating
        If IsNothing(ocVents) Then
            ocVents = VentDB.GetCVents
        End If
        If IsNumeric(txtDischargeCoeff.Text) Then
            If (CDbl(txtDischargeCoeff.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtDischargeCoeff.Select(0, txtDischargeCoeff.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtDischargeCoeff, "Invalid Entry. Discharge coefficient must be greater than or equal to 0.")

    End Sub
End Class