Public Class frmPowerlaw

    Private Sub TxtAlphaT_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAlphaT.Validated

        ErrorProvider1.Clear()

        AlphaT = CSng(TxtAlphaT.Text)

        frmOptions1.Update_Distributions("Alpha T")

    End Sub

    Private Sub TxtAlphaT_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TxtAlphaT.Validating
        If IsNumeric(TxtAlphaT.Text) Then
            If (CSng(TxtAlphaT.Text) <= 10 And CSng(TxtAlphaT.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        TxtAlphaT.Select(0, TxtAlphaT.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(TxtAlphaT, "Invalid Entry. Alpha must be in the range 0 to 10. ")

    End Sub

    Private Sub txtPeakHRR_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPeakHRR.Validated
        ErrorProvider1.Clear()

        PeakHRR = CSng(txtPeakHRR.Text)

        frmOptions1.Update_Distributions("Peak HRR")

    End Sub

    Private Sub txtPeakHRR_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtPeakHRR.Validating
        If IsNumeric(txtPeakHRR.Text) Then
            If (CSng(txtPeakHRR.Text) <= 500001 And CSng(txtPeakHRR.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtPeakHRR.Select(0, txtPeakHRR.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtPeakHRR, "Invalid Entry. Peak HRR must be in the range 0 to 500,000 kW. ")

    End Sub

    Private Sub cmdDist_alphaT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_alphaT.Click
        Dim param As String = "Alpha T"
        Dim units As String = "kW/s2"
        If useT2fire = False Then
            units = "kW/s3/m"
        End If
        Dim instruction As String = "alpha constant for power law fire."

        Call frmDistParam.ShowDistributionForm(param, units, instruction)
    End Sub

    Private Sub cmdPeakHRR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPeakHRR.Click
        Dim param As String = "Peak HRR"
        Dim units As String = "kW"
        Dim instruction As String = "Peak HRR for power law fire."

        Call frmDistParam.ShowDistributionForm(param, units, instruction)
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click

        If optAlpha2.Checked = True Then
            useT2fire = True
            txtStoreHeight.Enabled = False
            Label1.Text = "alpha (kW/s2)"
        Else
            useT2fire = False
            txtStoreHeight.Enabled = True
            Label1.Text = "alpha (kW/s3/m)"
        End If

        Me.Close()
    End Sub

    Private Sub txtStoreHeight_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtStoreHeight.Validated
        ErrorProvider1.Clear()

        StoreHeight = CSng(txtStoreHeight.Text)

        frmOptions1.Update_Distributions("Storage Height")

    End Sub

    Private Sub txtStoreHeight_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtStoreHeight.Validating
        If IsNumeric(txtStoreHeight.Text) Then
            If (CSng(txtStoreHeight.Text) <= 50 And CSng(txtStoreHeight.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtStoreHeight.Select(0, txtStoreHeight.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtStoreHeight, "Invalid Entry. Storage height must be in the range 0 to 50 m. ")

    End Sub

    Private Sub optAlpha2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optAlpha2.CheckedChanged
        If optAlpha2.Checked = True Then
            '    useT2fire = True
            txtStoreHeight.Enabled = False
            '    Label1.Text = "alpha (kW/s2)"
        Else
            '    useT2fire = False
            txtStoreHeight.Enabled = True
            '    Label1.Text = "alpha (kW/s3/m)"
        End If
    End Sub


    Private Sub optAlpha3_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optAlpha3.CheckedChanged
        If optAlpha2.Checked = True Then
            '    useT2fire = True
            txtStoreHeight.Enabled = False
            '    Label1.Text = "alpha (kW/s2)"
        Else
            '    useT2fire = False
            txtStoreHeight.Enabled = True
            '    Label1.Text = "alpha (kW/s3/m)"

        End If
    End Sub


End Class