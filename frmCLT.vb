Imports System.ComponentModel

Public Class frmCLT
    Inherits System.Windows.Forms.Form
    Private Sub frmCLT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'call procedure to centre form
        Centre_Form(Me)
        If useCLTmodel = True Then
            optCLTON.Checked = True
        Else
            optCLTOFF.Checked = True
        End If

        numeric_ceilareapercent.Value = CLTceilingpercent
        numeric_wallareapercent.Value = CLTwallpercent
        txtCharTemp.Text = chartemp
        txtCLTcalibration.Text = CLTcalibrationfactor
        txtLamellaDepth.Text = Lamella
        txtFlameFlux.Text = CLTflameflux
        txtCLTigtemp.Text = CLTigtemp
        txtCLTLoG.Text = CLTLoG
        txtCritFlux.Text = CLTQcrit
        txtDebondTemp.Text = DebondTemp

        'CLTflameflux = 17
        'CLTigtemp = 384
        'CLTLoG = 1.6
        'CLTQcrit = 16
        'CLTcalibrationfactor = 1

    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click

        If optCLTON.Checked = True Then
            useCLTmodel = True
        Else
            useCLTmodel = False
        End If

        If chkWoodIntegralModel.Checked = True Then
            IntegralModel = True
        Else
            IntegralModel = False
        End If

        If chkKineticModel.Checked = True Then
            KineticModel = True
        Else
            KineticModel = False
        End If

        Me.Close()

    End Sub

    Private Sub txtCharTemp_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCharTemp.Validated
        ErrorProvider1.Clear()

        chartemp = CDbl(txtCharTemp.Text)

    End Sub

    Private Sub txtCharTemp_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCharTemp.Validating
        If IsNumeric(txtCharTemp.Text) Then
            If (CDbl(txtCharTemp.Text) > 30) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtCharTemp.Select(0, txtCharTemp.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtCharTemp, "Invalid Entry. Char temperature must be greater than 30 C. ")

    End Sub

    Private Sub numeric_wallareapercent_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numeric_wallareapercent.ValueChanged
        CLTwallpercent = CInt(numeric_wallareapercent.Value)
    End Sub

    Private Sub numeric_ceilareapercent_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numeric_ceilareapercent.ValueChanged
        CLTceilingpercent = CInt(numeric_ceilareapercent.Value)
    End Sub



    Private Sub txtLamellaDepth_Validated(sender As Object, e As EventArgs) Handles txtLamellaDepth.Validated
        ErrorProvider1.Clear()

        Lamella = CSng(txtLamellaDepth.Text)
    End Sub

    Private Sub txtLamellaDepth_Validating(sender As Object, e As CancelEventArgs) Handles txtLamellaDepth.Validating
        If IsNumeric(txtLamellaDepth.Text) Then
            If (CDbl(txtLamellaDepth.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtLamellaDepth.Select(0, txtLamellaDepth.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtLamellaDepth, "Invalid Entry. Lamella depth must be greater than 0. ")
    End Sub

    Private Sub txtCritFlux_Validated(sender As Object, e As EventArgs) Handles txtCritFlux.Validated
        ErrorProvider1.Clear()

        CLTQcrit = CDbl(txtCritFlux.Text)
    End Sub

    Private Sub txtCritFlux_Validating(sender As Object, e As CancelEventArgs) Handles txtCritFlux.Validating
        If IsNumeric(txtCritFlux.Text) Then
            If (CDbl(txtCritFlux.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtCritFlux.Select(0, txtCritFlux.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtCritFlux, "Invalid Entry. Critical flux must be greater than 0. ")
    End Sub

    Private Sub txtFlameFlux_Validating(sender As Object, e As CancelEventArgs) Handles txtFlameFlux.Validating
        If IsNumeric(txtFlameFlux.Text) Then
            If (CDbl(txtFlameFlux.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtFlameFlux.Select(0, txtFlameFlux.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtFlameFlux, "Invalid Entry. Flame flux must be greater than 0. ")
    End Sub

    Private Sub txtFlameFlux_Validated(sender As Object, e As EventArgs) Handles txtFlameFlux.Validated
        ErrorProvider1.Clear()

        CLTflameflux = CDbl(txtFlameFlux.Text)
    End Sub

    Private Sub txtCLTLoG_Validating(sender As Object, e As CancelEventArgs) Handles txtCLTLoG.Validating
        If IsNumeric(txtCLTLoG.Text) Then
            If (CDbl(txtCLTLoG.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtCLTLoG.Select(0, txtCLTLoG.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtCLTLoG, "Invalid Entry. Latent heat of gasification must be greater than 0. ")
    End Sub

    Private Sub txtCLTLoG_Validated(sender As Object, e As EventArgs) Handles txtCLTLoG.Validated
        ErrorProvider1.Clear()

        CLTLoG = CDbl(txtCLTLoG.Text)
    End Sub

    Private Sub txtCLTigtemp_Validating(sender As Object, e As CancelEventArgs) Handles txtCLTigtemp.Validating
        If IsNumeric(txtCLTigtemp.Text) Then
            If (CDbl(txtCLTigtemp.Text) > 200) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtCLTigtemp.Select(0, txtCLTigtemp.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtCLTigtemp, "Invalid Entry. Ignition temperature must be greater than 200. ")
    End Sub

    Private Sub txtCLTigtemp_Validated(sender As Object, e As EventArgs) Handles txtCLTigtemp.Validated
        ErrorProvider1.Clear()

        CLTigtemp = CDbl(txtCLTigtemp.Text)
    End Sub

    Private Sub txtCLTcalibration_Validated(sender As Object, e As EventArgs) Handles txtCLTcalibration.Validated
        ErrorProvider1.Clear()

        CLTcalibrationfactor = CDbl(txtCLTcalibration.Text)
    End Sub

    Private Sub txtCLTcalibration_Validating(sender As Object, e As CancelEventArgs) Handles txtCLTcalibration.Validating
        If IsNumeric(txtCLTcalibration.Text) Then
            If (CDbl(txtCLTcalibration.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtCLTcalibration.Select(0, txtCLTcalibration.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtCLTcalibration, "Invalid Entry. Calibration factor must be  must be greater than 0. ")
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            System.Diagnostics.Process.Start(URL2)

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in frmCLT.vb LinkLabel1_LinkClicked")
        End Try
    End Sub


    Private Sub txtDebondTemp_Validating(sender As Object, e As CancelEventArgs) Handles txtDebondTemp.Validating
        If IsNumeric(txtDebondTemp.Text) Then
            If (CDbl(txtDebondTemp.Text) > 50) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtDebondTemp.Select(0, txtDebondTemp.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtDebondTemp, "Invalid Entry. Adhesive debonding temperature must be must be greater than 50 C. ")
    End Sub

    Private Sub txtDebondTemp_Validated(sender As Object, e As EventArgs) Handles txtDebondTemp.Validated
        ErrorProvider1.Clear()

        DebondTemp = CDbl(txtDebondTemp.Text)
    End Sub

End Class