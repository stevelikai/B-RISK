Imports System.ComponentModel

Public Class frmCLT
    Inherits System.Windows.Forms.Form
    Private Sub frmCLT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'call procedure to centre form
        Centre_Form(Me)

        If DevKey = True Then

        Else
            RB_dynamic.Checked = True
            IntegralModel = False
            KineticModel = False
        End If

        If useCLTmodel = True Then
            optCLTON.Checked = True
        Else
            optCLTOFF.Checked = True
        End If

        RB_dynamic.Checked = True

        If IntegralModel = True Then
            RB_Integral.Checked = True
            RB_dynamic.Checked = False
            RB_Kinetic.Checked = False

            Panel_integralmodel.Visible = True
        Else
            Panel_integralmodel.Visible = False
        End If

        If KineticModel = True Then
            RB_Kinetic.Checked = True
            RB_dynamic.Checked = False
            RB_Integral.Checked = False
            panel_kinetic.Visible = True
        Else
            panel_kinetic.Visible = False
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
        TXT_MoistureContent.Text = init_moisturecontent * 100

        txtA_cell.Text = Format(E_array(1), "0.00E+00")
        txtA_hemi.Text = Format(E_array(2), "0.00E+00")
        txtA_lignin.Text = Format(E_array(3), "0.00E+00")
        txtA_water.Text = Format(E_array(0), "0.00E+00")
        txtE_cell.Text = Format(A_array(1), "0.00E+00")
        txtE_hemi.Text = Format(A_array(2), "0.00E+00")
        txtE_lignin.Text = Format(A_array(3), "0.00E+00")
        txtE_water.Text = Format(A_array(0), "0.00E+00")
        txtReact_cell.Text = n_array(1)
        txtReact_hemi.Text = n_array(2)
        txtReact_lignin.Text = n_array(3)
        txtReact_water.Text = n_array(0)
        txtInit_cell.Text = mf_compinit(1)
        txtInit_hemi.Text = mf_compinit(2)
        txtInit_lignin.Text = mf_compinit(3)
        txtInit_water.Text = mf_compinit(0)
        TextBox_charyield_cell.Text = char_yield(1)
        TextBox_charyield_hemi.Text = char_yield(2)
        TextBox_charyield_lignin.Text = char_yield(3)

        Me.Show()


    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click

        If optCLTON.Checked = True Then
            useCLTmodel = True
        Else
            useCLTmodel = False
        End If

        If RB_Integral.Checked = True Then
            IntegralModel = True
            KineticModel = False
        ElseIf RB_Kinetic.Checked = True Then
            KineticModel = True
            IntegralModel = False
        Else
            KineticModel = False
            IntegralModel = False
        End If

        CLTwallpercent = numeric_wallareapercent.Value
        CLTceilingpercent = numeric_ceilareapercent.Value

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


    Private Sub RB_Integral_Click(sender As Object, e As EventArgs) Handles RB_Integral.Click
        If RB_Integral.Checked = True Then
            Panel_integralmodel.Visible = True
        Else
            Panel_integralmodel.Visible = False
        End If
    End Sub



    Private Sub txtA_hemi_Validating(sender As Object, e As CancelEventArgs) Handles txtA_hemi.Validating
        If IsNumeric(txtA_hemi.Text) Then
            If (CDbl(txtA_hemi.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtA_hemi.Select(0, txtA_hemi.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtA_hemi, "Invalid Entry. ")

    End Sub

    Private Sub txtA_cell_TextChanged(sender As Object, e As EventArgs) Handles txtA_cell.TextChanged

    End Sub

    Private Sub txtA_cell_Validating(sender As Object, e As CancelEventArgs) Handles txtA_cell.Validating
        If IsNumeric(txtA_cell.Text) Then
            If (CDbl(txtA_cell.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtA_cell.Select(0, txtA_cell.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtA_cell, "Invalid Entry. ")
    End Sub

    Private Sub txtA_lignin_Validating(sender As Object, e As CancelEventArgs) Handles txtA_lignin.Validating
        If IsNumeric(txtA_lignin.Text) Then
            If (CDbl(txtA_lignin.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtA_lignin.Select(0, txtA_lignin.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtA_lignin, "Invalid Entry. ")
    End Sub

    Private Sub txtA_water_Validating(sender As Object, e As CancelEventArgs) Handles txtA_water.Validating
        If IsNumeric(txtA_water.Text) Then
            If (CDbl(txtA_water.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtA_water.Select(0, txtA_water.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtA_water, "Invalid Entry. ")
    End Sub

    Private Sub txtE_hemi_TextChanged(sender As Object, e As EventArgs) Handles txtE_hemi.TextChanged

    End Sub

    Private Sub txtE_hemi_Validating(sender As Object, e As CancelEventArgs) Handles txtE_hemi.Validating
        If IsNumeric(txtE_hemi.Text) Then
            If (CDbl(txtE_hemi.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtE_hemi.Select(0, txtE_hemi.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtE_hemi, "Invalid Entry. ")
    End Sub

    Private Sub txtE_cell_Validating(sender As Object, e As CancelEventArgs) Handles txtE_cell.Validating
        If IsNumeric(txtE_cell.Text) Then
            If (CDbl(txtE_cell.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtE_cell.Select(0, txtE_cell.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtE_cell, "Invalid Entry. ")
    End Sub

    Private Sub txtE_lignin_Validating(sender As Object, e As CancelEventArgs) Handles txtE_lignin.Validating

        If IsNumeric(txtE_lignin.Text) Then
            If (CDbl(txtE_lignin.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtE_lignin.Select(0, txtE_lignin.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtE_lignin, "Invalid Entry. ")
    End Sub

    Private Sub txtE_water_Validating(sender As Object, e As CancelEventArgs) Handles txtE_water.Validating

        If IsNumeric(txtE_water.Text) Then
            If (CDbl(txtE_water.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtE_water.Select(0, txtE_water.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtE_water, "Invalid Entry. ")
    End Sub

    Private Sub txtReact_hemi_Validating(sender As Object, e As CancelEventArgs) Handles txtReact_hemi.Validating

        If IsNumeric(txtReact_hemi.Text) Then
            If (CDbl(txtReact_hemi.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtReact_hemi.Select(0, txtReact_hemi.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtReact_hemi, "Invalid Entry. ")
    End Sub

    Private Sub txtReact_cell_Validating(sender As Object, e As CancelEventArgs) Handles txtReact_cell.Validating
        If IsNumeric(txtReact_cell.Text) Then
            If (CDbl(txtReact_cell.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtReact_cell.Select(0, txtReact_cell.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtReact_cell, "Invalid Entry. ")
    End Sub

    Private Sub txtReact_lignin_Validating(sender As Object, e As CancelEventArgs) Handles txtReact_lignin.Validating
        If IsNumeric(txtReact_lignin.Text) Then
            If (CDbl(txtReact_lignin.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtReact_lignin.Select(0, txtReact_lignin.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtReact_lignin, "Invalid Entry. ")

    End Sub

    Private Sub txtReact_water_Validating(sender As Object, e As CancelEventArgs) Handles txtReact_water.Validating
        If IsNumeric(txtReact_water.Text) Then
            If (CDbl(txtReact_water.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtReact_water.Select(0, txtReact_water.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtReact_water, "Invalid Entry. ")
    End Sub

    Private Sub txtInit_water_Validating(sender As Object, e As CancelEventArgs) Handles txtInit_water.Validating
        If IsNumeric(txtInit_water.Text) Then
            If (CDbl(txtInit_water.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtInit_water.Select(0, txtInit_water.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtInit_water, "Invalid Entry. ")
    End Sub

    Private Sub txtInit_lignin_Validating(sender As Object, e As CancelEventArgs) Handles txtInit_lignin.Validating
        If IsNumeric(txtInit_lignin.Text) Then
            If (CDbl(txtInit_lignin.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtInit_lignin.Select(0, txtInit_lignin.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtInit_lignin, "Invalid Entry. ")
    End Sub

    Private Sub txtInit_cell_Validating(sender As Object, e As CancelEventArgs) Handles txtInit_cell.Validating
        If IsNumeric(txtInit_cell.Text) Then
            If (CDbl(txtInit_cell.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtInit_cell.Select(0, txtInit_cell.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtInit_cell, "Invalid Entry. ")
    End Sub

    Private Sub txtInit_hemi_Validating(sender As Object, e As CancelEventArgs) Handles txtInit_hemi.Validating
        If IsNumeric(txtInit_hemi.Text) Then
            If (CDbl(txtInit_hemi.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtInit_hemi.Select(0, txtInit_hemi.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtInit_hemi, "Invalid Entry. ")
    End Sub

    Private Sub txtA_hemi_Validated(sender As Object, e As EventArgs) Handles txtA_hemi.Validated
        ErrorProvider1.Clear()
        E_array(2) = CDbl(txtA_hemi.Text)
    End Sub

    Private Sub txtA_water_Validated(sender As Object, e As EventArgs) Handles txtA_water.Validated
        ErrorProvider1.Clear()
        E_array(0) = CDbl(txtA_water.Text)
    End Sub

    Private Sub txtA_cell_Validated(sender As Object, e As EventArgs) Handles txtA_cell.Validated
        ErrorProvider1.Clear()
        E_array(1) = CDbl(txtA_cell.Text)
    End Sub

    Private Sub txtA_lignin_Validated(sender As Object, e As EventArgs) Handles txtA_lignin.Validated
        ErrorProvider1.Clear()
        E_array(3) = CDbl(txtA_lignin.Text)
    End Sub

    Private Sub txtE_hemi_Validated(sender As Object, e As EventArgs) Handles txtE_hemi.Validated
        ErrorProvider1.Clear()
        A_array(2) = CDbl(txtE_hemi.Text)
    End Sub

    Private Sub txtE_cell_Validated(sender As Object, e As EventArgs) Handles txtE_cell.Validated
        ErrorProvider1.Clear()
        A_array(1) = CDbl(txtE_hemi.Text)
    End Sub

    Private Sub txtE_lignin_Validated(sender As Object, e As EventArgs) Handles txtE_lignin.Validated
        ErrorProvider1.Clear()
        A_array(3) = CDbl(txtE_lignin.Text)
    End Sub

    Private Sub txtE_water_Validated(sender As Object, e As EventArgs) Handles txtE_water.Validated
        ErrorProvider1.Clear()
        A_array(0) = CDbl(txtE_water.Text)
    End Sub

    Private Sub txtReact_hemi_Validated(sender As Object, e As EventArgs) Handles txtReact_hemi.Validated
        ErrorProvider1.Clear()
        n_array(2) = CDbl(txtReact_hemi.Text)
    End Sub

    Private Sub txtReact_water_Validated(sender As Object, e As EventArgs) Handles txtReact_water.Validated
        ErrorProvider1.Clear()
        n_array(0) = CDbl(txtReact_water.Text)
    End Sub

    Private Sub txtReact_lignin_Validated(sender As Object, e As EventArgs) Handles txtReact_lignin.Validated
        ErrorProvider1.Clear()
        n_array(3) = CDbl(txtReact_lignin.Text)
    End Sub

    Private Sub txtReact_cell_Validated(sender As Object, e As EventArgs) Handles txtReact_cell.Validated
        ErrorProvider1.Clear()
        n_array(1) = CDbl(txtReact_cell.Text)
    End Sub

    Private Sub RB_Kinetic_CheckedChanged(sender As Object, e As EventArgs) Handles RB_Kinetic.CheckedChanged
        If RB_Kinetic.Checked = True Then
            Panel_integralmodel.Visible = False
            panel_kinetic.Visible = True
        End If
    End Sub

    Private Sub RB_Integral_CheckedChanged(sender As Object, e As EventArgs) Handles RB_Integral.CheckedChanged
        If RB_Integral.Checked = True Then
            Panel_integralmodel.Visible = True
            panel_kinetic.Visible = False
        End If
    End Sub

    Private Sub RB_dynamic_CheckedChanged(sender As Object, e As EventArgs) Handles RB_dynamic.CheckedChanged
        If RB_dynamic.Checked = True Then
            Panel_integralmodel.Visible = False
            panel_kinetic.Visible = False
        End If
    End Sub

    Private Sub TextBox_charyield_hemi_Validating(sender As Object, e As CancelEventArgs) Handles TextBox_charyield_hemi.Validating
        If IsNumeric(TextBox_charyield_hemi.Text) Then
            If (CDbl(TextBox_charyield_hemi.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        TextBox_charyield_hemi.Select(0, TextBox_charyield_hemi.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(TextBox_charyield_hemi, "Invalid Entry. ")
    End Sub

    Private Sub TextBox_charyield_cell_Validating(sender As Object, e As CancelEventArgs) Handles TextBox_charyield_cell.Validating
        If IsNumeric(TextBox_charyield_cell.Text) Then
            If (CDbl(TextBox_charyield_cell.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        TextBox_charyield_cell.Select(0, TextBox_charyield_cell.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(TextBox_charyield_cell, "Invalid Entry. ")
    End Sub

    Private Sub TextBox_charyield_lignin_Validating(sender As Object, e As CancelEventArgs) Handles TextBox_charyield_lignin.Validating
        If IsNumeric(TextBox_charyield_lignin.Text) Then
            If (CDbl(TextBox_charyield_lignin.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        TextBox_charyield_lignin.Select(0, TextBox_charyield_lignin.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(TextBox_charyield_lignin, "Invalid Entry. ")
    End Sub

    Private Sub TextBox_charyield_lignin_Validated(sender As Object, e As EventArgs) Handles TextBox_charyield_lignin.Validated
        ErrorProvider1.Clear()
        char_yield(3) = CDbl(TextBox_charyield_lignin.Text)
    End Sub

    Private Sub TextBox_charyield_cell_Validated(sender As Object, e As EventArgs) Handles TextBox_charyield_cell.Validated
        ErrorProvider1.Clear()
        char_yield(1) = CDbl(TextBox_charyield_cell.Text)
    End Sub

    Private Sub TextBox_charyield_hemi_Validated(sender As Object, e As EventArgs) Handles TextBox_charyield_hemi.Validated
        ErrorProvider1.Clear()
        char_yield(2) = CDbl(TextBox_charyield_hemi.Text)
    End Sub

    Private Sub TXT_MoistureContent_TextChanged(sender As Object, e As EventArgs) Handles TXT_MoistureContent.TextChanged

    End Sub

    Private Sub TXT_MoistureContent_Validating(sender As Object, e As CancelEventArgs) Handles TXT_MoistureContent.Validating
        If IsNumeric(TXT_MoistureContent.Text) Then
            If (CDbl(TXT_MoistureContent.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtA_hemi.Select(0, TXT_MoistureContent.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(TXT_MoistureContent, "Invalid Entry. ")
    End Sub

    Private Sub TXT_MoistureContent_Validated(sender As Object, e As EventArgs) Handles TXT_MoistureContent.Validated
        ErrorProvider1.Clear()
        init_moisturecontent = CDbl(TXT_MoistureContent.Text) / 100
    End Sub
End Class