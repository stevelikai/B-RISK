Option Strict Off
Option Explicit On
Friend Class frmQuintiere
	Inherits System.Windows.Forms.Form
	
	Private Sub frmQuintiere_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        Me.Close()
		eventArgs.Cancel = Cancel
	End Sub
	Private Sub frmQuintiere_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'QuitHelp()
		Me.Close()
		
	End Sub

	Private Sub optFTP_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optFTP.CheckedChanged
		If eventSender.Checked Then
			IgnCorrelation = vbFTP
		End If
	End Sub

    Private Sub optJanssens_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optJanssens.CheckedChanged
        If eventSender.Checked Then
            IgnCorrelation = vbJanssens
        End If
    End Sub

	Private Sub optUseAllCurves_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optUseAllCurves.CheckedChanged
		If eventSender.Checked Then
			UseOneCurve = False
		End If
	End Sub
	
    Private Sub optUseOneCurve_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optUseOneCurve.CheckedChanged
        If eventSender.Checked Then
            UseOneCurve = True
        End If
    End Sub
	
	Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        If chkPessimiseWall.Checked = True Then
            PessimiseCombWall = True
        Else
            PessimiseCombWall = False
        End If

        Me.Hide()
	End Sub

    Private Sub frmQuintiere_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Centre_Form(Me)
    End Sub


    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericWallPercent.ValueChanged
        wallpercent = CInt(NumericWallPercent.Value)
    End Sub

    Private Sub NumericUpDown2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericCeilingPercent.ValueChanged
        ceilingpercent = CInt(NumericCeilingPercent.Value)
    End Sub

    Private Sub chkPessimiseWall_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPessimiseWall.CheckedChanged
        If chkPessimiseWall.Checked Then
            lblwallpercent.Visible = True
            NumericWallPercent.Visible = True
            LblHFS.Visible = False
            txtHFSlimit.Visible = False
            LblVFS.Visible = False
            txtVFSlimit.Visible = False
        Else
            lblwallpercent.Visible = False
            NumericWallPercent.Visible = False
            LblHFS.Visible = True
            txtHFSlimit.Visible = True
            LblVFS.Visible = True
            txtVFSlimit.Visible = True
        End If
    End Sub

    Private Sub txtHFSlimit_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHFSlimit.Validated
        ErrorProvider1.Clear()
        wallHFSlimit = CDbl(txtHFSlimit.Text)
    End Sub

    Private Sub txtHFSlimit_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtHFSlimit.Validating
        If IsNumeric(txtHFSlimit.Text) Then
            If (CDbl(txtHFSlimit.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtHFSlimit.Select(0, txtHFSlimit.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtHFSlimit, "Invalid Entry. Horizontal flame spread limit must be greater than 0.")

    End Sub

    Private Sub txtVFSlimit_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVFSlimit.Validated
        ErrorProvider1.Clear()
        wallVFSlimit = CDbl(txtVFSlimit.Text)
    End Sub

    Private Sub txtVFSlimit_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtVFSlimit.Validating
        If IsNumeric(txtVFSlimit.Text) Then
            If (CDbl(txtVFSlimit.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtVFSlimit.Select(0, txtVFSlimit.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtVFSlimit, "Invalid Entry. Vertical flame spread limit must be greater than 0.")

    End Sub

    Private Sub frmQuintiere_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

    End Sub

    Private Sub txtVFSlimit_TextChanged(sender As Object, e As EventArgs) Handles txtVFSlimit.TextChanged

    End Sub
End Class