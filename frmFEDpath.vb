Option Strict Off
Option Explicit On
Imports System.IO

Friend Class frmFEDpath
    Inherits System.Windows.Forms.Form


    Private Sub frmfedpath_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try

            For i = 1 To NumberRooms
                lstFED_room_1.Items.Add(i.ToString)
                lstFED_room_2.Items.Add(i.ToString)
                lstFED_room_3.Items.Add(i.ToString)

            Next

            txtFED_start_1.Text = FEDPath(0, 0) 'path 1 - start time
            txtFED_start_2.Text = FEDPath(0, 1) 'path 2 - start time
            txtFED_start_3.Text = FEDPath(0, 2) 'path 3 - start time

            txtFED_end_1.Text = FEDPath(1, 0) 'path 1 - end time
            txtFED_end_2.Text = FEDPath(1, 1) 'path 2 - end time
            txtFED_end_3.Text = FEDPath(1, 2) 'path 3 - end time

            lstFED_room_1.SelectedItem = FEDPath(2, 0).ToString 'path 1 - room
            lstFED_room_2.SelectedItem = FEDPath(2, 1).ToString 'path 2  - room
            lstFED_room_3.SelectedItem = FEDPath(2, 2).ToString 'path 3  - room

        Catch ex As Exception

        End Try

    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click

        FEDPath(2, 0) = CInt(lstFED_room_1.SelectedItem) 'path 1 - room
        FEDPath(2, 1) = CInt(lstFED_room_2.SelectedItem) 'path 2 - room
        FEDPath(2, 2) = CInt(lstFED_room_3.SelectedItem) 'path 2 - room

        FEDPath(0, 0) = CSng(txtFED_start_1.Text) 'start times
        FEDPath(0, 1) = CSng(txtFED_start_2.Text)
        FEDPath(0, 2) = CSng(txtFED_start_3.Text)

        FEDPath(1, 0) = CSng(txtFED_end_1.Text) 'end times
        FEDPath(1, 1) = CSng(txtFED_end_2.Text)
        FEDPath(1, 2) = CSng(txtFED_end_3.Text)

        Me.Close()

    End Sub

 

    Private Sub txtFED_start_1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFED_start_1.Validated
        ErrorProvider1.Clear()
        FEDPath(0, 0) = CSng(txtFED_start_1.Text)
    End Sub

    Private Sub txtFED_start_1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtFED_start_1.Validating
        If IsNumeric(txtFED_start_1.Text) Then
            If (CDbl(txtFED_start_1.Text) <= SimTime And CDbl(txtFED_start_1.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtFED_start_1.Select(0, txtFED_start_1.Text.Length)

        txtFED_start_1.Text = txtFED_end_1.Text
        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtFED_start_1, "Invalid Entry. Time must be in the range 0 to " & SimTime.ToString & " sec")
    End Sub

    Private Sub txtFED_start_2_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFED_start_2.Validated
        ErrorProvider1.Clear()
        FEDPath(0, 1) = CSng(txtFED_start_2.Text)
    End Sub

    Private Sub txtFED_start_2_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtFED_start_2.Validating
        If IsNumeric(txtFED_start_2.Text) Then
            If (CDbl(txtFED_start_2.Text) <= SimTime And CDbl(txtFED_start_2.Text) >= CDbl(txtFED_end_1.Text)) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtFED_start_2.Select(0, txtFED_start_2.Text.Length)
        txtFED_start_2.Text = txtFED_end_1.Text
        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtFED_start_2, "Invalid Entry. Time must be in the range " & txtFED_end_1.Text.ToString & " to " & SimTime.ToString & " sec")

    End Sub

    Private Sub txtFED_start_3_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFED_start_3.Validated
        ErrorProvider1.Clear()
        FEDPath(0, 2) = CSng(txtFED_start_3.Text)
    End Sub

    Private Sub txtFED_start_3_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtFED_start_3.Validating
        If IsNumeric(txtFED_start_3.Text) Then
            If (CDbl(txtFED_start_3.Text) <= SimTime And CDbl(txtFED_start_3.Text) >= CDbl(txtFED_end_2.Text)) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtFED_start_3.Select(0, txtFED_start_3.Text.Length)

        txtFED_start_3.Text = txtFED_end_2.Text
        ' Give the ErrorProvider the error message to
        ' display.

        ErrorProvider1.SetError(txtFED_start_3, "Invalid Entry. Time must be in the range " & txtFED_end_2.Text.ToString & " to " & SimTime.ToString & " sec")
    End Sub

   

    Private Sub txtFED_end_1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFED_end_1.Validated
        ErrorProvider1.Clear()
        FEDPath(1, 0) = CSng(txtFED_end_1.Text)
    End Sub

    Private Sub txtFED_end_1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtFED_end_1.Validating
        If IsNumeric(txtFED_end_1.Text) Then
            If (CDbl(txtFED_end_1.Text) <= SimTime And CDbl(txtFED_end_1.Text) >= CDbl(txtFED_start_1.Text)) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtFED_end_1.Select(0, txtFED_end_1.Text.Length)
        txtFED_end_1.Text = SimTime
        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtFED_end_1, "Invalid Entry. Time must be in the range 0 to " & SimTime.ToString & " sec")
    End Sub

    

    Private Sub txtFED_end_2_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFED_end_2.Validated
        ErrorProvider1.Clear()
        FEDPath(1, 1) = CSng(txtFED_end_2.Text)
    End Sub

    Private Sub txtFED_end_2_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtFED_end_2.Validating
        If IsNumeric(txtFED_end_2.Text) Then
            If (CDbl(txtFED_end_2.Text) <= SimTime And CDbl(txtFED_end_2.Text) >= CDbl(txtFED_start_2.Text)) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtFED_end_2.Select(0, txtFED_end_2.Text.Length)
        txtFED_end_2.Text = SimTime

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtFED_end_2, "Invalid Entry. Time must be in the range " & txtFED_start_2.Text.ToString & " to " & SimTime.ToString & " sec")
    End Sub

    Private Sub txtFED_end_3_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFED_end_3.Validated
        ErrorProvider1.Clear()
        FEDPath(1, 2) = CSng(txtFED_end_3.Text)
    End Sub

    Private Sub txtFED_end_3_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtFED_end_3.Validating
        If IsNumeric(txtFED_end_3.Text) Then
            If (CDbl(txtFED_end_3.Text) <= SimTime And CDbl(txtFED_end_3.Text) >= CDbl(txtFED_start_3.Text)) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtFED_end_3.Select(0, txtFED_end_3.Text.Length)
        txtFED_end_3.Text = SimTime

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtFED_end_3, "Invalid Entry. Time must be in the range " & txtFED_start_3.Text.ToString & " to " & SimTime.ToString & " sec")
    End Sub

    Private Sub cmdUpdateFED_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdateFED.Click

        FEDPath(2, 0) = CInt(lstFED_room_1.SelectedItem) 'path 1 - room
        FEDPath(2, 1) = CInt(lstFED_room_2.SelectedItem) 'path 2 - room
        FEDPath(2, 2) = CInt(lstFED_room_3.SelectedItem) 'path 2 - room

        FEDPath(0, 0) = CSng(txtFED_start_1.Text) 'start times
        FEDPath(0, 1) = CSng(txtFED_start_2.Text)
        FEDPath(0, 2) = CSng(txtFED_start_3.Text)

        FEDPath(1, 0) = CSng(txtFED_end_1.Text) 'end times
        FEDPath(1, 1) = CSng(txtFED_end_2.Text)
        FEDPath(1, 2) = CSng(txtFED_end_3.Text)

        Call FED_thermal_iso13571_multi()

        If FEDCO = True Then
            Call FED_CO_iso13571_multi()
        Else
            Call FED_gases_multi()
        End If

    End Sub




End Class