Imports System
Imports System.Collections.Generic

Public Class frmFanList
    Inherits System.Windows.Forms.Form
    Public NewFanForm As frmNewFan
    Dim oFans As List(Of oFan)
    Dim oFandistributions As List(Of oDistribution)

    Private Sub txtFanReliability_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFanReliability.TextChanged
        'store fan reliability
        If IsNumeric(txtFanReliability.Text) = True Then
            FanReliability = CDec(txtFanReliability.Text)
        End If
    End Sub

    Private Sub txtFanReliability_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFanReliability.Validated
        ErrorProvider1.Clear()
        FanReliability = CSng(Me.txtFanReliability.Text)
    End Sub

    Private Sub txtFanReliability_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtFanReliability.Validating
        If IsNumeric(txtFanReliability.Text) Then
            If (CSng(txtFanReliability.Text) >= 0) And (CSng(txtFanReliability.Text) <= 1) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtFanReliability.Select(0, txtFanReliability.Text.Length)

        ' Give the ErrorProvider the error message to display.
        ErrorProvider1.SetError(txtFanReliability, "Invalid Entry. Value must be between 0 and 1.")

    End Sub

    Private Sub cmdAddFan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddFan.Click
        'add fan

        ofans = FanDB.GetFans
        Dim ofan As oFan
        ofandistributions = FanDB.GetFanDistributions

        NewFanForm = New frmNewFan
        NewFanForm.Text = "New Fan"
        ofan = New oFan(1, 1, 0, 0, RoomHeight(1), 50, 0, True, True, True)

        ofan.fanid = oFans.Count + 1
        NewFanForm.Tag = oFans.Count + 1

        If ofan IsNot Nothing Then
            oFans.Add(ofan)

            NewFanForm.txtFanElevation.Text = ofan.fanelevation
            NewFanForm.txtFanReliability.Text = ofan.fanreliability
            NewFanForm.NumericUpDownRoom.Value = ofan.fanroom
            NewFanForm.NumericUpDownRoom.Maximum = NumberRooms
            NewFanForm.NumericUpDownRoom.Minimum = 1
            If ofan.fanextract = True Then
                NewFanForm.optExtract.Checked = True
            Else
                NewFanForm.optPressurise.Checked = True
            End If
            If ofan.fanmanual = 0 Then
                NewFanForm.optFanManual.Checked = True
            ElseIf ofan.fanmanual = 1 Then
                NewFanForm.optFanAuto.Checked = True
            ElseIf ofan.fanmanual = 2 Then
                NewFanForm.optFanAutoFR.Checked = True
            End If
            NewFanForm.chkUseFanCurve.Checked = ofan.fancurve

            Dim oDistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            oDistribution.id = ofan.fanid
            oDistribution.varname = "fanflowrate"
            oDistribution.varvalue = 0
            NewFanForm.txtFanFlowRate.Text = oDistribution.varvalue
            oFandistributions.Add(oDistribution)

            oDistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            oDistribution.id = ofan.fanid
            oDistribution.varname = "fanstarttime"
            oDistribution.varvalue = 0
            NewFanForm.txtFanStartTime.Text = oDistribution.varvalue
            oFandistributions.Add(oDistribution)

            oDistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            oDistribution.id = ofan.fanid
            oDistribution.varname = "fanpressurelimit"
            oDistribution.varvalue = 50
            NewFanForm.txtPressureLimit.Text = oDistribution.varvalue
            oFandistributions.Add(oDistribution)

            oDistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            oDistribution.id = ofan.fanid
            oDistribution.varname = "fanreliability"
            oDistribution.varvalue = 50
            NewFanForm.txtPressureLimit.Text = oDistribution.varvalue
            oFandistributions.Add(oDistribution)

            FanDB.SaveFans(oFans, oFandistributions)
        End If

        Me.FillFanList(oFans)
        lstFan.SelectedIndex = ofan.fanid - 1
        NewFanForm.ShowDialog()
    End Sub
    Public Sub FillFanList(ByVal oFans)
        lstFan.Items.Clear()
        Dim j As Integer = 0

        For Each p As oFan In oFans
            j = j + 1
            lstFan.Items.Add(j.ToString & vbTab & p.GetDisplayText(vbTab))
        Next

    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Dim oDistributions As New List(Of oDistribution)
        oDistributions = DistributionClass.GetDistributions

        For Each oDistribution In oDistributions
            If oDistribution.varname = "Mechanical Ventilation Reliability" Then
                oDistribution.varvalue = txtFanReliability.Text
            End If
        Next
        DistributionClass.SaveDistributions(oDistributions)

        Me.Close()
    End Sub

    Private Sub frmFanList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
           Dim sep As String = vbTab
        Dim text As String
        oFans = FanDB.GetFans
        Me.FillFanList(oFans)

        text = "ID" & sep & "Room" & sep & "flow rate (m3/s)" & sep & "elevation (m)"
        ListBox1.Items.Add(text)

        Me.CenterToParent()
    End Sub

    Private Sub cmdEditFan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEditFan.Click
        Dim i As Integer = lstFan.SelectedIndex
        If i <> -1 Then
            oFans = FanDB.GetFans()
            Dim ofan As oFan = oFans(i)
            oFandistributions = FanDB.GetFanDistributions

            Dim NewfanForm As New frmNewFan

            NewfanForm.Text = "Edit Fan"
            NewfanForm.oFan = ofan

            If NewfanForm.oFan IsNot Nothing Then
                NewfanForm.Tag = ofan.fanid
                NewfanForm.txtFanElevation.Text = ofan.fanelevation
                NewfanForm.txtFanReliability.Text = ofan.fanreliability
                NewfanForm.NumericUpDownRoom.Value = ofan.fanroom
                NewfanForm.NumericUpDownRoom.Maximum = NumberRooms
                NewfanForm.NumericUpDownRoom.Minimum = 1
                If ofan.fanextract = True Then
                    NewfanForm.optExtract.Checked = True
                Else
                    NewfanForm.optPressurise.Checked = True
                End If
                If ofan.fanmanual = 0 Then
                    NewfanForm.optFanManual.Checked = True
                ElseIf ofan.fanmanual = 1 Then
                    NewfanForm.optFanAuto.Checked = True
                ElseIf ofan.fanmanual = 2 Then
                    NewfanForm.optFanAutoFR.Checked = True
                End If
                NewfanForm.chkUseFanCurve.Checked = ofan.fancurve

                For Each x In oFandistributions
                    If x.id = ofan.fanid Then
                        If x.varname = "fanflowrate" Then
                            NewfanForm.txtFanFlowRate.Text = x.varvalue
                            If x.distribution <> "None" Then
                                NewfanForm.txtFanFlowRate.BackColor = distbackcolor
                            Else
                                NewfanForm.txtFanFlowRate.BackColor = distnobackcolor
                            End If
                        End If
                        If x.varname = "fanstarttime" Then
                            NewfanForm.txtFanStartTime.Text = x.varvalue
                            If x.distribution <> "None" Then
                                NewfanForm.txtFanStartTime.BackColor = distbackcolor
                            Else
                                NewfanForm.txtFanStartTime.BackColor = distnobackcolor
                            End If
                        End If
                        If x.varname = "fanpressurelimit" Then
                            NewfanForm.txtPressureLimit.Text = x.varvalue
                            If x.distribution <> "None" Then
                                NewfanForm.txtPressureLimit.BackColor = distbackcolor
                            Else
                                NewfanForm.txtPressureLimit.BackColor = distnobackcolor
                            End If
                        End If
                        If x.varname = "fanreliability" Then
                            NewfanForm.txtFanReliability.Text = x.varvalue
                            If x.distribution <> "None" Then
                                NewfanForm.txtFanReliability.BackColor = distbackcolor
                            Else
                                NewfanForm.txtFanReliability.BackColor = distnobackcolor
                            End If
                        End If
                    End If
                Next

                NewfanForm.ShowDialog()

            End If

        End If
    End Sub

    Private Sub cmdCopyFan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopyFan.Click
        Dim oFans As List(Of oFan)
        Dim ofandistributions As List(Of oDistribution)
        Dim i As Integer = lstFan.SelectedIndex
        Dim id As Integer
        Dim j As Integer = 0
        If i <> -1 Then
            oFans = FanDB.GetFans
            Dim oFan As oFan = oFans(i)
            id = oFan.fanid
            ofandistributions = FanDB.GetFanDistributions
            Dim odistribution As New oDistribution
            Dim NewFanForm As New frmNewFan

            NewFanForm.Text = "Copy Fan"
            NewFanForm.oFan = oFan
            Dim counter As Integer
            If NewFanForm.oFan IsNot Nothing Then

                counter = 0

                For z = 0 To ofandistributions.Count - 1
                    If ofandistributions(z).id = id Then

                        Dim x As oDistribution = ofandistributions(z)
                        Dim y As New oDistribution(x.varname, x.units, x.distribution, x.varvalue, x.mean, x.variance, x.lbound, x.ubound, x.mode, x.alpha, x.beta)
                        y.id = oFans.Count + 1
                        ofandistributions.Add(y)

                    End If
                Next
                NewFanForm.oFan.fanid = oFans.Count + 1
                oFans.Add(NewFanForm.oFan)

                FanDB.SaveFans(oFans, ofandistributions)
                Me.FillFanList(oFans)
            End If

        End If
        Exit Sub
    End Sub

    Private Sub cmdRemoveFan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveFan.Click
        Dim i As Integer = lstFan.SelectedIndex
        oFans = FanDB.GetFans()

        Dim j As Integer = 0
        If i <> -1 Then
            Dim oFan As oFan = oFans(i)
            Dim ofandistributions As List(Of oDistribution)

            Dim message As String = "Are you sure you want to delete " _
                & "Fan " & CStr(i + 1) & "?"
            Dim button As DialogResult = MessageBox.Show(message, _
                "Confirm Delete", MessageBoxButtons.YesNo)
            If button = Windows.Forms.DialogResult.Yes Then
                ofandistributions = FanDB.GetFanDistributions()

here:
                For Each oDistribution In ofandistributions
                    If oDistribution.id = oFan.fanid Then
                        ofandistributions.Remove(oDistribution)
                        GoTo here
                    End If
                Next

                oFans.Remove(oFan)

                'sort and reindex
                Dim count As Integer = 1
                For Each oFan In oFans

                    For Each oDistribution In ofandistributions
                        If oDistribution.id = oFan.fanid Then
                            oDistribution.id = count
                        End If
                    Next

                    oFan.fanid = count
                    count = count + 1
                Next

                FanDB.SaveFans(oFans, ofandistributions)
                Me.FillFanList(oFans)
            End If
        End If
        Exit Sub
    End Sub

    Private Sub cmdDist_fanreliability_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_fanreliability.Click
        Dim param As String = "Mechanical Ventilation Reliability"
        Dim units As String = "-"
        Dim instruction As String = txtFanReliability.Tag

        Call frmDistParam.ShowDistributionForm(param, units, instruction)
    End Sub
End Class