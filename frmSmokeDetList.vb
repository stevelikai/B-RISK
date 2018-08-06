Imports System
Imports System.Collections.Generic

Public Class frmSmokeDetList
    Inherits System.Windows.Forms.Form
    Public NewSmokeDetForm As frmNewSmokeDetector
    Dim oSmokeDets As List(Of oSmokeDet)
    Dim oSDdistributions As List(Of oDistribution)
    Private Sub frmSmokeDetList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sep As String = vbTab
        Dim text As String
        oSmokeDets = SmokeDetDB.GetSmokDets
        Me.FillSmokeDetList(oSmokeDets)

        text = "ID" & sep & "Room" & sep & "x (m)" & sep & "y (m)"
        ListBox1.Items.Add(text)

        If VM2 = True Then
            chkCalcSRadialDist.Checked = False
            chkCalcSRadialDist.Enabled = False
        Else
            chkCalcSRadialDist.Enabled = True
        End If

        Me.CenterToScreen()

    End Sub
    Public Sub FillSmokeDetList(ByVal oSmokeDets)
        lstSmokeDet.Items.Clear()
        Dim j As Integer = 0

        For Each p As oSmokeDet In oSmokeDets
            j = j + 1
            lstSmokeDet.Items.Add(j.ToString & vbTab & p.GetDisplayText(vbTab))
        Next

    End Sub

    Private Sub cmdAddSmokeDet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddSmokeDet.Click
        'add sd

        oSmokeDets = SmokeDetDB.GetSmokDets
        Dim oSmokeDet As oSmokeDet
        oSDdistributions = SmokeDetDB.GetSDDistributions

        NewSmokeDetForm = New frmNewSmokeDetector
        NewSmokeDetForm.Text = "New Smoke Detector"
        oSmokeDet = New oSmokeDet(1, RoomLength(1) / 2, RoomWidth(1) / 2, 0.025, 0, 0.097, 15, True, False, 12, 0.2)

        oSmokeDet.sdid = oSmokeDets.Count + 1
        NewSmokeDetForm.Tag = oSmokeDets.Count + 1

        If oSmokeDet IsNot Nothing Then
            oSmokeDets.Add(oSmokeDet)

            NewSmokeDetForm.TextBox3.Text = oSmokeDet.sdx
            NewSmokeDetForm.TextBox4.Text = oSmokeDet.sdy

            Dim oDistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            oDistribution.id = oSmokeDet.sdid
            oDistribution.varname = "sdr"
            oDistribution.varvalue = 7.0
            NewSmokeDetForm.txtValue7.Text = oDistribution.varvalue
            oSDdistributions.Add(oDistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            oDistribution.id = oSmokeDet.sdid
            oDistribution.varname = "sdz"
            odistribution.varvalue = 0.025
            NewSmokeDetForm.txtValue11.Text = oDistribution.varvalue
            oSDdistributions.Add(oDistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            oDistribution.id = oSmokeDet.sdid
            oDistribution.varname = "OD"
            oDistribution.varvalue = 0.097
            NewSmokeDetForm.txtOD.Text = oDistribution.varvalue
            oSDdistributions.Add(oDistribution)

            oDistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            oDistribution.id = oSmokeDet.sdid
            oDistribution.varname = "charlength"
            oDistribution.varvalue = 15
            NewSmokeDetForm.txtCharLength.Text = oDistribution.varvalue
            oSDdistributions.Add(oDistribution)

            SmokeDetDB.SaveSmokeDets(oSmokeDets, oSDdistributions)
        End If

        Me.FillSmokeDetList(oSmokeDets)
        lstSmokeDet.SelectedIndex = oSmokeDet.sdid - 1

        'show the beam detection inputs
        If DevKey = True Then
            NewSmokeDetForm.gbBeamDetect.Visible = True
        Else
            NewSmokeDetForm.gbBeamDetect.Visible = False
        End If

        NewSmokeDetForm.ShowDialog()

    End Sub

    Private Sub cmdRemoveSmokeDet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveSmokeDet.Click
        Dim i As Integer = lstSmokeDet.SelectedIndex
        oSmokeDets = SmokeDetDB.GetSmokDets()

        Dim j As Integer = 0
        If i <> -1 Then
            Dim oSmokeDet As oSmokeDet = oSmokeDets(i)
            Dim osddistributions As List(Of oDistribution)

            Dim message As String = "Are you sure you want to delete " _
                & "Smoke Detector " & CStr(i + 1) & "?"
            Dim button As DialogResult = MessageBox.Show(message, _
                "Confirm Delete", MessageBoxButtons.YesNo)
            If button = Windows.Forms.DialogResult.Yes Then
                osddistributions = SmokeDetDB.GetSDDistributions()

here:
                For Each oDistribution In osddistributions
                    If oDistribution.id = oSmokeDet.sdid Then
                        osddistributions.Remove(oDistribution)
                        GoTo here
                    End If
                Next

                oSmokeDets.Remove(oSmokeDet)

                'sort and reindex
                Dim count As Integer = 1
                For Each oSmokeDet In oSmokeDets

                    For Each oDistribution In osddistributions
                        If oDistribution.id = oSmokeDet.sdid Then
                            oDistribution.id = count
                        End If
                    Next

                    oSmokeDet.sdid = count
                    count = count + 1
                Next

                SmokeDetDB.SaveSmokeDets(oSmokeDets, osddistributions)
                Me.FillSmokeDetList(oSmokeDets)
            End If
        End If
        Exit Sub

    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Dim oDistributions As New List(Of oDistribution)
        oDistributions = DistributionClass.GetDistributions

        For Each oDistribution In oDistributions
            If oDistribution.varname = "Smoke Detector Reliability" Then
                oDistribution.varvalue = txtSmokeDetReliability.Text
            End If
        Next
        DistributionClass.SaveDistributions(oDistributions)

        Me.Close()
    End Sub

    Private Sub cmdEditSmokeDet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEditSmokeDet.Click
        Dim i As Integer = lstSmokeDet.SelectedIndex
        If i <> -1 Then
            oSmokeDets = SmokeDetDB.GetSmokDets()
            Dim oSmokeDet As oSmokeDet = oSmokeDets(i)
            oSDdistributions = SmokeDetDB.GetSDDistributions

            Dim NewSmokeDetForm As New frmNewSmokeDetector

            NewSmokeDetForm.Text = "Edit Smoke Detector"
            NewSmokeDetForm.oSmokeDet = oSmokeDet

            If NewSmokeDetForm.oSmokeDet IsNot Nothing Then
                NewSmokeDetForm.Tag = oSmokeDet.sdid
                NewSmokeDetForm.NumericUpDown1.Value = oSmokeDet.room
                NewSmokeDetForm.TextBox3.Text = oSmokeDet.sdx
                NewSmokeDetForm.TextBox4.Text = oSmokeDet.sdy
                'NewSmokeDetForm.txtCharLength.Text = oSmokeDet.charlength
                NewSmokeDetForm.chkSDinside.Checked = oSmokeDet.sdinside
                NewSmokeDetForm.txtPathLength.Text = oSmokeDet.sdbeampathlength
                NewSmokeDetForm.txtAlarmTrans.Text = oSmokeDet.sdbeamalarmtrans
                If oSmokeDet.sdbeam = True Then
                    NewSmokeDetForm.optBeamDetect.Checked = True
                Else
                    NewSmokeDetForm.optBeamDetect.Checked = False
                End If

                For Each x In oSDdistributions
                    If x.id = oSmokeDet.sdid Then
                        If x.varname = "charlength" Then
                            NewSmokeDetForm.txtCharLength.Text = x.varvalue
                            If x.distribution <> "None" Then
                                NewSmokeDetForm.txtCharLength.BackColor = distbackcolor
                            Else
                                NewSmokeDetForm.txtCharLength.BackColor = distnobackcolor
                            End If
                        End If
                        If x.varname = "OD" Then
                            NewSmokeDetForm.txtOD.Text = x.varvalue
                            If x.distribution <> "None" Then
                                NewSmokeDetForm.txtOD.BackColor = distbackcolor
                            Else
                                NewSmokeDetForm.txtOD.BackColor = distnobackcolor
                            End If
                        End If
                        If x.varname = "sdr" Then
                            NewSmokeDetForm.txtValue7.Text = x.varvalue
                            If x.distribution <> "None" Then
                                NewSmokeDetForm.txtValue7.BackColor = distbackcolor
                            Else
                                NewSmokeDetForm.txtValue7.BackColor = distnobackcolor
                            End If
                        End If
                        If x.varname = "sdz" Then
                            NewSmokeDetForm.txtValue11.Text = x.varvalue
                            If x.distribution <> "None" Then
                                NewSmokeDetForm.txtValue11.BackColor = distbackcolor
                            Else
                                NewSmokeDetForm.txtValue11.BackColor = distnobackcolor
                            End If
                        End If
                    End If
                Next

                'show the beam detection inputs
                If DevKey = True Then
                    NewSmokeDetForm.gbBeamDetect.Visible = True
                Else
                    NewSmokeDetForm.gbBeamDetect.Visible = False
                End If


                NewSmokeDetForm.ShowDialog()


            End If

        End If
    End Sub

    Private Sub cmdCopySmokeDet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopySmokeDet.Click
        Dim oSmokeDets As List(Of oSmokeDet)
        Dim osddistributions As List(Of oDistribution)
        Dim i As Integer = lstSmokeDet.SelectedIndex
        Dim id As Integer
        Dim j As Integer = 0
        If i <> -1 Then
            oSmokeDets = SmokeDetDB.GetSmokDets
            Dim oSmokeDet As oSmokeDet = oSmokeDets(i)
            id = oSmokeDet.sdid
            osddistributions = SmokeDetDB.GetSDDistributions
            Dim odistribution As New oDistribution
            Dim NewSmokeDetForm As New frmNewSmokeDetector

            NewSmokeDetForm.Text = "Copy Smoke Detector"
            NewSmokeDetForm.oSmokeDet = oSmokeDet
            Dim counter As Integer
            If NewSmokeDetForm.oSmokeDet IsNot Nothing Then

                counter = 0

                For z = 0 To osddistributions.Count - 1
                    If osddistributions(z).id = id Then

                        Dim x As oDistribution = osddistributions(z)
                        Dim y As New oDistribution(x.varname, x.units, x.distribution, x.varvalue, x.mean, x.variance, x.lbound, x.ubound, x.mode, x.alpha, x.beta)
                        y.id = oSmokeDets.Count + 1
                        osddistributions.Add(y)

                    End If
                Next
                NewSmokeDetForm.oSmokeDet.sdid = oSmokeDets.Count + 1
                oSmokeDets.Add(NewSmokeDetForm.oSmokeDet)

                SmokeDetDB.SaveSmokeDets(oSmokeDets, osddistributions)
                Me.FillSmokeDetList(oSmokeDets)
            End If

        End If
        Exit Sub

    End Sub

    Private Sub cmdDist_sprreliability_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_sprreliability.Click
        Dim param As String = "Smoke Detector Reliability"
        Dim units As String = "-"
        Dim instruction As String = txtSmokeDetReliability.Tag

        Call frmDistParam.ShowDistributionForm(param, units, instruction)
    End Sub

    Private Sub txtSmokeDetReliability_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSmokeDetReliability.TextChanged
        'store sd reliability
        If IsNumeric(txtSmokeDetReliability.Text) = True Then
            SDReliability = CDec(txtSmokeDetReliability.Text)
        End If
    End Sub

    Private Sub chkCalcSRadialDist_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCalcSRadialDist.CheckedChanged
        If chkCalcSRadialDist.Checked = True Then
            calc_sddist = True
        Else
            calc_sddist = False
        End If
    End Sub

    Private Sub txtSmokeDetReliability_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSmokeDetReliability.Validated
        ErrorProvider1.Clear()
        SDReliability = CSng(Me.txtSmokeDetReliability.Text)

    End Sub

    Private Sub txtSmokeDetReliability_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtSmokeDetReliability.Validating
        If IsNumeric(txtSmokeDetReliability.Text) Then
            If (CSng(txtSmokeDetReliability.Text) >= 0) And (CSng(txtSmokeDetReliability.Text) <= 1) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtSmokeDetReliability.Select(0, txtSmokeDetReliability.Text.Length)

        ' Give the ErrorProvider the error message to display.
        ErrorProvider1.SetError(txtSmokeDetReliability, "Invalid Entry. Value must be between 0 and 1.")

    End Sub
End Class