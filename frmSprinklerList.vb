Imports System
Imports System.Collections.Generic

Public Class frmSprinklerList
    Inherits System.Windows.Forms.Form
    Public NewSprinklerForm As frmNewSprinkler
    Dim oSprinklers As List(Of oSprinkler)
    Dim osprdistributions As List(Of oDistribution)
    Dim odistributions As List(Of oDistribution)


    Private Sub frmSprinklerList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sep As String = vbTab
        Dim text As String
        oSprinklers = SprinklerDB.GetSprinklers2
        Me.FillSprinklerList(oSprinklers)

        text = "ID" & sep & "Room" & sep & "x (m)" & sep & "y (m)"
        ListBox1.Items.Add(text)

        If VM2 = True Then
            chkCalcSprinkRadialDist.Checked = UNCHECKED
            chkCalcSprinkRadialDist.Enabled = False
            txtSprSuppressProb.Text = 0
            txtSprSuppressProb.Enabled = False
            txtSprSuppressProb.BackColor = Color.White
            cmdDist_sprsuppression.Enabled = False
            txtSprCoolingCoeff.Text = 1
            txtSprCoolingCoeff.Enabled = False
            txtSprCoolingCoeff.BackColor = Color.White
            cmdDist_sprcooling.Enabled = False
            cmdNoSprink.Enabled = False
            sprnum_prob(0) = 1
            sprnum_prob(1) = 0
            sprnum_prob(2) = 0
            sprnum_prob(3) = 0

            osprdistributions = SprinklerDB.GetSprDistributions
            For Each oDistribution In osprdistributions
                oDistribution.distribution = "None"
            Next
            SprinklerDB.SaveSprinklers2(oSprinklers, osprdistributions)

            oDistributions = DistributionClass.GetDistributions
            For Each oDistribution In odistributions
                If oDistribution.varname = "Sprinkler Suppression Probability" Then
                    oDistribution.distribution = "None"
                    oDistribution.varvalue = 0
                End If
                If oDistribution.varname = "Sprinkler Cooling Coefficient" Then
                    oDistribution.distribution = "None"
                    oDistribution.varvalue = 1
                End If
            Next
            DistributionClass.SaveDistributions(odistributions)

        Else
            chkCalcSprinkRadialDist.Enabled = True
            txtSprReliability.Enabled = True
            txtSprSuppressProb.Enabled = True
            txtSprCoolingCoeff.Enabled = True
            cmdDist_sprreliability.Enabled = True
            cmdDist_sprsuppression.Enabled = True
            cmdDist_sprcooling.Enabled = True
            cmdNoSprink.Enabled = True

        End If

        Me.CenterToScreen()

    End Sub

    Public Sub FillSprinklerList(ByVal oSprinklers)
        lstSprinklers.Items.Clear()
        Dim j As Integer = 0

        For Each p As oSprinkler In oSprinklers
            j = j + 1
            lstSprinklers.Items.Add(j.ToString & vbTab & p.GetDisplayText(vbTab))
        Next
    End Sub

    Private Sub cmdAddSprinkler_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddSprinkler.Click
        'add standard response sprinkler

        oSprinklers = SprinklerDB.GetSprinklers2
        Dim osprinkler As oSprinkler
        osprdistributions = SprinklerDB.GetSprDistributions

        NewSprinklerForm = New frmNewSprinkler
        NewSprinklerForm.Text = "New Sprinkler"
        osprinkler = New oSprinkler(68, 1, 135, 0.85, 1, 1, 0.025, 3.25, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

        osprinkler.sprid = oSprinklers.Count + 1
        NewSprinklerForm.Tag = oSprinklers.Count + 1

        If osprinkler IsNot Nothing Then
            oSprinklers.Add(osprinkler)

            NewSprinklerForm.TextBox3.Text = osprinkler.sprx
            NewSprinklerForm.TextBox4.Text = osprinkler.spry

            Dim odistribution As New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "rti"
            odistribution.varvalue = 135
            NewSprinklerForm.txtValue8.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "cfactor"
            odistribution.varvalue = 0.85
            NewSprinklerForm.txtValue9.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "sprdensity"
            odistribution.varvalue = 0.07 * 60
            NewSprinklerForm.txtValue12.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "sprr"
            odistribution.varvalue = 3.25
            NewSprinklerForm.txtValue7.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "sprz"
            odistribution.varvalue = 0.025
            NewSprinklerForm.txtValue11.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "acttemp"
            odistribution.varvalue = 68
            NewSprinklerForm.txtValue10.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            SprinklerDB.SaveSprinklers2(oSprinklers, osprdistributions)
        End If

        Me.FillSprinklerList(oSprinklers)
        lstSprinklers.SelectedIndex = osprinkler.sprid - 1
        NewSprinklerForm.ShowDialog()

    End Sub

    Private Sub cmdRemoveSprinkler_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveSprinkler.Click

        Dim i As Integer = lstSprinklers.SelectedIndex
        oSprinklers = SprinklerDB.GetSprinklers2()

        Dim j As Integer = 0
        If i <> -1 Then
            Dim osprinkler As oSprinkler = oSprinklers(i)
            Dim osprdistributions As List(Of oDistribution)

            Dim message As String = "Are you sure you want to delete " _
                & "Sprinkler " & CStr(i + 1) & "?"
            Dim button As DialogResult = MessageBox.Show(message, _
                "Confirm Delete", MessageBoxButtons.YesNo)
            If button = Windows.Forms.DialogResult.Yes Then
                osprdistributions = SprinklerDB.GetSprDistributions()

here:
                For Each oDistribution In osprdistributions
                    If oDistribution.id = osprinkler.sprid Then
                        osprdistributions.Remove(oDistribution)
                        GoTo here
                    End If
                Next

                oSprinklers.Remove(osprinkler)

                'sort and reindex
                Dim count As Integer = 1
                For Each osprinkler In oSprinklers

                    For Each oDistribution In osprdistributions
                        If oDistribution.id = osprinkler.sprid Then
                            oDistribution.id = count
                        End If
                    Next

                    osprinkler.sprid = count
                    count = count + 1
                Next

                SprinklerDB.SaveSprinklers2(oSprinklers, osprdistributions)
                Me.FillSprinklerList(oSprinklers)
            End If
        End If
        Exit Sub

        '        Dim i As Integer = lstSprinklers.SelectedIndex
        '        Dim j As Integer = 0
        '        If i <> -1 Then
        '            Dim osprinkler As oSprinkler = oSprinklers(i)
        '            Dim message As String = "Are you sure you want to delete " _
        '                & "Sprinkler " & CStr(i + 1) & "?"
        '            Dim button As DialogResult = MessageBox.Show(message, _
        '                "Confirm Delete", MessageBoxButtons.YesNo)
        '            If button = Windows.Forms.DialogResult.Yes Then

        '                osprdistributions = SprinklerDB.GetSprDistributions

        'here:
        '                For Each oDistribution In osprdistributions
        '                    If oDistribution.id = osprinkler.sprid Then
        '                        osprdistributions.Remove(oDistribution)
        '                        GoTo here
        '                    End If
        '                Next

        '                oSprinklers.Remove(osprinkler)

        '                'sort and reindex
        '                Dim count As Integer = 1
        '                For Each osprinkler In oSprinklers

        '                    For Each oDistribution In osprdistributions
        '                        If oDistribution.id = osprinkler.sprid Then
        '                            oDistribution.id = count
        '                        End If
        '                    Next

        '                    osprinkler.sprid = count

        '                    count = count + 1
        '                Next

        '                SprinklerDB.SaveSprinklers2(oSprinklers, osprdistributions)
        '                Me.FillSprinklerList(oSprinklers)
        '            End If
        '        End If
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Dim oDistributions As New List(Of oDistribution)
        oDistributions = DistributionClass.GetDistributions

        For Each oDistribution In oDistributions
            If oDistribution.varname = "Sprinkler Reliability" Then
                oDistribution.varvalue = txtSprReliability.Text
            End If
            If oDistribution.varname = "Sprinkler Suppression Probability" Then
                oDistribution.varvalue = txtSprSuppressProb.Text
            End If
            If oDistribution.varname = "Sprinkler Cooling Coefficient" Then
                oDistribution.varvalue = txtSprCoolingCoeff.Text
            End If
        Next
        DistributionClass.SaveDistributions(oDistributions)

        Me.Close()
    End Sub


    Private Sub cmdEditSprinkler_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEditSprinkler.Click
        Dim i As Integer = lstSprinklers.SelectedIndex
        If i <> -1 Then
            oSprinklers = SprinklerDB.GetSprinklers2()
            Dim osprinkler As oSprinkler = oSprinklers(i)
            osprdistributions = SprinklerDB.GetSprDistributions

            Dim NewSprinklerForm As New frmNewSprinkler

            NewSprinklerForm.Text = "Edit Sprinkler"
            NewSprinklerForm.osprinkler = osprinkler

            If NewSprinklerForm.osprinkler IsNot Nothing Then
                'oSprinklers.Add(NewSprinklerForm.osprinkler)
                NewSprinklerForm.Tag = osprinkler.sprid
                NewSprinklerForm.NumericUpDown1.Value = osprinkler.room
                'NewSprinklerForm.txtValue8.Text = osprinkler.rti
                'NewSprinklerForm.txtValue10.Text = osprinkler.acttemp
                'NewSprinklerForm.txtValue9.Text = osprinkler.cfactor
                'NewSprinklerForm.txtValue12.Text = osprinkler.sprdensity
                'NewSprinklerForm.txtValue7.Text = osprinkler.sprr
                NewSprinklerForm.TextBox3.Text = osprinkler.sprx
                NewSprinklerForm.TextBox4.Text = osprinkler.spry

                For Each x In osprdistributions
                    If x.id = osprinkler.sprid Then
                        If x.varname = "rti" Then
                            NewSprinklerForm.txtValue8.Text = x.varvalue
                            If x.distribution <> "None" Then
                                NewSprinklerForm.txtValue8.BackColor = distbackcolor
                            Else
                                NewSprinklerForm.txtValue8.BackColor = distnobackcolor
                            End If
                        End If
                        If x.varname = "cfactor" Then
                            NewSprinklerForm.txtValue9.Text = x.varvalue
                            If x.distribution <> "None" Then
                                NewSprinklerForm.txtValue9.BackColor = distbackcolor
                            Else
                                NewSprinklerForm.txtValue9.BackColor = distnobackcolor
                            End If
                        End If
                        If x.varname = "sprdensity" Then
                            NewSprinklerForm.txtValue12.Text = x.varvalue
                            If x.distribution <> "None" Then
                                NewSprinklerForm.txtValue12.BackColor = distbackcolor
                            Else
                                NewSprinklerForm.txtValue12.BackColor = distnobackcolor
                            End If
                        End If
                        If x.varname = "sprr" Then
                            NewSprinklerForm.txtValue7.Text = x.varvalue
                            If x.distribution <> "None" Then
                                NewSprinklerForm.txtValue7.BackColor = distbackcolor
                            Else
                                NewSprinklerForm.txtValue7.BackColor = distnobackcolor
                            End If
                        End If
                        If x.varname = "sprz" Then
                            NewSprinklerForm.txtValue11.Text = x.varvalue
                            If x.distribution <> "None" Then
                                NewSprinklerForm.txtValue11.BackColor = distbackcolor
                            Else
                                NewSprinklerForm.txtValue11.BackColor = distnobackcolor
                            End If
                        End If
                        If x.varname = "acttemp" Then
                            NewSprinklerForm.txtValue10.Text = x.varvalue
                            If x.distribution <> "None" Then
                                NewSprinklerForm.txtValue10.BackColor = distbackcolor
                            Else
                                NewSprinklerForm.txtValue10.BackColor = distnobackcolor
                            End If
                        End If
                    End If
                Next


                NewSprinklerForm.ShowDialog()


                'SprinklerDB.SaveSprinklers(oSprinklers)
                'Me.FillSprinklerList()
            End If

        End If
    End Sub


    Private Sub cmdCopySprinkler_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopySprinkler.Click

        Dim oSprinklers As List(Of oSprinkler)
        Dim osprdistributions As List(Of oDistribution)
        Dim i As Integer = lstSprinklers.SelectedIndex
        Dim id As Integer
        Dim j As Integer = 0
        If i <> -1 Then
            oSprinklers = SprinklerDB.GetSprinklers2
            Dim osprinkler As oSprinkler = oSprinklers(i)
            id = osprinkler.sprid
            osprdistributions = SprinklerDB.GetSprDistributions
            Dim odistribution As New oDistribution
            Dim NewsprinklerForm As New frmNewSprinkler

            NewsprinklerForm.Text = "Copy Sprinkler"
            NewsprinklerForm.osprinkler = osprinkler
            Dim counter As Integer
            If NewsprinklerForm.osprinkler IsNot Nothing Then

                counter = 0

                For z = 0 To osprdistributions.Count - 1
                    If osprdistributions(z).id = id Then

                        Dim x As oDistribution = osprdistributions(z)
                        Dim y As New oDistribution(x.varname, x.units, x.distribution, x.varvalue, x.mean, x.variance, x.lbound, x.ubound, x.mode, x.alpha, x.beta)
                        y.id = oSprinklers.Count + 1
                        osprdistributions.Add(y)

                    End If
                Next
                NewsprinklerForm.osprinkler.sprid = oSprinklers.Count + 1
                oSprinklers.Add(NewsprinklerForm.osprinkler)

                SprinklerDB.SaveSprinklers2(oSprinklers, osprdistributions)
                Me.FillSprinklerList(oSprinklers)
            End If

        End If
        Exit Sub

    End Sub


    Private Sub cmdDist_sprreliability_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_sprreliability.Click

        Dim param As String = "Sprinkler Reliability"
        Dim units As String = "-"
        Dim instruction As String = txtSprReliability.Tag

        Call frmDistParam.ShowDistributionForm(param, units, instruction)

    End Sub

    Private Sub txtSprReliability_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSprReliability.TextChanged
        'store sprinkler reliability
        If IsNumeric(txtSprReliability.Text) = True Then
            SprReliability = CDec(txtSprReliability.Text)
        End If

    End Sub

    Private Sub txtSprSuppressProb_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSprSuppressProb.TextChanged
        'store sprinkler suppression probability
        If IsNumeric(txtSprSuppressProb.Text) = True Then
            SprSuppressionProb = CDec(txtSprSuppressProb.Text)
        End If
    End Sub

    Private Sub cmdDist_sprsuppression_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_sprsuppression.Click

        Dim param As String = "Sprinkler Suppression Probability"
        Dim units As String = "-"
        Dim instruction As String = "Probability sprinkler system suppresses the fire given system has activated"

        Call frmDistParam.ShowDistributionForm(param, units, instruction)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNoSprink.Click

        frmDiscreteProb.txtSpr1.Text = sprnum_prob(0)
        frmDiscreteProb.txtSpr2.Text = sprnum_prob(1)
        frmDiscreteProb.txtSpr3.Text = sprnum_prob(2)
        frmDiscreteProb.txtSpr4.Text = sprnum_prob(3)

        Call frmDiscreteProb.Show()

    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            System.Diagnostics.Process.Start(URL1)
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in frmSprinklerList.vb LinkLabel1_LinkClicked")
        End Try
    End Sub

    Private Sub txtSprCoolingCoeff_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSprCoolingCoeff.TextChanged
        If IsNumeric(txtSprCoolingCoeff.Text) = True Then
            SprCooling = CDec(txtSprCoolingCoeff.Text)
        End If
    End Sub

    Private Sub cmdDist_sprcooling_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_sprcooling.Click
        Dim param As String = "Sprinkler Cooling Coefficient"
        Dim units As String = "-"
        Dim instruction As String = "Sprinkler cooling coefficient reduces the mass flow of gases through wall openings"

        Call frmDistParam.ShowDistributionForm(param, units, instruction)

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'add quick response sprinkler

        oSprinklers = SprinklerDB.GetSprinklers2
        Dim osprinkler As oSprinkler
        osprdistributions = SprinklerDB.GetSprDistributions

        NewSprinklerForm = New frmNewSprinkler
        NewSprinklerForm.Text = "New Sprinkler"
        osprinkler = New oSprinkler(68, 1, 50, 0.65, 1, 1, 0.025, 3.25, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

        osprinkler.sprid = oSprinklers.Count + 1
        NewSprinklerForm.Tag = oSprinklers.Count + 1

        If osprinkler IsNot Nothing Then
            oSprinklers.Add(osprinkler)

            NewSprinklerForm.TextBox3.Text = osprinkler.sprx
            NewSprinklerForm.TextBox4.Text = osprinkler.spry

            Dim odistribution As New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "rti"
            odistribution.varvalue = 50
            NewSprinklerForm.txtValue8.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "cfactor"
            odistribution.varvalue = 0.65
            NewSprinklerForm.txtValue9.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "sprdensity"
            odistribution.varvalue = 0.07 * 60
            NewSprinklerForm.txtValue12.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "sprr"
            odistribution.varvalue = 3.25
            NewSprinklerForm.txtValue7.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "sprz"
            odistribution.varvalue = 0.025
            NewSprinklerForm.txtValue11.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "acttemp"
            odistribution.varvalue = 68
            NewSprinklerForm.txtValue10.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            SprinklerDB.SaveSprinklers2(oSprinklers, osprdistributions)
        End If

        Me.FillSprinklerList(oSprinklers)
        lstSprinklers.SelectedIndex = osprinkler.sprid - 1
        NewSprinklerForm.ShowDialog()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        'add extended coverage sprinkler

        oSprinklers = SprinklerDB.GetSprinklers2
        Dim osprinkler As oSprinkler
        osprdistributions = SprinklerDB.GetSprDistributions

        NewSprinklerForm = New frmNewSprinkler
        NewSprinklerForm.Text = "New Sprinkler"
        osprinkler = New oSprinkler(68, 1, 50, 0.65, 1, 1, 0.025, 4.3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

        osprinkler.sprid = oSprinklers.Count + 1
        NewSprinklerForm.Tag = oSprinklers.Count + 1

        If osprinkler IsNot Nothing Then
            oSprinklers.Add(osprinkler)

            NewSprinklerForm.TextBox3.Text = osprinkler.sprx
            NewSprinklerForm.TextBox4.Text = osprinkler.spry

            Dim odistribution As New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "rti"
            odistribution.varvalue = 50
            NewSprinklerForm.txtValue8.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "cfactor"
            odistribution.varvalue = 0.65
            NewSprinklerForm.txtValue9.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "sprdensity"
            odistribution.varvalue = 0.07 * 60
            NewSprinklerForm.txtValue12.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "sprr"
            odistribution.varvalue = 4.3
            NewSprinklerForm.txtValue7.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "sprz"
            odistribution.varvalue = 0.025
            NewSprinklerForm.txtValue11.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "acttemp"
            odistribution.varvalue = 68
            NewSprinklerForm.txtValue10.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            SprinklerDB.SaveSprinklers2(oSprinklers, osprdistributions)
        End If

        Me.FillSprinklerList(oSprinklers)
        lstSprinklers.SelectedIndex = osprinkler.sprid - 1
        NewSprinklerForm.ShowDialog()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'add heat detector

        oSprinklers = SprinklerDB.GetSprinklers2
        Dim osprinkler As oSprinkler
        osprdistributions = SprinklerDB.GetSprDistributions

        NewSprinklerForm = New frmNewSprinkler
        NewSprinklerForm.Text = "New Heat Detector"
        osprinkler = New oSprinkler(57, 1, 30, 0, 1, 1, 0.025, 4.2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

        osprinkler.sprid = oSprinklers.Count + 1
        NewSprinklerForm.Tag = oSprinklers.Count + 1

        If osprinkler IsNot Nothing Then
            oSprinklers.Add(osprinkler)

            NewSprinklerForm.TextBox3.Text = osprinkler.sprx
            NewSprinklerForm.TextBox4.Text = osprinkler.spry

            Dim odistribution As New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "rti"
            odistribution.varvalue = 30
            NewSprinklerForm.txtValue8.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "cfactor"
            odistribution.varvalue = 0
            NewSprinklerForm.txtValue9.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "sprdensity"
            odistribution.varvalue = 0
            NewSprinklerForm.txtValue12.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "sprr"
            odistribution.varvalue = 4.2
            NewSprinklerForm.txtValue7.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "sprz"
            odistribution.varvalue = 0.025
            NewSprinklerForm.txtValue11.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = osprinkler.sprid
            odistribution.varname = "acttemp"
            odistribution.varvalue = 57
            NewSprinklerForm.txtValue10.Text = odistribution.varvalue
            osprdistributions.Add(odistribution)

            SprinklerDB.SaveSprinklers2(oSprinklers, osprdistributions)
        End If

        Me.FillSprinklerList(oSprinklers)
        lstSprinklers.SelectedIndex = osprinkler.sprid - 1
        NewSprinklerForm.ShowDialog()

    End Sub


    Private Sub chkCalcSprinkRadialDist_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCalcSprinkRadialDist.CheckedChanged
        If chkCalcSprinkRadialDist.Checked = True Then
            calc_sprdist = True
        Else
            calc_sprdist = False
        End If
    End Sub

    Private Sub cmdRemoveAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveAll.Click

        oSprinklers = SprinklerDB.GetSprinklers2()

        Dim icount As Integer = oSprinklers.Count

        Dim j As Integer = 0
        If icount > 0 Then

            Dim osprdistributions As List(Of oDistribution)

            Dim message As String = "Are you sure you want to delete " _
                & "all sensors " & "?"
            Dim button As DialogResult = MessageBox.Show(message, _
                "Confirm Delete", MessageBoxButtons.YesNo)
            If button = Windows.Forms.DialogResult.Yes Then

                oSprinklers.Clear()

                osprdistributions = SprinklerDB.GetSprDistributions()
                osprdistributions.Clear()

                SprinklerDB.SaveSprinklers2(oSprinklers, osprdistributions)
                Me.FillSprinklerList(oSprinklers)
            End If
        End If
    End Sub

    Private Sub txtSprReliability_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSprReliability.Validated
        ErrorProvider1.Clear()
        SprReliability = CSng(Me.txtSprReliability.Text)

    End Sub

    Private Sub txtSprReliability_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtSprReliability.Validating
        If IsNumeric(txtSprReliability.Text) Then
            If (CSng(txtSprReliability.Text) >= 0) And (CSng(txtSprReliability.Text) <= 1) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtSprReliability.Select(0, txtSprReliability.Text.Length)

        ' Give the ErrorProvider the error message to display.
        ErrorProvider1.SetError(txtSprReliability, "Invalid Entry. Value must be between 0 and 1.")
    End Sub

    Private Sub txtSprSuppressProb_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSprSuppressProb.Validated
        ErrorProvider1.Clear()
        SprSuppressionProb = CSng(Me.txtSprSuppressProb.Text)
    End Sub

    Private Sub txtSprSuppressProb_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtSprSuppressProb.Validating
        If IsNumeric(txtSprSuppressProb.Text) Then
            If (CSng(txtSprSuppressProb.Text) >= 0) And (CSng(txtSprSuppressProb.Text) <= 1) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtSprSuppressProb.Select(0, txtSprSuppressProb.Text.Length)

        ' Give the ErrorProvider the error message to display.
        ErrorProvider1.SetError(txtSprSuppressProb, "Invalid Entry. Value must be between 0 and 1.")

    End Sub

    Private Sub txtSprCoolingCoeff_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSprCoolingCoeff.Validated
        ErrorProvider1.Clear()
        SprCooling = CSng(Me.txtSprCoolingCoeff.Text)
    End Sub

    Private Sub txtSprCoolingCoeff_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtSprCoolingCoeff.Validating
        If IsNumeric(txtSprCoolingCoeff.Text) Then
            If (CSng(txtSprCoolingCoeff.Text) >= 0) And (CSng(txtSprCoolingCoeff.Text) <= 1) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtSprCoolingCoeff.Select(0, txtSprCoolingCoeff.Text.Length)

        ' Give the ErrorProvider the error message to display.
        ErrorProvider1.SetError(txtSprCoolingCoeff, "Invalid Entry. Value must be between 0 and 1.")


    End Sub
End Class