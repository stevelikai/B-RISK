Imports System.Collections.Generic

Public Class frmNewSmokeDetector
    Public oSmokeDet As oSmokeDet
    Public oSmokeDets As List(Of oSmokeDet)
    Public osddistribution As oDistribution
    Public osddistributions As List(Of oDistribution)
    Private Function IsValidData() As Boolean
        Return Validator.IsPresent(TextBox3) AndAlso _
               Validator.IsPresent(TextBox4) AndAlso _
               Validator.IsPresent(txtValue11) AndAlso _
               Validator.IsPresent(txtValue7) AndAlso _
               Validator.IsPresent(txtOD)
    End Function
    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        If IsValidData() Then
            oSmokeDets = SmokeDetDB.GetSmokDets
            osddistributions = SmokeDetDB.GetSDDistributions

            For Each Me.oSmokeDet In oSmokeDets

                If Me.oSmokeDet.sdid = Me.Tag Then
                    oSmokeDet.room = NumericUpDown1.Value
                    oSmokeDet.sdx = TextBox3.Text
                    oSmokeDet.sdy = TextBox4.Text

                    oSmokeDet.sdinside = chkSDinside.CheckState
                    'oSmokeDet.charlength = txtCharLength.Text
                    oSmokeDet.sdbeamalarmtrans = txtAlarmTrans.Text
                    oSmokeDet.sdbeampathlength = txtPathLength.Text

                    If Me.optBeamDetect.Checked = True Then
                        oSmokeDet.sdbeam = True
                    Else
                        oSmokeDet.sdbeam = False
                    End If

                    oSmokeDet.sdinside = chkSDinside.CheckState

                    For Each x In osddistributions
                        If x.id = Me.Tag Then
                            If x.varname = "sdr" Then x.varvalue = txtValue7.Text
                            If x.varname = "sdz" Then x.varvalue = txtValue11.Text
                            If x.varname = "OD" Then x.varvalue = txtOD.Text
                            If x.varname = "charlength" Then x.varvalue = txtCharLength.Text
                        End If
                    Next
                End If
            Next

            If oSmokeDet IsNot Nothing Then
                SmokeDetDB.SaveSmokeDets(oSmokeDets, osddistributions)
                frmSmokeDetList.FillSmokeDetList(oSmokeDets)
            End If

        End If
        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub frmNewSmokeDetector_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        NumericUpDown1.Maximum = NumberRooms
        If VM2 = True Then
            chkSDinside.Checked = False
            chkSDinside.Enabled = False
        Else
            chkSDinside.Enabled = True
        End If
    End Sub

    Private Sub cmdDist_sdr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_sdr.Click

        Dim oSmokeDets As List(Of oSmokeDet)
        Dim osddistributions As List(Of oDistribution)
        Dim sdid = Me.Tag - 1
        Dim param As String = "sdr"
        Dim units As String = "(m)"
        Dim instruction As String = "Smoke Detector Radial Distance"
        Dim paramdist As String

        Try

            oSmokeDets = SmokeDetDB.GetSmokDets()
            osddistributions = SmokeDetDB.GetSdDistributions()

            Call frmDistParam.ShowSmokeDetDistributionForm(param, units, instruction, oSmokeDets, osddistributions, sdid, paramdist)

            For Each odistribution In osddistributions
                If odistribution.id = sdid + 1 And odistribution.varname = param Then
                    txtValue7.Text = odistribution.varvalue
                    If paramdist <> "None" Then
                        txtValue7.BackColor = distbackcolor
                    Else
                        txtValue7.BackColor = distnobackcolor
                    End If
                    Exit For
                End If
            Next

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "cmdDist_sprr_Click")
        End Try
    End Sub

    Private Sub cmdDist_sdz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_sdz.Click
        Dim oSmokeDets As List(Of oSmokeDet)
        Dim osddistributions As List(Of oDistribution)
        Dim sdid = Me.Tag - 1
        Dim param As String = "sdz"
        Dim units As String = "(m)"
        Dim instruction As String = "Smoke Detector Distance Below Ceiling"
        Dim paramdist As String

        Try

            oSmokeDets = SmokeDetDB.GetSmokDets()
            osddistributions = SmokeDetDB.GetSDDistributions()

            Call frmDistParam.ShowSmokeDetDistributionForm(param, units, instruction, oSmokeDets, osddistributions, sdid, paramdist)

            For Each odistribution In osddistributions
                If odistribution.id = sdid + 1 And odistribution.varname = param Then
                    txtValue11.Text = odistribution.varvalue
                    If paramdist <> "None" Then
                        txtValue11.BackColor = distbackcolor
                    Else
                        txtValue11.BackColor = distnobackcolor
                    End If
                    Exit For
                End If
            Next

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "cmdDist_sdz_Click")
        End Try
    End Sub

    Private Sub cmdDist_OD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_OD.Click
        Dim oSmokeDets As List(Of oSmokeDet)
        Dim osddistributions As List(Of oDistribution)
        Dim sdid = Me.Tag - 1
        Dim param As String = "OD"
        Dim units As String = "(1/m)"
        Dim instruction As String = "Smoke Detector Optical Density for Alarm"
        Dim paramdist As String

        Try

            oSmokeDets = SmokeDetDB.GetSmokDets()
            osddistributions = SmokeDetDB.GetSDDistributions()

            Call frmDistParam.ShowSmokeDetDistributionForm(param, units, instruction, oSmokeDets, osddistributions, sdid, paramdist)

            For Each odistribution In osddistributions
                If odistribution.id = sdid + 1 And odistribution.varname = param Then
                    txtOD.Text = odistribution.varvalue
                    If paramdist <> "None" Then
                        txtOD.BackColor = distbackcolor
                    Else
                        txtOD.BackColor = distnobackcolor
                    End If
                    Exit For
                End If
            Next

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "cmdDist_OD_Click")
        End Try
    End Sub

    Private Sub cmdDist_charlength_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_charlength.Click
        Dim oSmokeDets As List(Of oSmokeDet)
        Dim osddistributions As List(Of oDistribution)
        Dim sdid = Me.Tag - 1
        Dim param As String = "charlength"
        Dim units As String = "(m)"
        Dim instruction As String = "Characteristic Length"
        Dim paramdist As String

        Try

            oSmokeDets = SmokeDetDB.GetSmokDets()
            osddistributions = SmokeDetDB.GetSDDistributions()

            Call frmDistParam.ShowSmokeDetDistributionForm(param, units, instruction, oSmokeDets, osddistributions, sdid, paramdist)

            For Each odistribution In osddistributions
                If odistribution.id = sdid + 1 And odistribution.varname = param Then
                    txtCharLength.Text = odistribution.varvalue
                    If paramdist <> "None" Then
                        txtCharLength.BackColor = distbackcolor
                    Else
                        txtCharLength.BackColor = distnobackcolor
                    End If
                    Exit For
                End If
            Next

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "cmdDist_charlength_Click")
        End Try
    End Sub

    Private Sub optBeamDetect_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optBeamDetect.CheckedChanged
        If optBeamDetect.Checked = True Then
            Me.txtCharLength.Visible = False
            Me.txtOD.Visible = False
            Me.TextBox3.Visible = False
            Me.TextBox4.Visible = False
            Me.txtValue7.Visible = False
            Me.Label3.Visible = False
            Me.Label5.Visible = False
            Me.Label6.Visible = False
            Me.Label30.Visible = False
            Me.Label1.Visible = False
            Me.Label31.Visible = False
            Me.Label33.Visible = False
            Me.Label2.Visible = False
            Me.Label11.Visible = False
            Me.Label12.Visible = False
            Me.cmdDist_charlength.Visible = False
            Me.cmdDist_OD.Visible = False
            Me.cmdDist_sdr.Visible = False
            Me.txtAlarmTrans.Visible = True
            Me.txtPathLength.Visible = True
            Me.lblAlarmTrans.Visible = True
            Me.lblPathLength.Visible = True
            Me.chkSDinside.Visible = False
        Else
            Me.txtCharLength.Visible = True
            Me.txtOD.Visible = True
            Me.TextBox3.Visible = True
            Me.TextBox4.Visible = True
            Me.txtValue7.Visible = True
            Me.Label3.Visible = True
            Me.Label5.Visible = True
            Me.Label6.Visible = True
            Me.Label30.Visible = True
            Me.Label1.Visible = True
            Me.Label31.Visible = True
            Me.Label33.Visible = True
            Me.Label2.Visible = True
            Me.Label11.Visible = True
            Me.Label12.Visible = True
            Me.cmdDist_charlength.Visible = True
            Me.cmdDist_OD.Visible = True
            Me.cmdDist_sdr.Visible = True
            Me.txtAlarmTrans.Visible = False
            Me.txtPathLength.Visible = False
            Me.lblAlarmTrans.Visible = False
            Me.lblPathLength.Visible = False
            Me.chkSDinside.Visible = True
        End If
    End Sub

    Private Sub optPointDetect_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optPointDetect.CheckedChanged
        If optBeamDetect.Checked = True Then
            Me.txtCharLength.Visible = False
            Me.txtOD.Visible = False
            Me.TextBox3.Visible = False
            Me.TextBox4.Visible = False
            Me.txtValue7.Visible = False
            Me.Label3.Visible = False
            Me.Label5.Visible = False
            Me.Label6.Visible = False
            Me.Label30.Visible = False
            Me.Label1.Visible = False
            Me.Label31.Visible = False
            Me.Label33.Visible = False
            Me.Label2.Visible = False
            Me.Label11.Visible = False
            Me.Label12.Visible = False
            Me.cmdDist_charlength.Visible = False
            Me.cmdDist_OD.Visible = False
            Me.cmdDist_sdr.Visible = False
            Me.txtAlarmTrans.Visible = True
            Me.txtPathLength.Visible = True
            Me.lblAlarmTrans.Visible = True
            Me.lblPathLength.Visible = True
            Me.chkSDinside.Visible = False
        Else
            Me.txtCharLength.Visible = True
            Me.txtOD.Visible = True
            Me.TextBox3.Visible = True
            Me.TextBox4.Visible = True
            Me.txtValue7.Visible = True
            Me.Label3.Visible = True
            Me.Label5.Visible = True
            Me.Label6.Visible = True
            Me.Label30.Visible = True
            Me.Label1.Visible = True
            Me.Label31.Visible = True
            Me.Label33.Visible = True
            Me.Label2.Visible = True
            Me.Label11.Visible = True
            Me.Label12.Visible = True
            Me.cmdDist_charlength.Visible = True
            Me.cmdDist_OD.Visible = True
            Me.cmdDist_sdr.Visible = True
            Me.txtAlarmTrans.Visible = False
            Me.txtPathLength.Visible = False
            Me.lblAlarmTrans.Visible = False
            Me.lblPathLength.Visible = False
            Me.chkSDinside.Visible = True
        End If
    End Sub
End Class