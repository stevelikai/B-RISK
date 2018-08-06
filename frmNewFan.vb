Imports System.Collections.Generic
Public Class frmNewFan
    Public oFan As oFan
    Public oFans As List(Of oFan)
    Public ofandistribution As oDistribution
    Public ofandistributions As List(Of oDistribution)

    Private Function IsValidData() As Boolean
        Return Validator.IsPresent(txtFanElevation) AndAlso _
               Validator.IsPresent(txtFanFlowRate) AndAlso _
               Validator.IsPresent(txtFanStartTime) AndAlso _
               Validator.IsPresent(txtPressureLimit)

    End Function

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        If IsValidData() Then
            oFans = FanDB.GetFans
            ofandistributions = FanDB.GetFanDistributions

            For Each Me.oFan In oFans

                If Me.oFan.fanid = Me.Tag Then
                    oFan.fanroom = NumericUpDownRoom.Value
                    oFan.fanelevation = txtFanElevation.Text
                    oFan.fanflowrate = txtFanFlowRate.Text
                    oFan.fanstarttime = txtFanStartTime.Text
                    oFan.fanpressurelimit = txtPressureLimit.Text
                    oFan.fanextract = optExtract.Checked
                    If optFanManual.Checked = True Then
                        oFan.fanmanual = 0
                    ElseIf optFanAuto.Checked = True Then
                        oFan.fanmanual = 1
                    Else
                        oFan.fanmanual = 2
                    End If

                    oFan.fancurve = chkUseFanCurve.Checked
                    oFan.fanreliability = txtFanReliability.Text

                    For Each x In ofandistributions
                        If x.id = Me.Tag Then
                            If x.varname = "fanflowrate" Then x.varvalue = txtFanFlowRate.Text
                            If x.varname = "fanstarttime" Then x.varvalue = txtFanStartTime.Text
                            If x.varname = "fanpressurelimit" Then x.varvalue = txtPressureLimit.Text
                            If x.varname = "fanreliability" Then x.varvalue = txtFanReliability.Text
                        End If
                    Next
                End If
            Next

            If oFan IsNot Nothing Then
                FanDB.SaveFans(oFans, ofandistributions)
                frmFanList.FillFanList(oFans)
            End If

        End If
        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub


    Private Sub cmdDist_fanflowrate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_fanflowrate.Click
        Dim oFans As List(Of oFan)
        Dim ofandistributions As List(Of oDistribution)
        Dim fanid = Me.Tag
        Dim param As String = "fanflowrate"
        Dim units As String = "(m3/s)"
        Dim instruction As String = "Fan Flow Rate"
        Dim paramdist As String

        Try

            oFans = FanDB.GetFans()
            ofandistributions = FanDB.GetFanDistributions()

            Call frmDistParam.ShowFanDistributionForm(param, units, instruction, oFans, ofandistributions, fanid, paramdist)

            For Each odistribution In ofandistributions
                If odistribution.id = fanid And odistribution.varname = param Then
                    txtFanFlowRate.Text = odistribution.varvalue
                    If paramdist <> "None" Then
                        txtFanFlowRate.BackColor = distbackcolor
                    Else
                        txtFanFlowRate.BackColor = distnobackcolor
                    End If
                    Exit For
                End If
            Next

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "cmdDist_fanflowrate_Click")
        End Try
    End Sub

    Private Sub cmdDist_fanpressurelimit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_fanpressurelimit.Click
        Dim oFans As List(Of oFan)
        Dim ofandistributions As List(Of oDistribution)
        Dim fanid = Me.Tag
        Dim param As String = "fanpressurelimit"
        Dim units As String = "(Pa)"
        Dim instruction As String = "Fan Pressure Limit"
        Dim paramdist As String

        Try

            oFans = FanDB.GetFans()
            ofandistributions = FanDB.GetFanDistributions()

            Call frmDistParam.ShowFanDistributionForm(param, units, instruction, oFans, ofandistributions, fanid, paramdist)

            For Each odistribution In ofandistributions
                If odistribution.id = fanid And odistribution.varname = param Then
                    txtPressureLimit.Text = odistribution.varvalue
                    If paramdist <> "None" Then
                        txtPressureLimit.BackColor = distbackcolor
                    Else
                        txtPressureLimit.BackColor = distnobackcolor
                    End If
                    Exit For
                End If
            Next

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "cmdDist_fanpressurelimit_Click")
        End Try
    End Sub

    Private Sub cmdDist_fanstarttime_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_fanstarttime.Click
        Dim oFans As List(Of oFan)
        Dim ofandistributions As List(Of oDistribution)
        Dim fanid = Me.Tag
        Dim param As String = "fanstarttime"
        Dim units As String = "(sec)"
        Dim instruction As String = "Fan Start Time"
        Dim paramdist As String

        Try

            oFans = FanDB.GetFans()
            ofandistributions = FanDB.GetFanDistributions()

            Call frmDistParam.ShowFanDistributionForm(param, units, instruction, oFans, ofandistributions, fanid, paramdist)

            For Each odistribution In ofandistributions
                If odistribution.id = fanid And odistribution.varname = param Then
                    txtFanStartTime.Text = odistribution.varvalue
                    If paramdist <> "None" Then
                        txtFanStartTime.BackColor = distbackcolor
                    Else
                        txtFanStartTime.BackColor = distnobackcolor
                    End If
                    Exit For
                End If
            Next

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "cmdDist_fanstarttime_Click")
        End Try
    End Sub

    Private Sub cmdDist_fanreliability_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_fanreliability.Click
        Dim oFans As List(Of oFan)
        Dim ofandistributions As List(Of oDistribution)
        Dim fanid = Me.Tag
        Dim param As String = "fanreliability"
        Dim units As String = "-"
        Dim instruction As String = "Fan Reliability"
        Dim paramdist As String

        Try

            oFans = FanDB.GetFans()
            ofandistributions = FanDB.GetFanDistributions()

            Call frmDistParam.ShowFanDistributionForm(param, units, instruction, oFans, ofandistributions, fanid, paramdist)

            For Each odistribution In ofandistributions
                If odistribution.id = fanid And odistribution.varname = param Then
                    txtFanReliability.Text = odistribution.varvalue
                    If paramdist <> "None" Then
                        txtFanReliability.BackColor = distbackcolor
                    Else
                        txtFanReliability.BackColor = distnobackcolor
                    End If
                    Exit For
                End If
            Next

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "cmdDist_fanreliability_Click")
        End Try
    End Sub

    Private Sub frmNewFan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class