Imports System.Collections.Generic

Public Class frmNewSprinkler

    Public osprinkler As oSprinkler
    Public oSprinklers As List(Of oSprinkler)
    Public osprdistribution As oDistribution
    Public osprdistributions As List(Of oDistribution)

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        If IsValidData() Then
            oSprinklers = SprinklerDB.GetSprinklers2
            osprdistributions = SprinklerDB.GetSprDistributions

            For Each Me.osprinkler In oSprinklers

                If Me.osprinkler.sprid = Me.Tag Then
                    osprinkler.room = NumericUpDown1.Value
                    osprinkler.sprx = TextBox3.Text
                    osprinkler.spry = TextBox4.Text

                    For Each x In osprdistributions
                        If x.id = Me.Tag Then
                            If x.varname = "rti" Then x.varvalue = txtValue8.Text
                            If x.varname = "acttemp" Then x.varvalue = txtValue10.Text
                            If x.varname = "cfactor" Then x.varvalue = txtValue9.Text
                            If x.varname = "sprdensity" Then x.varvalue = txtValue12.Text
                            If x.varname = "sprr" Then x.varvalue = txtValue7.Text
                            If x.varname = "sprz" Then x.varvalue = txtValue11.Text
                        End If
                    Next
                End If
            Next

            If osprinkler IsNot Nothing Then
                SprinklerDB.SaveSprinklers2(oSprinklers, osprdistributions)
                frmSprinklerList.FillSprinklerList(oSprinklers)
            End If

        End If
        Me.Close()

    End Sub
    Private Function IsValidData() As Boolean
        Return Validator.IsPresent(txtValue8) AndAlso _
               Validator.IsPresent(txtValue9) AndAlso _
               Validator.IsPresent(TextBox3) AndAlso _
               Validator.IsPresent(TextBox4) AndAlso _
               Validator.IsPresent(txtValue11) AndAlso _
               Validator.IsPresent(txtValue12) AndAlso _
               Validator.IsPresent(txtValue7) AndAlso _
               Validator.IsPresent(txtValue10)
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub frmNewSprinkler_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        NumericUpDown1.Maximum = NumberRooms
    End Sub


    Private Sub cmdDist_rti_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_rti.Click

        Dim osprinklers As List(Of oSprinkler)
        Dim osprdistributions As List(Of oDistribution)
        Dim sprid = Me.Tag - 1
        Dim param As String = "rti"
        Dim units As String = "(ms)^(1/2)"
        Dim instruction As String = "Sprinkler Response Time Index"
        Dim paramdist As String

        Try

            osprinklers = SprinklerDB.GetSprinklers2()
            osprdistributions = SprinklerDB.GetSprDistributions()

            Call frmDistParam.ShowSprinklerDistributionForm(param, units, instruction, osprinklers, osprdistributions, sprid, paramdist)

            For Each odistribution In osprdistributions
                If odistribution.id = sprid + 1 And odistribution.varname = param Then
                    txtValue8.Text = odistribution.varvalue
                    If paramdist <> "None" Then
                        txtValue8.BackColor = distbackcolor
                    Else
                        txtValue8.BackColor = distnobackcolor
                    End If
                    Exit For
                End If
            Next

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "cmdDist_rti_Click")
        End Try

    End Sub

   
    Private Sub cmdDist_acttemp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_acttemp.Click
        Dim osprinklers As List(Of oSprinkler)
        Dim osprdistributions As List(Of oDistribution)
        Dim sprid = Me.Tag - 1
        Dim param As String = "acttemp"
        Dim units As String = "(C)"
        Dim instruction As String = "Sprinkler Actuation Temperature"
        Dim paramdist As String

        Try

            osprinklers = SprinklerDB.GetSprinklers2()
            osprdistributions = SprinklerDB.GetSprDistributions()

            Call frmDistParam.ShowSprinklerDistributionForm(param, units, instruction, osprinklers, osprdistributions, sprid, paramdist)

            For Each odistribution In osprdistributions
                If odistribution.id = sprid + 1 And odistribution.varname = param Then
                    txtValue10.Text = odistribution.varvalue
                    If paramdist <> "None" Then
                        txtValue10.BackColor = distbackcolor
                    Else
                        txtValue10.BackColor = distnobackcolor
                    End If
                    Exit For
                End If
            Next

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "cmdDist_acttemp_Click")
        End Try

    End Sub

    Private Sub cmdDist_cfactor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_cfactor.Click

        Dim osprinklers As List(Of oSprinkler)
        Dim osprdistributions As List(Of oDistribution)
        Dim sprid = Me.Tag - 1
        Dim param As String = "cfactor"
        Dim units As String = "(m/s)^(1/2)"
        Dim instruction As String = "Sprinkler Conduction Factor"
        Dim paramdist As String

        Try

            osprinklers = SprinklerDB.GetSprinklers2()
            osprdistributions = SprinklerDB.GetSprDistributions()

            Call frmDistParam.ShowSprinklerDistributionForm(param, units, instruction, osprinklers, osprdistributions, sprid, paramdist)

            For Each odistribution In osprdistributions
                If odistribution.id = sprid + 1 And odistribution.varname = param Then
                    txtValue9.Text = odistribution.varvalue
                    If paramdist <> "None" Then
                        txtValue9.BackColor = distbackcolor
                    Else
                        txtValue9.BackColor = distnobackcolor
                    End If
                    Exit For
                End If
            Next

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "cmdDist_cfactor_Click")
        End Try
    End Sub

    Private Sub cmdDist_sprdensity_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_sprdensity.Click
        Dim osprinklers As List(Of oSprinkler)
        Dim osprdistributions As List(Of oDistribution)
        Dim sprid = Me.Tag - 1
        Dim param As String = "sprdensity"
        Dim units As String = "(mm/min)"
        Dim instruction As String = "Sprinkler Water Spray Density"
        Dim paramdist As String

        Try

            osprinklers = SprinklerDB.GetSprinklers2()
            osprdistributions = SprinklerDB.GetSprDistributions()

            Call frmDistParam.ShowSprinklerDistributionForm(param, units, instruction, osprinklers, osprdistributions, sprid, paramdist)

            For Each odistribution In osprdistributions
                If odistribution.id = sprid + 1 And odistribution.varname = param Then
                    txtValue12.Text = odistribution.varvalue
                    If paramdist <> "None" Then
                        txtValue12.BackColor = distbackcolor
                    Else
                        txtValue12.BackColor = distnobackcolor
                    End If
                    Exit For
                End If
            Next

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "cmdDist_sprdensity_Click")
        End Try
    End Sub

    Private Sub cmdDist_sprr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_sprr.Click
        Dim osprinklers As List(Of oSprinkler)
        Dim osprdistributions As List(Of oDistribution)
        Dim sprid = Me.Tag - 1
        Dim param As String = "sprr"
        Dim units As String = "(m)"
        Dim instruction As String = "Sprinkler Radial Distance"
        Dim paramdist As String

        Try

            osprinklers = SprinklerDB.GetSprinklers2()
            osprdistributions = SprinklerDB.GetSprDistributions()

            Call frmDistParam.ShowSprinklerDistributionForm(param, units, instruction, osprinklers, osprdistributions, sprid, paramdist)

            For Each odistribution In osprdistributions
                If odistribution.id = sprid + 1 And odistribution.varname = param Then
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

    Private Sub cmdDist_sprz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_sprz.Click
        Dim osprinklers As List(Of oSprinkler)
        Dim osprdistributions As List(Of oDistribution)
        Dim sprid = Me.Tag - 1
        Dim param As String = "sprz"
        Dim units As String = "(m)"
        Dim instruction As String = "Sprinkler Distance Below Ceiling"
        Dim paramdist As String

        Try

            osprinklers = SprinklerDB.GetSprinklers2()
            osprdistributions = SprinklerDB.GetSprDistributions()

            Call frmDistParam.ShowSprinklerDistributionForm(param, units, instruction, osprinklers, osprdistributions, sprid, paramdist)

            For Each odistribution In osprdistributions
                If odistribution.id = sprid + 1 And odistribution.varname = param Then
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
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "cmdDist_sprz_Click")
        End Try
    End Sub

End Class