Imports System.Collections.Generic
Imports System.Net
Imports System.IO
Imports System.Xml

Public Class frmVentList
    Public NewVentForm As frmNewVents
    Dim oVents As List(Of oVent)
    Dim oVentdistributions As List(Of oDistribution)

    Private Sub frmVentList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim sep As String = vbTab
        Dim text As String
        oVents = VentDB.GetVents()

        Dim lab1 As String = "Description"
        Dim lab2 As String = "Room"
        Dim lab3 As String = "Room"
        Dim lab4 As String = "ID"

        text = LeftJustified(lab4, 3) & _
         RightJustified(lab1, 24) & _
         RightJustified(lab2, 8) & _
        RightJustified(lab3, 8)
        ' text = "ID" & sep & "Description" & sep & "From Room" & sep & "To Room"
        ListBox1.Items.Add(text)
        Me.FillVentList()

        Me.CenterToScreen()
        Me.AddOwnedForm(frmNewVents)

    End Sub

    Public Sub FillVentList()

        lstVents.Items.Clear()

        oVents = VentDB.GetVents()
        Dim j As Integer = 0

        For Each p As oVent In oVents
            j = j + 1

            Dim str = p.GetDisplayText(j.ToString)

            lstVents.Items.Add(str)

        Next

        'sorts the list
        lstVents.Items.Clear()
        For i = 1 To NumberRooms
            For Each p As oVent In oVents
                If p.fromroom = i Then
                    Dim str = p.GetDisplayText(p.id)

                    lstVents.Items.Add(str)
                End If
            Next
        Next

    End Sub
   

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click

        oVents = VentDB.GetVents()
        Dim ovent As oVent
        Dim vc As Integer = oVents.Count + 1
        oVentdistributions = VentDB.GetVentDistributions
      
        NewVentForm = New frmNewVents
        NewVentForm.Text = "New Vent"
        ovent = New oVent(0, 200, False, 10, 100, 0, 0, 1, False, False, False, 0, gcs_DischargeCoeff, False, False, False, False, 0, 0, 0, 0, 500, 1, False, 0, False, 0, 0, 0, False, 72000, 0.00000042, 47, 20, 6, 0.0000083, 1, 0.937, "Vent " & vc.ToString, 1, 2, 0, 1, 1, 0, 0, 0, 0, 0, False, False)
        ovent.id = vc
        NewVentForm.Tag = vc
        Dim fromroom As Integer = ovent.fromroom
        Dim toroom As Integer = ovent.toroom

        If ovent IsNot Nothing Then

            oVents.Add(ovent)

            number_vents = NumberVents(fromroom, toroom) + 1
            NumberVents(fromroom, toroom) = number_vents
            NumberVents(toroom, fromroom) = number_vents
            Resize_Vents()

            NewVentForm.lstRoom1.Items.Clear()
            NewVentForm.lstRoom2.Items.Clear()
            For i = 1 To NumberRooms
                NewVentForm.lstRoom1.Items.Add(i)
            Next

            NewVentForm.lstRoom1.SelectedItem = fromroom

            NewVentForm.lstRoom2.SelectedItem = toroom
            NewVentForm.lstRoomFace.SelectedIndex = ovent.face
            NewVentForm.txtOffset.Text = ovent.offset
            NewVentForm.txtVentSillHeight.Text = ovent.sillheight
            NewVentForm.txtVentOpenTime.Text = ovent.opentime
            NewVentForm.txtVentCloseTime.Text = ovent.closetime
            If ovent.autobreakglass = True Then
                NewVentForm.optGlassAutoBreak.Checked = True
            Else
                NewVentForm.optGlassManualOpen.Checked = True
            End If

            NewVentForm.txtBalconyWidth.Enabled = True
            If ovent.spillplumemodel = 0 Then
                NewVentForm.txtBalconyWidth.Enabled = False
                NewVentForm.optStandard.Checked = True
            ElseIf ovent.spillplumemodel = 1 Then
                NewVentForm.opt2Dadhered.Checked = True
            ElseIf ovent.spillplumemodel = 2 Then
                NewVentForm.opt2Dbalcony.Checked = True
            ElseIf ovent.spillplumemodel = 3 Then
                NewVentForm.opt3Dadhered.Checked = True
            ElseIf ovent.spillplumemodel = 4 Then
                NewVentForm.opt3Dchanbalcony.Checked = True
            ElseIf ovent.spillplumemodel = 5 Then
                NewVentForm.opt3Dunchanbalcony.Checked = True
            Else
            End If

            NewVentForm.txtHOreliability.Text = ovent.horeliability
            NewVentForm.txtVentProb.Text = ovent.probventclosed
            NewVentForm.txtVentCD.Text = ovent.cd
            NewVentForm.txtDownstand.Text = ovent.downstand
            NewVentForm.txtBalconyWidth.Text = ovent.spillbalconyprojection

            Dim odistribution As New oDistribution("", "", "None", 1, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "height"
            NewVentForm.txtVentHeight.Text = odistribution.varvalue
            oVentdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 1, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "width"
            NewVentForm.txtVentWidth.Text = odistribution.varvalue
            oVentdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "prob"
            NewVentForm.txtVentProb.Text = odistribution.varvalue
            oVentdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 1, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "HOreliability"
            NewVentForm.txtHOreliability.Text = odistribution.varvalue
            oVentdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "integrity"
            frmFireResistance.txtFRintegrity.Text = odistribution.varvalue
            oVentdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 100, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "maxopening"
            frmFireResistance.txtMaxOpening.Text = odistribution.varvalue
            oVentdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 10, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "maxopeningtime"
            frmFireResistance.txtMaxOpeningTime.Text = odistribution.varvalue
            oVentdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 200, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "gastemp"
            frmFireResistance.txtFRgastemp.Text = odistribution.varvalue
            oVentdistributions.Add(odistribution)

            VentDB.SaveVents(oVents, oVentdistributions)
            NewVentForm.oVent = ovent
            NewVentForm.oVents = oVents
        End If


        NewVentForm.ShowDialog()
        NewVentForm.BringToFront()

    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click

        For fromroom = 1 To NumberRooms
            For toroom = 1 To NumberRooms + 1
                NumberVents(fromroom, toroom) = 0
            Next
        Next

        For Each oVent In oVents
            Dim i = oVent.fromroom
            Dim j = oVent.toroom
            NumberVents(i, j) = NumberVents(i, j) + 1
            NumberVents(j, i) = NumberVents(i, j)
        Next

        Me.Close()
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Dim i As Integer = lstVents.SelectedIndex

        Dim str As String = lstVents.SelectedItem
        If str = "" Then Exit Sub

        Dim part = str.Split(" ")
        i = CInt(part(0)) - 1

        If lstVents.SelectedIndex <> -1 Then
            oVents = VentDB.GetVents
            Dim ovent As oVent = oVents(i)
            oVentdistributions = VentDB.GetVentDistributions
            VentDB.SaveVents(oVents, oVentdistributions)

            Dim NewventForm As New frmNewVents
            Dim toroom As Integer = ovent.toroom
            Dim fromroom As Integer = ovent.fromroom

            NewventForm.Text = "Edit Vent"
            NewventForm.oVent = ovent

            If NewventForm.oVent IsNot Nothing Then

                NewventForm.Tag = ovent.id
                NewventForm.txtVentDescription.Text = ovent.description

                NewventForm.lstRoom1.Items.Clear()
                NewventForm.lstRoom2.Items.Clear()

                For i = 1 To NumberRooms
                    NewventForm.lstRoom1.Items.Add(i)
                    'NewventForm.lstRoom2.Items.Add(i + 1)
                Next

                NewventForm.lstRoom1.SelectedItem = fromroom

                'For i = fromroom To NumberRooms
                '    'NewventForm.lstRoom1.Items.Add(i)
                '    NewventForm.lstRoom2.Items.Add(i + 1)
                'Next
                If toroom = NumberRooms + 1 Then
                    NewventForm.lstRoom2.SelectedItem = "outside"
                Else
                    NewventForm.lstRoom2.SelectedItem = toroom
                End If

                NewventForm.lstRoomFace.SelectedIndex = ovent.face
                NewventForm.txtOffset.Text = ovent.offset
                NewventForm.txtVentCloseTime.Text = ovent.closetime
                NewventForm.txtVentOpenTime.Text = ovent.opentime
                NewventForm.txtVentSillHeight.Text = ovent.sillheight

                If ovent.autobreakglass = True Then
                    NewventForm.optGlassAutoBreak.Checked = True
                Else
                    NewventForm.optGlassManualOpen.Checked = True
                End If


                NewventForm.txtDownstand.Enabled = True
                NewventForm.txtBalconyWidth.Enabled = True

                If ovent.spillplumemodel = 0 Then
                    NewventForm.optStandard.Checked = True
                    NewventForm.txtBalconyWidth.Enabled = False
                ElseIf ovent.spillplumemodel = 1 Then
                    NewventForm.opt2Dadhered.Checked = True
                ElseIf ovent.spillplumemodel = 2 Then
                    NewventForm.opt2Dbalcony.Checked = True
                ElseIf ovent.spillplumemodel = 3 Then
                    NewventForm.opt3Dadhered.Checked = True
                ElseIf ovent.spillplumemodel = 4 Then
                    NewventForm.opt3Dchanbalcony.Checked = True
                ElseIf ovent.spillplumemodel = 5 Then
                    NewventForm.opt3Dunchanbalcony.Checked = True
                Else
                End If

                NewventForm.txtVentCD.Text = ovent.cd
                NewventForm.txtDownstand.Text = ovent.downstand
                NewventForm.txtBalconyWidth.Text = ovent.spillbalconyprojection

                For Each oDistribution In oVentdistributions
                    If oDistribution.id = ovent.id Then
                        If oDistribution.varname = "height" Then
                            NewventForm.txtVentHeight.Text = oDistribution.varvalue
                            If oDistribution.distribution <> "None" Then
                                NewventForm.txtVentHeight.BackColor = distbackcolor
                            Else
                                NewventForm.txtVentHeight.BackColor = distnobackcolor
                            End If
                        End If
                        If oDistribution.varname = "width" Then
                            NewventForm.txtVentWidth.Text = oDistribution.varvalue
                            If oDistribution.distribution <> "None" Then
                                NewventForm.txtVentWidth.BackColor = distbackcolor
                            Else
                                NewventForm.txtVentWidth.BackColor = distnobackcolor
                            End If
                        End If

                        If oDistribution.varname = "prob" Then
                            NewventForm.txtVentProb.Text = oDistribution.varvalue
                            If oDistribution.distribution <> "None" Then
                                NewventForm.txtVentProb.BackColor = distbackcolor
                            Else
                                NewventForm.txtVentProb.BackColor = distnobackcolor
                            End If
                        End If

                        If oDistribution.varname = "HOreliability" Then
                            NewventForm.txtHOreliability.Text = oDistribution.varvalue
                            If oDistribution.distribution <> "None" Then
                                NewventForm.txtHOreliability.BackColor = distbackcolor
                            Else
                                NewventForm.txtHOreliability.BackColor = distnobackcolor
                            End If
                        End If

                    End If
                Next



                NewventForm.ShowDialog()
                NewventForm.BringToFront()


            End If
        End If

    End Sub

    Private Sub cmdRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemove.Click
        Dim i As Integer = lstVents.SelectedIndex
        Dim j As Integer = 0

        Dim str As String = lstVents.SelectedItem
        Dim part = str.Split(" ")
        i = CInt(part(0)) - 1

        If lstVents.SelectedIndex <> -1 Then
            Dim ovent As oVent = oVents(i)
            Dim oventdistributions As List(Of oDistribution)

            Dim message As String = "Are you sure you want to delete " _
                & "Vent " & CStr(i + 1) & "?"
            Dim button As DialogResult = MessageBox.Show(message, _
                "Confirm Delete", MessageBoxButtons.YesNo)
            If button = Windows.Forms.DialogResult.Yes Then
                oventdistributions = VentDB.GetVentDistributions()

here:
                For Each oDistribution In oventdistributions
                    If oDistribution.id = ovent.id Then
                        oventdistributions.Remove(oDistribution)
                        GoTo here
                    End If
                Next

                oVents.Remove(ovent)

                'sort and reindex
                Dim count As Integer = 1
                For Each ovent In oVents

                    For Each oDistribution In oventdistributions
                        If oDistribution.id = ovent.id Then
                            oDistribution.id = count
                        End If
                    Next

                    ovent.id = count

                    count = count + 1
                Next

                VentDB.SaveVents(oVents, oventdistributions)
                Me.FillVentList()
            End If
        End If
    End Sub

    Private Sub cmdCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopy.Click
        Dim i As Integer = lstVents.SelectedIndex
        Dim str As String = lstVents.SelectedItem

        If IsNothing(str) = True Then Exit Sub

        Dim part = str.Split(" ")
        i = CInt(part(0)) - 1

        If lstVents.SelectedIndex <> -1 Then
            oVents = VentDB.GetVents
            oVentdistributions = VentDB.GetVentDistributions

            Dim oventx As oVent = oVents(i)
            Dim copyvent As oVent = CType(oventx.Clone, oVent)

            If copyvent IsNot Nothing Then
                Dim vc = oVents.Count + 1

                copyvent.id = vc
                copyvent.description = "vent " & vc.ToString
                oVents.Add(copyvent)

                Dim odistribution As New oDistribution("", "", "None", 1, 0, 0, 0, 0, 0, 0, 0)
                odistribution.id = vc
                odistribution.varname = "height"
                odistribution.varvalue = copyvent.height
                oVentdistributions.Add(odistribution)

                odistribution = New oDistribution("", "", "None", 1, 0, 0, 0, 0, 0, 0, 0)
                odistribution.id = vc
                odistribution.varname = "width"
                odistribution.varvalue = copyvent.width
                oVentdistributions.Add(odistribution)

                odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
                odistribution.id = vc
                odistribution.varname = "prob"
                odistribution.varvalue = copyvent.probventclosed
                oVentdistributions.Add(odistribution)

                odistribution = New oDistribution("", "", "None", 1, 0, 0, 0, 0, 0, 0, 0)
                odistribution.id = vc
                odistribution.varname = "HOreliability"
                odistribution.varvalue = copyvent.horeliability
                oVentdistributions.Add(odistribution)

                odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
                odistribution.id = vc
                odistribution.varname = "integrity"
                odistribution.varvalue = copyvent.integrity
                oVentdistributions.Add(odistribution)

                odistribution = New oDistribution("", "", "None", 100, 0, 0, 0, 0, 0, 0, 0)
                odistribution.id = vc
                odistribution.varname = "maxopening"
                odistribution.varvalue = copyvent.maxopening
                oVentdistributions.Add(odistribution)

                odistribution = New oDistribution("", "", "None", 10, 0, 0, 0, 0, 0, 0, 0)
                odistribution.id = vc
                odistribution.varname = "maxopeningtime"
                odistribution.varvalue = copyvent.maxopeningtime
                oVentdistributions.Add(odistribution)

                odistribution = New oDistribution("", "", "None", 200, 0, 0, 0, 0, 0, 0, 0)
                odistribution.id = vc
                odistribution.varname = "gastemp"
                odistribution.varvalue = copyvent.gastemp
                oVentdistributions.Add(odistribution)

                VentDB.SaveVents(oVents, oVentdistributions)

                Call frmInputs.wallventarrays(oVents, oVentdistributions, 0)

                Me.FillVentList()
            End If

        End If
    End Sub

End Class