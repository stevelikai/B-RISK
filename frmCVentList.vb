Imports System.Collections.Generic
Imports System.Net
Imports System.IO
Imports System.Xml

Public Class frmCVentList
    Public NewCVentForm As frmNewCVents
    Dim oCVents As List(Of oCVent)
    Dim oCVentdistributions As List(Of oDistribution)

    Private Sub frmCVentList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim sep As String = vbTab
        Dim text As String
        oCVents = VentDB.GetCVents()

        Dim lab1 As String = "Description"
        Dim lab2 As String = "U Room"
        Dim lab3 As String = "L Room"
        Dim lab4 As String = "ID"
        Dim lab5 As String = "Area (m2)"

        text = LeftJustified(lab4, 3) & _
         RightJustified(lab1, 18) & _
         RightJustified(lab2, 9) & _
          RightJustified(lab3, 9) & _
        RightJustified(lab5, 11)
        ' text = "ID" & sep & "Description" & sep & "From Room" & sep & "To Room"
        ListBox1.Items.Add(text)
        Me.FillVentList()

        Me.CenterToScreen()
        Me.AddOwnedForm(frmNewCVents)

    End Sub

    Public Sub FillVentList()

        lstVents.Items.Clear()

        oCVents = VentDB.GetCVents()
        Dim j As Integer = 0

        For Each p As oCVent In oCVents
            j = j + 1

            Dim str = p.GetDisplayText(j.ToString)

            lstVents.Items.Add(str)

        Next

        'sorts the list
        lstVents.Items.Clear()
        For i = 1 To NumberRooms + 1
            For Each p As oCVent In oCVents
                If p.upperroom = i Then
                    Dim str = p.GetDisplayText(p.id)

                    lstVents.Items.Add(str)
                End If
            Next
        Next

    End Sub


    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click

        Try

            oCVents = VentDB.GetCVents()
            Dim ocvent As oCVent
            Dim vc As Integer = oCVents.Count + 1
            oCVentdistributions = VentDB.GetCVentDistributions

            NewCVentForm = New frmNewCVents
            NewCVentForm.Text = "New Vent"
            ocvent = New oCVent(200, 10, 100, 0, "ceiling vent", 2, 1, 1, 0, 0, False, False, 0, 0.6)
            ocvent.id = vc
            NewCVentForm.Tag = vc
            Dim upperroom As Integer = ocvent.upperroom
            Dim lowerroom As Integer = ocvent.lowerroom

            If ocvent IsNot Nothing Then

                oCVents.Add(ocvent)

                number_cvents = NumberCVents(upperroom, lowerroom) + 1
                NumberCVents(upperroom, lowerroom) = number_cvents
                Resize_CVents()

                NewCVentForm.lstCVentIDUpper.Items.Clear()
                NewCVentForm.lstCVentIDlower.Items.Clear()

                For i = 1 To NumberRooms
                    NewCVentForm.lstCVentIDUpper.Items.Add(i)
                    NewCVentForm.lstCVentIDlower.Items.Add(i)
                Next
                NewCVentForm.lstCVentIDUpper.Items.Add("outside")
                NewCVentForm.lstCVentIDlower.Items.Add("outside")

                NewCVentForm.lstCVentIDUpper.SelectedItem = "outside"
                NewCVentForm.lstCVentIDlower.SelectedItem = 1

                'NewCVentForm.txtCVentArea.Text = ocvent.area
                NewCVentForm.txtCVentOpenTime.Text = ocvent.opentime
                NewCVentForm.txtCVentCloseTime.Text = ocvent.closetime
                NewCVentForm.optCventManual.Checked = True
                NewCVentForm.txtDischargeCoeff.Text = ocvent.dischargecoeff

                Dim odistribution As New oDistribution("", "", "None", 1, 0, 0, 0, 0, 0, 0, 0)
                odistribution.id = ocvent.id
                odistribution.varname = "area"
                NewCVentForm.txtCVentArea.Text = odistribution.varvalue
                oCVentdistributions.Add(odistribution)

                odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
                odistribution.id = ocvent.id
                odistribution.varname = "integrity"
                frmFireResistance.txtFRintegrity.Text = odistribution.varvalue
                oCVentdistributions.Add(odistribution)

                odistribution = New oDistribution("", "", "None", 100, 0, 0, 0, 0, 0, 0, 0)
                odistribution.id = ocvent.id
                odistribution.varname = "maxopening"
                frmFireResistance.txtMaxOpening.Text = odistribution.varvalue
                oCVentdistributions.Add(odistribution)

                odistribution = New oDistribution("", "", "None", 10, 0, 0, 0, 0, 0, 0, 0)
                odistribution.id = ocvent.id
                odistribution.varname = "maxopeningtime"
                frmFireResistance.txtMaxOpeningTime.Text = odistribution.varvalue
                oCVentdistributions.Add(odistribution)

                odistribution = New oDistribution("", "", "None", 200, 0, 0, 0, 0, 0, 0, 0)
                odistribution.id = ocvent.id
                odistribution.varname = "gastemp"
                frmFireResistance.txtFRgastemp.Text = odistribution.varvalue
                oCVentdistributions.Add(odistribution)

                VentDB.SaveCVents(oCVents, oCVentdistributions)
                NewCVentForm.ocVent = ocvent
                NewCVentForm.ocVents = oCVents

            End If

            NewCVentForm.ShowDialog()
            NewCVentForm.BringToFront()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in frmCVentList.vb cmdAdd_click - line " & Err.Erl)
        End Try

    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click

        Me.Close()

    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Try
            If lstVents.SelectedIndex < 0 Then Exit Sub

            Dim i As Integer = lstVents.SelectedIndex
            Dim str As String = lstVents.SelectedItem
            Dim part = str.Split(" ")
            i = CInt(part(0)) - 1

            If lstVents.SelectedIndex <> -1 Then
                oCVents = VentDB.GetCVents
                Dim ocvent As oCVent = oCVents(i)
                oCVentdistributions = VentDB.GetCVentDistributions
                VentDB.SaveCVents(oCVents, oCVentdistributions)

                Dim NewcventForm As New frmNewCVents
                Dim upperroom As Integer = ocvent.upperroom
                Dim lowerroom As Integer = ocvent.lowerroom

                NewcventForm.Text = "Edit Vent"
                NewcventForm.ocVent = ocvent

                If NewcventForm.ocVent IsNot Nothing Then

                    NewcventForm.Tag = ocvent.id
                    NewcventForm.txtVentDescription.Text = ocvent.description

                    NewcventForm.lstCVentIDUpper.Items.Clear()
                    NewcventForm.lstCVentIDlower.Items.Clear()

                    For i = 1 To NumberRooms
                        NewcventForm.lstCVentIDUpper.Items.Add(i)
                        NewcventForm.lstCVentIDlower.Items.Add(i)
                    Next
                    NewcventForm.lstCVentIDUpper.Items.Add("outside")
                    NewcventForm.lstCVentIDlower.Items.Add("outside")

                    If upperroom > NumberRooms Then
                        NewcventForm.lstCVentIDUpper.SelectedItem = "outside"
                    Else
                        NewcventForm.lstCVentIDUpper.SelectedItem = upperroom
                    End If
                    If lowerroom > NumberRooms Then
                        NewcventForm.lstCVentIDlower.SelectedItem = "outside"
                    Else
                        NewcventForm.lstCVentIDlower.SelectedItem = lowerroom
                    End If

                    NewcventForm.txtCVentCloseTime.Text = ocvent.closetime
                    NewcventForm.txtCVentOpenTime.Text = ocvent.opentime
                    NewcventForm.txtDischargeCoeff.Text = ocvent.dischargecoeff

                    For Each oDistribution In oCVentdistributions
                        If oDistribution.id = ocvent.id Then
                            If oDistribution.varname = "area" Then
                                NewcventForm.txtCVentArea.Text = oDistribution.varvalue
                                If oDistribution.distribution <> "None" Then
                                    NewcventForm.txtCVentArea.BackColor = distbackcolor
                                Else
                                    NewcventForm.txtCVentArea.BackColor = distnobackcolor
                                End If
                            End If
                        End If

                        If oDistribution.id = ocvent.id Then
                            If oDistribution.varname = "integrity" Then
                                frmFireResistance.txtFRintegrity.Text = oDistribution.varvalue
                                If oDistribution.distribution <> "None" Then
                                    frmFireResistance.txtFRintegrity.BackColor = distbackcolor
                                Else
                                    frmFireResistance.txtFRintegrity.BackColor = distnobackcolor
                                End If
                            End If
                        End If

                        If oDistribution.id = ocvent.id Then
                            If oDistribution.varname = "maxopening" Then
                                frmFireResistance.txtMaxOpening.Text = oDistribution.varvalue
                                If oDistribution.distribution <> "None" Then
                                    frmFireResistance.txtMaxOpening.BackColor = distbackcolor
                                Else
                                    frmFireResistance.txtMaxOpening.BackColor = distnobackcolor
                                End If
                            End If
                        End If

                        If oDistribution.id = ocvent.id Then
                            If oDistribution.varname = "maxopeningtime" Then
                                frmFireResistance.txtMaxOpeningTime.Text = oDistribution.varvalue
                                If oDistribution.distribution <> "None" Then
                                    frmFireResistance.txtMaxOpeningTime.BackColor = distbackcolor
                                Else
                                    frmFireResistance.txtMaxOpeningTime.BackColor = distnobackcolor
                                End If
                            End If
                        End If

                        If oDistribution.id = ocvent.id Then
                            If oDistribution.varname = "gastemp" Then
                                frmFireResistance.txtFRgastemp.Text = oDistribution.varvalue
                                If oDistribution.distribution <> "None" Then
                                    frmFireResistance.txtFRgastemp.BackColor = distbackcolor
                                Else
                                    frmFireResistance.txtFRgastemp.BackColor = distnobackcolor
                                End If
                            End If
                        End If

                    Next

                    NewcventForm.ShowDialog()
                    NewcventForm.BringToFront()

                End If
            End If

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in frmCVentList.vb cmdEdit_click - line " & Err.Erl)
        End Try

    End Sub

    Private Sub cmdRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemove.Click
        Try

            Dim i As Integer = lstVents.SelectedIndex
            Dim j As Integer = 0

            Dim str As String = lstVents.SelectedItem
            Dim part = str.Split(" ")
            i = CInt(part(0)) - 1

            If lstVents.SelectedIndex <> -1 Then
                Dim ocvent As oCVent = oCVents(i)
                Dim ocventdistributions As List(Of oDistribution)

                Dim message As String = "Are you sure you want to delete " _
                    & "Vent " & CStr(i + 1) & "?"
                Dim button As DialogResult = MessageBox.Show(message, _
                    "Confirm Delete", MessageBoxButtons.YesNo)
                If button = Windows.Forms.DialogResult.Yes Then
                    ocventdistributions = VentDB.GetCVentDistributions()

here:
                    For Each oDistribution In ocventdistributions
                        If oDistribution.id = ocvent.id Then
                            ocventdistributions.Remove(oDistribution)
                            GoTo here
                        End If
                    Next

                    oCVents.Remove(ocvent)

                    'sort and reindex
                    Dim count As Integer = 1
                    For Each ocvent In oCVents

                        For Each oDistribution In ocventdistributions
                            If oDistribution.id = ocvent.id Then
                                oDistribution.id = count
                            End If
                        Next

                        ocvent.id = count

                        count = count + 1
                    Next

                    VentDB.SaveCVents(oCVents, ocventdistributions)
                    Me.FillVentList()
                End If
            End If


        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in frmCVentList.vb cmdRemove_click - line " & Err.Erl)
        End Try

    End Sub

    Private Sub cmdCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopy.Click

        Try

            Dim i As Integer = lstVents.SelectedIndex
            If i = -1 Then Exit Sub

            Dim str As String = lstVents.SelectedItem
            Dim part = str.Split(" ")
            i = CInt(part(0)) - 1

            If lstVents.SelectedIndex <> -1 Then
                oCVents = VentDB.GetCVents
                oCVentdistributions = VentDB.GetCVentDistributions

                Dim oventx As oCVent = oCVents(i)
                Dim copyvent As oCVent = CType(oventx.Clone, oCVent)

                If copyvent IsNot Nothing Then
                    Dim vc = oCVents.Count + 1

                    copyvent.id = vc
                    copyvent.description = "vent " & vc.ToString
                    oCVents.Add(copyvent)

                    Dim odistribution As New oDistribution("", "", "None", 1, 0, 0, 0, 0, 0, 0, 0)
                    odistribution.id = vc
                    odistribution.varname = "area"
                    odistribution.varvalue = copyvent.area
                    oCVentdistributions.Add(odistribution)

                    odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
                    odistribution.id = vc
                    odistribution.varname = "integrity"
                    odistribution.varvalue = copyvent.integrity
                    oCVentdistributions.Add(odistribution)

                    odistribution = New oDistribution("", "", "None", 100, 0, 0, 0, 0, 0, 0, 0)
                    odistribution.id = vc
                    odistribution.varname = "maxopening"
                    odistribution.varvalue = copyvent.maxopening
                    oCVentdistributions.Add(odistribution)

                    odistribution = New oDistribution("", "", "None", 10, 0, 0, 0, 0, 0, 0, 0)
                    odistribution.id = vc
                    odistribution.varname = "maxopeningtime"
                    odistribution.varvalue = copyvent.maxopeningtime
                    oCVentdistributions.Add(odistribution)

                    odistribution = New oDistribution("", "", "None", 200, 0, 0, 0, 0, 0, 0, 0)
                    odistribution.id = vc
                    odistribution.varname = "gastemp"
                    odistribution.varvalue = copyvent.gastemp
                    oCVentdistributions.Add(odistribution)

                    VentDB.SaveCVents(oCVents, oCVentdistributions)

                    Call frmInputs.ceilingventarrays(oCVents, oCVentdistributions, 0)

                    Me.FillVentList()
                End If

            End If

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in frmCVentList.vb cmdCopy_click - line " & Err.Erl)
        End Try

    End Sub
    Public Sub deletecvent(ByVal i As Integer, ByVal ocvents As Object)

        Dim j As Integer = 0
        If i <> -1 Then
            Dim ocvent As oCVent = ocvents(i)
            Dim ocventdistributions As List(Of oDistribution)

            ocventdistributions = VentDB.GetCVentDistributions()

here:
            For Each oDistribution In ocventdistributions
                If oDistribution.id = ocvent.id Then
                    ocventdistributions.Remove(oDistribution)
                    GoTo here
                End If
            Next

            ocvents.Remove(ocvent)

            'sort and reindex
            Dim count As Integer = 1
            For Each ocvent In ocvents

                For Each oDistribution In ocventdistributions
                    If oDistribution.id = ocvent.id Then
                        oDistribution.id = count
                    End If
                Next

                ocvent.id = count

                count = count + 1
            Next

            VentDB.SaveCVents(ocvents, ocventdistributions)
            Me.FillVentList()
        End If

    End Sub
End Class