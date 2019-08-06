Imports System.Collections.Generic
Imports System.Net
Imports System.IO
Imports System.Xml
Imports System.Xml.Linq
Imports VB = Microsoft.VisualBasic

Public Class frmItemList
    Inherits System.Windows.Forms.Form
    Public NewItemForm As frmNewItem
    Dim oItems As List(Of oItem)
    Dim oitemdistributions As List(Of oDistribution)

    Private Sub frmItemList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sep As String = vbTab
        Dim text As String
        oItems = ItemDB.GetItemsv2()
        Me.FillItemList()

        text = "ID" & sep & "Description"
        ListBox1.Items.Add(text)

        If usepowerlawdesignfire = True Then chkAlphaTFire.Checked = True

        If DevKey = True Then
            GroupBox3.Visible = True
            If FuelResponseEffects = True Then
                optFREon.Checked = True
            Else
                optFREoff.Checked = True
            End If
        Else
                GroupBox3.Visible = False
        End If


        If VM2 = True Then
            If lstItems.Items.Count > 0 Then

            Else
                oItems = ItemDB.GetItemsv2()
                Dim oitem As oItem
                oitemdistributions = ItemDB.GetItemDistributions

                NewItemForm = New frmNewItem

                NewItemForm.Text = "New Item"
                oitem = New oItem("0,0", 98.4, 0.338, 2, 684, 60, 0.101, 1, 250, "new item", "OBJ", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 427, 481, 0, 0, 0, 0, 9.5, 22, 0, 0.3, 0.3, 0.3, 0, 1.5, 0.07, "generic", "description", 20, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, 1, RoomLength(fireroom) / 2, RoomWidth(fireroom) / 2, 0.35, 0, 0, 3, 0)

                oitem.id = oItems.Count + 1
                NewItemForm.Tag = oItems.Count + 1
                If oitem IsNot Nothing Then
                    oItems.Add(oitem)
                    NewItemForm.txtItemDescription.Text = oitem.description
                    NewItemForm.txtDetailedDescription.Text = oitem.detaileddescription
                    NewItemForm.txtUserLabel.Text = oitem.userlabel
                    NewItemForm.ComboType.Text = oitem.type
                    NewItemForm.txtItemElevation.Text = oitem.elevation
                    NewItemForm.txtItemHeight.Text = oitem.height
                    NewItemForm.txtItemLength.Text = oitem.length
                    NewItemForm.txtItemWidth.Text = oitem.width
                    NewItemForm.txtCritFluxauto.Text = oitem.critfluxauto
                    NewItemForm.txtFTPlimitpilot.Text = oitem.ftplimitpilot
                    NewItemForm.txtFTPindexpilot.Text = oitem.ftpindexpilot
                    NewItemForm.txtCritFlux.Text = oitem.critflux
                    NewItemForm.txtFTPlimitauto.Text = oitem.ftplimitauto
                    NewItemForm.txtFTPindexauto.Text = oitem.ftpindexauto
                    NewItemForm.txtMass.Text = oitem.mass
                    NewItemForm.txtItemProb.Text = oitem.prob
                    NewItemForm.Time___Heat_ReleaseTextBox.Text = oitem.hrr
                    NewItemForm.txtFreeBurnMLR.Text = oitem.mlrfreeburn
                    NewItemForm.txtIgnitionTime.Text = oitem.ignitiontime
                    NewItemForm.txtMaxNumItem.Text = oitem.maxnumitem
                    NewItemForm.txtXleft.Text = oitem.xleft
                    NewItemForm.txtYbottom.Text = oitem.ybottom
                    NewItemForm.txtConstA.Text = oitem.constantA
                    NewItemForm.txtConstB.Text = oitem.constantB
                    If oitem.windeffect > 1 Then
                        NewItemForm.optwind.Checked = True
                    Else
                        NewItemForm.optnowind.Checked = True
                    End If
                    If oitem.pyrolysisoption = 0 Then
                        NewItemForm.optFreeBurn.Checked = True
                    ElseIf oitem.pyrolysisoption = 1 Then
                        NewItemForm.optPoolFire.Checked = True
                    ElseIf oitem.pyrolysisoption = 2 Then
                        NewItemForm.optWoodCrib.Checked = True
                    End If

                    Dim odistribution As New oDistribution("", "", "None", oitem.hoc, 0, 0, 0, 0, 0, 0, 0)
                    odistribution.id = oitem.id
                    odistribution.varname = "heat of combustion"
                    NewItemForm.txtValue8.Text = odistribution.varvalue
                    oitemdistributions.Add(odistribution)

                    odistribution = New oDistribution("", "", "None", oitem.soot, 0, 0, 0, 0, 0, 0, 0)
                    odistribution.id = oitem.id
                    odistribution.varname = "soot yield"
                    NewItemForm.txtSootValue.Text = odistribution.varvalue
                    oitemdistributions.Add(odistribution)

                    odistribution = New oDistribution("", "", "None", oitem.co2, 0, 0, 0, 0, 0, 0, 0)
                    odistribution.id = oitem.id
                    odistribution.varname = "co2 yield"
                    NewItemForm.txtCO2Value.Text = odistribution.varvalue
                    oitemdistributions.Add(odistribution)

                    odistribution = New oDistribution("", "", "None", oitem.LHoG, 0, 0, 0, 0, 0, 0, 0)
                    odistribution.id = oitem.id
                    odistribution.varname = "Latent Heat of Gasification"
                    NewItemForm.txtLHOG.Text = odistribution.varvalue
                    oitemdistributions.Add(odistribution)

                    odistribution = New oDistribution("", "", "None", oitem.radlossfrac, 0, 0, 0, 0, 0, 0, 0)
                    odistribution.id = oitem.id
                    odistribution.varname = "Radiant Loss Fraction"
                    NewItemForm.txtRadLossFracItem.Text = odistribution.varvalue
                    oitemdistributions.Add(odistribution)

                    odistribution = New oDistribution("", "", "None", oitem.HRRUA, 0, 0, 0, 0, 0, 0, 0)
                    odistribution.id = oitem.id
                    odistribution.varname = "HRRUA"
                    NewItemForm.txtHRRUA.Text = odistribution.varvalue
                    oitemdistributions.Add(odistribution)

                    ItemDB.SaveItemsv2(oItems, oitemdistributions)

                End If

                FillItemList()
            End If

            lstItems.SelectedIndex = 0
            chkAlphaTFire.Enabled = False
            cmdRoomPopulate.Enabled = False
            cmdImport.Enabled = False
            cmdCopy.Enabled = False
            cmdAdd.Enabled = False
            cmdRemove.Enabled = False
            cmdRandomIgnite.Enabled = False
            cmdIgniteFirst.Enabled = False

        Else

            chkAlphaTFire.Enabled = True
            cmdRoomPopulate.Enabled = True
            cmdImport.Enabled = True
            cmdCopy.Enabled = True
            cmdAdd.Enabled = True
            cmdRemove.Enabled = True
            cmdRandomIgnite.Enabled = True
            cmdIgniteFirst.Enabled = True

        End If

    End Sub

    Public Sub FillItemList()

        Try

            lstItems.Items.Clear()

            oItems = ItemDB.GetItemsv2()
            If oItems.Count = 0 Then
                frmInputs.StartToolStripLabel1.Visible = False
                frmInputs.StopToolStripButton1.Visible = False
                Exit Sub
            End If

            frmInputs.StartToolStripLabel1.Visible = True
            frmInputs.StopToolStripButton1.Visible = False

            Dim j As Integer = 0

            For Each p As oItem In oItems
                j = j + 1

                Dim str = p.GetDisplayText(vbTab)
                If str.Length > 40 Then str = str.Remove(40, str.Length - 40) 'truncate to length 40 char
                'If j = firstitem Then
                If p.id = firstitem Then
                    If usepowerlawdesignfire = False Then
                        'lstItems.Items.Add(j.ToString & vbTab & str & vbTab & "IGNITE")
                        lstItems.Items.Add(p.id.ToString & vbTab & str & vbTab & "IGNITE")
                    End If
                Else
                    If usepowerlawdesignfire = False Then
                        'lstItems.Items.Add(j.ToString & vbTab & str)
                        lstItems.Items.Add(p.id.ToString & vbTab & str)
                    End If
                End If
            Next
            If usepowerlawdesignfire = True Then
                Dim str = oItems(0).GetDisplayText(vbTab)
                If str.Length > 40 Then str = str.Remove(40, str.Length - 40) 'truncate to length 40 char
                lstItems.Items.Add(1.ToString & vbTab & str)
            End If

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in frmItemList.vb FillItemList")
        End Try
    End Sub
    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click

        oItems = ItemDB.GetItemsv2()
        Dim oitem As oItem
        oitemdistributions = ItemDB.GetItemDistributions

        NewItemForm = New frmNewItem
       
        NewItemForm.Text = "New Item"
        oitem = New oItem("0,0", 98.4, 0.338, 2, 684, 60, 0.101, 1, 250, "new item", "OBJ", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 427, 481, 0, 0, 0, 0, 9.5, 22, 0, 0.3, 0.3, 0.3, 0, 1.5, 0.07, "generic", "description", 20, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, 1, RoomLength(fireroom) / 2, RoomWidth(fireroom) / 2, 0.35, 0, 0, 3, 0)

        oitem.id = oItems.Count + 1
        NewItemForm.Tag = oItems.Count + 1
        If oitem IsNot Nothing Then
            oItems.Add(oitem)
            NewItemForm.txtItemDescription.Text = oitem.description
            NewItemForm.txtDetailedDescription.Text = oitem.detaileddescription
            NewItemForm.txtUserLabel.Text = oitem.userlabel
            NewItemForm.ComboType.Text = oitem.type
            NewItemForm.txtItemElevation.Text = oitem.elevation
            NewItemForm.txtItemHeight.Text = oitem.height
            NewItemForm.txtItemLength.Text = oitem.length
            NewItemForm.txtItemWidth.Text = oitem.width
            NewItemForm.txtCritFluxauto.Text = oitem.critfluxauto
            NewItemForm.txtFTPlimitpilot.Text = oitem.ftplimitpilot
            NewItemForm.txtFTPindexpilot.Text = oitem.ftpindexpilot
            NewItemForm.txtCritFlux.Text = oitem.critflux
            NewItemForm.txtFTPlimitauto.Text = oitem.ftplimitauto
            NewItemForm.txtFTPindexauto.Text = oitem.ftpindexauto
            NewItemForm.txtMass.Text = oitem.mass
            NewItemForm.txtItemProb.Text = oitem.prob
            NewItemForm.Time___Heat_ReleaseTextBox.Text = oitem.hrr
            NewItemForm.txtFreeBurnMLR.Text = oitem.mlrfreeburn
            NewItemForm.txtIgnitionTime.Text = oitem.ignitiontime
            NewItemForm.txtMaxNumItem.Text = oitem.maxnumitem
            NewItemForm.txtXleft.Text = oitem.xleft
            NewItemForm.txtYbottom.Text = oitem.ybottom
            'NewItemForm.txtRadLossFracItem.Text = oitem.radlossfrac
            NewItemForm.txtConstA.Text = oitem.constantA
            NewItemForm.txtConstB.Text = oitem.constantB
            'NewItemForm.txtLHOG.Text = oitem.LHoG
            NewItemForm.txtHRRUA.Text = oitem.HRRUA
            If oitem.windeffect > 1 Then
                NewItemForm.optwind.Checked = True
            Else
                NewItemForm.optnowind.Checked = True
            End If
            If oitem.pyrolysisoption = 0 Then
                NewItemForm.optFreeBurn.Checked = True
            ElseIf oitem.pyrolysisoption = 1 Then
                NewItemForm.optPoolFire.Checked = True
            ElseIf oitem.pyrolysisoption = 2 Then
                NewItemForm.optWoodCrib.Checked = True
            End If
            If optFREon.Checked = True Then
                FuelResponseEffects = True
            Else
                FuelResponseEffects = False
            End If
            If sootmode = True Then
                'manual setting
                NewItemForm.txtSootValue.Visible = False
                NewItemForm.Label36.Visible = False
                NewItemForm.Label1.Visible = False
                NewItemForm.cmdDist_soot.Visible = False
            Else
                NewItemForm.txtSootValue.Visible = True
                NewItemForm.Label36.Visible = True
                NewItemForm.Label1.Visible = True
                NewItemForm.cmdDist_soot.Visible = True
            End If

            Dim odistribution As New oDistribution("", "", "None", oitem.hoc, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = oitem.id
            odistribution.varname = "heat of combustion"
            NewItemForm.txtValue8.Text = odistribution.varvalue
            oitemdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", oitem.soot, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = oitem.id
            odistribution.varname = "soot yield"
            NewItemForm.txtSootValue.Text = odistribution.varvalue
            oitemdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", oitem.co2, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = oitem.id
            odistribution.varname = "co2 yield"
            NewItemForm.txtCO2Value.Text = odistribution.varvalue
            oitemdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", oitem.LHoG, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = oitem.id
            odistribution.varname = "Latent Heat of Gasification"
            NewItemForm.txtLHOG.Text = odistribution.varvalue
            oitemdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", oitem.radlossfrac, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = oitem.id
            odistribution.varname = "Radiant Loss Fraction"
            NewItemForm.txtRadLossFracItem.Text = odistribution.varvalue
            oitemdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", oitem.HRRUA, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = oitem.id
            odistribution.varname = "HRRUA"
            NewItemForm.txtHRRUA.Text = odistribution.varvalue
            oitemdistributions.Add(odistribution)

            ItemDB.SaveItemsv2(oItems, oitemdistributions)

            If VM2 = True Then
                NewItemForm.txtItemLength.Enabled = False
                NewItemForm.txtItemWidth.Enabled = False
                NewItemForm.txtItemHeight.Enabled = False
                NewItemForm.txtItemProb.Enabled = False
                NewItemForm.txtLHOG.Enabled = False
                NewItemForm.txtMaxNumItem.Enabled = False
            Else
                NewItemForm.txtItemLength.Enabled = True
                NewItemForm.txtItemWidth.Enabled = True
                NewItemForm.txtItemHeight.Enabled = True
                NewItemForm.txtItemProb.Enabled = True
                NewItemForm.txtLHOG.Enabled = True
                NewItemForm.txtMaxNumItem.Enabled = True
            End If

        End If

        FillItemList()

        Me.AddOwnedForm(frmNewItem)
        Me.AddOwnedForm(NewItemForm)

        NewItemForm.ShowDialog()

    End Sub

    Private Sub cmdRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemove.Click
        Dim i As Integer = lstItems.SelectedIndex
        Dim j As Integer = 0
        If i <> -1 Then
            Dim oitem As oItem = oItems(i)
            Dim oitemdistributions As List(Of oDistribution)

            Dim message As String = "Are you sure you want to delete " _
                & "Item " & CStr(i + 1) & "?"
            Dim button As DialogResult = MessageBox.Show(message, _
                "Confirm Delete", MessageBoxButtons.YesNo)
            If button = Windows.Forms.DialogResult.Yes Then
                oitemdistributions = ItemDB.GetItemDistributions()

here:
                For Each oDistribution In oitemdistributions
                    If oDistribution.id = oitem.id Then
                        oitemdistributions.Remove(oDistribution)
                        GoTo here
                    End If
                Next

                oItems.Remove(oitem)

                'sort and reindex
                Dim count As Integer = 1
                For Each oitem In oItems

                    For Each oDistribution In oitemdistributions
                        If oDistribution.id = oitem.id Then
                            oDistribution.id = count
                        End If
                    Next

                    oitem.id = count
                    count = count + 1
                Next

                ItemDB.SaveItemsv2(oItems, oitemdistributions)
                Me.FillItemList()

                NumberObjects = oItems.Count
                Resize_Objects()

            End If
        End If
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        If optFREon.Checked = True Then
            FuelResponseEffects = True
        Else
            FuelResponseEffects = False
        End If

        ' frmPopulate.cmdPopulate.PerformClick()
        If usepowerlawdesignfire = True Then
            n_max = 1
        End If

        ReDim Item1X(0 To frmInputs.txtNumberIterations.Text - 1)
        ReDim Item1Y(0 To frmInputs.txtNumberIterations.Text - 1)
        ReDim itime(0 To n_max)
        ReDim ignmode(0 To n_max)

        Call populate_items_manual(n_max, 1, 0, 0)

        Me.Close()
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Dim i As Integer = lstItems.SelectedIndex
        If i <> -1 Then
            oItems = ItemDB.GetItemsv2()
            Dim oitem As oItem = oItems(i)
            oitemdistributions = ItemDB.GetItemDistributions

            Dim NewItemForm As New frmNewItem

            NewItemForm.Text = "Edit Item"
            NewItemForm.oitem = oitem

            If NewItemForm.oitem IsNot Nothing Then

                NewItemForm.Tag = oitem.id
                NewItemForm.txtItemDescription.Text = oitem.description
                NewItemForm.txtDetailedDescription.Text = oitem.detaileddescription
                NewItemForm.txtUserLabel.Text = oitem.userlabel
                NewItemForm.ComboType.Text = oitem.type
                NewItemForm.txtItemElevation.Text = oitem.elevation
                NewItemForm.txtItemHeight.Text = oitem.height
                NewItemForm.txtItemLength.Text = oitem.length
                NewItemForm.txtItemWidth.Text = oitem.width
                NewItemForm.txtCritFluxauto.Text = oitem.critfluxauto
                NewItemForm.txtFTPlimitpilot.Text = oitem.ftplimitpilot
                NewItemForm.txtFTPindexpilot.Text = oitem.ftpindexpilot
                NewItemForm.txtCritFlux.Text = oitem.critflux
                NewItemForm.txtFTPlimitauto.Text = oitem.ftplimitauto
                NewItemForm.txtFTPindexauto.Text = oitem.ftpindexauto
                NewItemForm.txtMass.Text = oitem.mass
                NewItemForm.txtItemProb.Text = oitem.prob
                NewItemForm.Time___Heat_ReleaseTextBox.Text = oitem.hrr
                NewItemForm.txtFreeBurnMLR.Text = oitem.mlrfreeburn
                NewItemForm.txtIgnitionTime.Text = oitem.ignitiontime
                NewItemForm.txtMaxNumItem.Text = oitem.maxnumitem
                NewItemForm.txtXleft.Text = oitem.xleft
                NewItemForm.txtYbottom.Text = oitem.ybottom
                NewItemForm.txtConstA.Text = oitem.constantA
                NewItemForm.txtConstB.Text = oitem.constantB
                NewItemForm.txtHRRUA.Text = oitem.HRRUA
                If oitem.windeffect > 1 Then
                    NewItemForm.optwind.Checked = True
                Else
                    NewItemForm.optnowind.Checked = True
                End If
                If oitem.pyrolysisoption = 0 Then
                    NewItemForm.optFreeBurn.Checked = True
                ElseIf oitem.pyrolysisoption = 1 Then
                    NewItemForm.optPoolFire.Checked = True
                ElseIf oitem.pyrolysisoption = 2 Then
                    NewItemForm.optWoodCrib.Checked = True
                End If
                If optFREon.Checked = True Then
                    FuelResponseEffects = True
                Else
                    FuelResponseEffects = False
                End If

                For Each oDistribution In oitemdistributions
                    If oDistribution.id = oitem.id Then
                        If oDistribution.varname = "heat of combustion" Then
                            NewItemForm.txtValue8.Text = oDistribution.varvalue
                            If oDistribution.distribution <> "None" Then
                                NewItemForm.txtValue8.BackColor = distbackcolor
                            Else
                                NewItemForm.txtValue8.BackColor = distnobackcolor
                            End If
                        End If

                        If oDistribution.varname = "soot yield" Then
                            NewItemForm.txtSootValue.Text = oDistribution.varvalue
                            If oDistribution.distribution <> "None" Then
                                NewItemForm.txtSootValue.BackColor = distbackcolor
                            Else
                                NewItemForm.txtSootValue.BackColor = distnobackcolor
                            End If
                        End If

                        If oDistribution.varname = "co2 yield" Then
                            NewItemForm.txtCO2Value.Text = oDistribution.varvalue
                            If oDistribution.distribution <> "None" Then
                                NewItemForm.txtCO2Value.BackColor = distbackcolor
                            Else
                                NewItemForm.txtCO2Value.BackColor = distnobackcolor
                            End If
                        End If

                        If oDistribution.varname = "Latent Heat of Gasification" Then
                            NewItemForm.txtLHOG.Text = oDistribution.varvalue
                            If oDistribution.distribution <> "None" Then
                                NewItemForm.txtLHOG.BackColor = distbackcolor
                            Else
                                NewItemForm.txtLHOG.BackColor = distnobackcolor
                            End If
                        End If

                        If oDistribution.varname = "Radiant Loss Fraction" Then
                            NewItemForm.txtRadLossFracItem.Text = oDistribution.varvalue
                            If oDistribution.distribution <> "None" Then
                                NewItemForm.txtRadLossFracItem.BackColor = distbackcolor
                            Else
                                NewItemForm.txtRadLossFracItem.BackColor = distnobackcolor
                            End If
                        End If

                        If oDistribution.varname = "HRRUA" Then
                            NewItemForm.txtHRRUA.Text = oDistribution.varvalue
                            If oDistribution.distribution <> "None" Then
                                NewItemForm.txtHRRUA.BackColor = distbackcolor
                            Else
                                NewItemForm.txtHRRUA.BackColor = distnobackcolor
                            End If
                        End If

                    End If
                Next

                If usepowerlawdesignfire = True Then
                    NewItemForm.Time___Heat_ReleaseTextBox.Visible = False
                    NewItemForm.txtFreeBurnMLR.Visible = False
                    NewItemForm.Button3.Visible = False
                    NewItemForm.Label50.Visible = False
                    NewItemForm.Label14.Visible = False
                    NewItemForm.Mass.Visible = False
                    NewItemForm.txtMass.Visible = False
                    NewItemForm.kg.Visible = False
                    NewItemForm.cmdMassCalc.Visible = False
                    NewItemForm.TableLayoutPanel2.Visible = False
                    NewItemForm.TableLayoutPanel3.Visible = False
                Else
                    NewItemForm.Time___Heat_ReleaseTextBox.Visible = True
                    NewItemForm.txtFreeBurnMLR.Visible = True
                    NewItemForm.Button3.Visible = True
                    NewItemForm.Label50.Visible = True
                    NewItemForm.Label14.Visible = True
                    NewItemForm.Mass.Visible = True
                    NewItemForm.txtMass.Visible = True
                    NewItemForm.kg.Visible = True
                    NewItemForm.cmdMassCalc.Visible = True
                    NewItemForm.TableLayoutPanel2.Visible = True
                    NewItemForm.TableLayoutPanel3.Visible = True
                End If

                If VM2 = True Then
                    'NewItemForm.TableLayoutPanel2.Visible = False
                    NewItemForm.txtItemLength.Enabled = False
                    NewItemForm.txtItemWidth.Enabled = False
                    NewItemForm.txtItemHeight.Enabled = False
                    NewItemForm.txtItemProb.Enabled = False
                    NewItemForm.txtLHOG.Enabled = False
                    NewItemForm.txtMaxNumItem.Enabled = False
                Else
                    'NewItemForm.TableLayoutPanel2.Visible = True
                    NewItemForm.txtItemLength.Enabled = True
                    NewItemForm.txtItemWidth.Enabled = True
                    NewItemForm.txtItemHeight.Enabled = True
                    NewItemForm.txtItemProb.Enabled = True
                    NewItemForm.txtLHOG.Enabled = True
                    NewItemForm.txtMaxNumItem.Enabled = True
                End If

                If sootmode = True Then
                    'manual setting
                    NewItemForm.txtSootValue.Visible = False
                    NewItemForm.Label36.Visible = False
                    NewItemForm.Label1.Visible = False
                    NewItemForm.cmdDist_soot.Visible = False
                Else
                    NewItemForm.txtSootValue.Visible = True
                    NewItemForm.Label36.Visible = True
                    NewItemForm.Label1.Visible = True
                    NewItemForm.cmdDist_soot.Visible = True
                End If

                NewItemForm.ShowDialog()

            End If

        End If
    End Sub

    Private Sub lstItems_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstItems.SelectedIndexChanged
        Dim a As Integer
        a = lstItems.SelectedIndex + 1

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopy.Click
        Dim oItems As List(Of oItem)
        Dim oitemdistributions As List(Of oDistribution)
        Dim i As Integer = lstItems.SelectedIndex
        Dim id As Integer
        Dim j As Integer = 0

        If optFREon.Checked = True Then
            FuelResponseEffects = True
        Else
            FuelResponseEffects = False
        End If

        If i <> -1 Then
            oItems = ItemDB.GetItemsv2()
            Dim oitem As oItem = oItems(i)
            id = oitem.id
            oitemdistributions = ItemDB.GetItemDistributions
            Dim odistribution As New oDistribution
            Dim NewItemForm As New frmNewItem

            NewItemForm.Text = "Copy Item"
            NewItemForm.oitem = oitem
            Dim counter As Integer
            If NewItemForm.oitem IsNot Nothing Then

                counter = 0

                For z = 0 To oitemdistributions.Count - 1
                    If oitemdistributions(z).id = id Then

                        Dim x As oDistribution = oitemdistributions(z)
                        Dim y As New oDistribution(x.varname, x.units, x.distribution, x.varvalue, x.mean, x.variance, x.lbound, x.ubound, x.mode, x.alpha, x.beta)
                        y.id = oItems.Count + 1
                        oitemdistributions.Add(y)

                    End If
                Next
                NewItemForm.oitem.id = oItems.Count + 1
                oItems.Add(NewItemForm.oitem)

                ItemDB.SaveItemsv2(oItems, oitemdistributions)
                Me.FillItemList()

                NumberObjects = oItems.Count
                Resize_Objects()
            End If

        End If
    End Sub

    Private Sub cmdIgniteFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdIgniteFirst.Click
        Dim i As Integer = lstItems.SelectedIndex
        If i <> -1 Then
            firstitem = i + 1
            Me.FillItemList()
        End If

    End Sub

    Private Sub cmdRandomIgnite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRandomIgnite.Click

        firstitem = 0
        Me.FillItemList()

    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdImport.Click
        frmFireObjectDB.ToolStripButton2.Visible = False
        frmFireObjectDB.ToolStripLabel1.Visible = True
        frmFireObjectDB.Show()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            'web 
            Dim URL As String
            URL = gcs_UOC_XML

            Dim m_nodelist As XmlNodeList
            'Dim m_node As XmlNode
            Dim oitem As oItem
            Dim oItems As New List(Of oItem)
            Dim odistribution As oDistribution
            Dim oitemdistributions As New List(Of oDistribution)

            'oItems = ItemDB.GetItemsv2()
            'oitemdistributions = ItemDB.GetItemDistributions

            Dim Request As WebRequest = HttpWebRequest.Create(URL)
            Dim Response As WebResponse = Request.GetResponse
            Dim sReader As New StreamReader(Response.GetResponseStream())
            Dim count = 0
            Dim m_xmld As New XmlDocument()
            m_xmld.Load(sReader)

            'Get the list of name nodes 
            m_nodelist = m_xmld.SelectNodes("/FireBaseXML/item")

            Dim nodelist As XmlNodeList = m_xmld.DocumentElement.ChildNodes
            For Each outerNode As XmlNode In nodelist
                Dim name = outerNode.Name
                Dim xtim As Array
                Dim xrhr As Array

                If name = "item" Then
                    'oitem = New oItem(1, 250, "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 20, 20, 0, 0, 0.3, 0.3, 0, 0, 0, "generic", "description", 20, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, 1, 0, 0, 0.3, 0, 0, 3)
                    oitem = New oItem("0,0", 98.4, 0.338, 2, 684, 60, 0.101, 1, 250, "", "OBJ", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 427, 481, 0, 0, 0, 0, 9.5, 22, 0, 0, 0.3, 0.3, 0, 0, 0, "generic", "description", 20, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, 1, RoomLength(fireroom) / 2, RoomWidth(fireroom) / 2, 0.35, 0, 0, 3, 0)

                    count = count + 1
                    oitem.id = count


                    For Each InnerNode As XmlNode In outerNode.ChildNodes

                        Dim iname = InnerNode.Name
                        Dim x = InnerNode.InnerText


                        Select Case iname
                            Case "initial_mass"
                                oitem.mass = x
                            Case "heat_of_combustion"
                                odistribution = New oDistribution("", "", "None", CSng(x / 1000), 0, 0, 0, 0, 0, 0, 0)
                                odistribution.id = oitem.id
                                odistribution.varname = "heat of combustion"
                                oitemdistributions.Add(odistribution)
                            Case Else
                        End Select

                        If InnerNode.HasChildNodes Then
                            For Each InnerNode2 As XmlNode In InnerNode.ChildNodes
                                Dim iname2 = InnerNode2.Name
                                Dim x2 = InnerNode2.InnerText

                                Select Case iname
                                    Case "description"
                                        Select Case iname2
                                            Case "summary"
                                                oitem.description = x2
                                            Case "details"
                                                oitem.detaileddescription = x2
                                        End Select

                                    Case "data"
                                        Select Case iname2
                                            Case "time"
                                                xtim = x2.Split(",")
                                            Case "rhr"
                                                xrhr = x2.Split(",")
                                        End Select
                                End Select

                            Next

                        End If

                    Next

                    For i = 0 To xtim.GetUpperBound(0)
                        If i > xrhr.GetUpperBound(0) Then
                            oitem.hrr = oitem.hrr & xtim(i) & "," & ","
                        Else
                            oitem.hrr = oitem.hrr & xtim(i) & "," & xrhr(i) & ","
                        End If
                    Next

                    If oitem IsNot Nothing Then
                        oItems.Add(oitem)
                    End If
                End If

            Next

            If oItems IsNot Nothing Then
                ItemDB.SaveWebItems(oItems, oitemdistributions)

                Dim j As Integer = 0
                frmwebitemlist.lstItems.Items.Clear()
                For Each p As oItem In oItems
                    j = j + 1

                    Dim str = p.GetDisplayText(vbTab)
                    If str.Length > 40 Then str = str.Remove(40, str.Length - 40) 'truncate to length 40 char

                    frmwebitemlist.lstItems.Items.Add(j.ToString & vbTab & str)
                Next

                frmwebitemlist.Show()

            End If

        Catch ex As Exception
            'Console.WriteLine(ex.Message)
            'Console.ReadKey()
        End Try

    End Sub


    Private Sub cmdAlphaT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAlphaT.Click

        If useT2fire = True Then
            frmPowerlaw.optAlpha2.Checked = True
            frmPowerlaw.txtStoreHeight.Enabled = False
        Else
            frmPowerlaw.optAlpha3.Checked = True
            frmPowerlaw.txtStoreHeight.Enabled = True
        End If

        frmPowerlaw.txtStoreHeight.Text = StoreHeight
        frmPowerlaw.txtPeakHRR.Text = PeakHRR
        frmPowerlaw.TxtAlphaT.Text = AlphaT

        Dim odistributions As New List(Of oDistribution)
        odistributions = DistributionClass.GetDistributions
        For Each oDistribution In odistributions
            Select Case oDistribution.varname
                Case "Alpha T"
                    If oDistribution.distribution <> "None" Then
                        frmPowerlaw.TxtAlphaT.BackColor = distbackcolor
                    Else
                        frmPowerlaw.TxtAlphaT.BackColor = distnobackcolor
                    End If
                Case "Peak HRR"
                    If oDistribution.distribution <> "None" Then
                        frmPowerlaw.txtPeakHRR.BackColor = distbackcolor
                    Else
                        frmPowerlaw.txtPeakHRR.BackColor = distnobackcolor
                    End If
            End Select
        Next

        frmPowerlaw.Show()

    End Sub

    Private Sub chkAlphaTFire_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAlphaTFire.CheckedChanged
        If chkAlphaTFire.Checked Then
            usepowerlawdesignfire = True
        Else
            usepowerlawdesignfire = False
        End If
        FillItemList()
    End Sub

    Private Sub lstFireRoom2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstFireRoom2.SelectedIndexChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================

        Dim id As Short
        On Error Resume Next

        If lstFireRoom2.SelectedIndex = -1 Then lstFireRoom2.SelectedIndex = 0
        id = lstFireRoom2.SelectedIndex + 1
        fireroom = id

        'need to re-centre the item in the new room
        If VM2 = True Then
            Dim oItems As List(Of oItem)
            Dim oItemDistributions As List(Of oDistribution)
            oItems = ItemDB.GetItemsv2
            oItemDistributions = ItemDB.GetItemDistributions
            For Each oitem In oItems
                oitem.xleft = RoomLength(fireroom) / 2
                oitem.ybottom = RoomWidth(fireroom) / 2
            Next
            ItemDB.SaveItemsv2(oItems, oItemDistributions)
        End If

    End Sub

    Private Sub lstFireRoom2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstFireRoom2.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*  ===================================================

        Dim id As Short

        id = lstFireRoom2.SelectedIndex + 1
        fireroom = id
    End Sub

    Private Sub cmdRoomPopulate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRoomPopulate.Click
        Dim numfiles As Integer
        numfiles = My.Computer.FileSystem.GetFiles(RiskDataDirectory, VB.FileIO.SearchOption.SearchTopLevelOnly, "input*.xml").Count
        frmPopulate.updown_itcounter.Maximum = numfiles
        frmPopulate.txtGridSize.Text = gridsize.ToString
        frmPopulate.txtVentClearance.Text = ventclearance.ToString
        frmPopulate.txtWindSpeed.Text = ISD_windspeed.ToString
        frmPopulate.txt_winddir.Text = ISD_winddir.ToString
        frmPopulate.BringToFront()
        frmPopulate.Show()

        frmPopulate.cmdPopulate.PerformClick()




    End Sub

    Private Sub txtFLED_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFLED.Validated
        ErrorProvider1.Clear()

        FLED = CSng(txtFLED.Text)

        frmOptions1.Update_Distributions("Fire Load Energy Density")

    End Sub

    Private Sub txtFLED_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtFLED.Validating
        If IsNumeric(txtFLED.Text) Then
            If (CSng(txtFLED.Text) <= 50000 And CSng(txtFLED.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtFLED.Select(0, txtFLED.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtFLED, "Invalid Entry. Fire Load Energy Density must be in the range 0 to 50,000 MJ/m2. ")

    End Sub

    Private Sub cmdDist_FLED_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_FLED.Click
        Dim param As String = "Fire Load Energy Density"
        Dim units As String = "MJ/m2"
        Dim instruction As String = "Fire load energy density per unit floor area"

        Call frmDistParam.ShowDistributionForm(param, units, instruction)

        'If oitemdistributions(2).distribution <> "None" Then
        '    txtFLED.BackColor = Color.LightSalmon
        'Else
        '    txtFLED.BackColor = Color.White
        'End If


    End Sub



    Private Sub cmdUtiskul_Click(sender As Object, e As EventArgs)

    End Sub
End Class