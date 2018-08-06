Imports System.Collections.Generic
Imports System.Windows.Forms.DataVisualization.Charting

Public Class frmNewItem
    Public itemid As Integer
    Public oitem As oItem
    Public oItems As List(Of oItem)
    Public oitemdistribution As oDistribution
    Public oitemdistributions As List(Of oDistribution)

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        'save new item data
        If IsValidData() Then

            Call cmdMassCalc.PerformClick() 'calculate the mass using area under the HRR curve

            oItems = ItemDB.GetItemsv2()
            oitemdistributions = ItemDB.GetItemDistributions()

            For Each Me.oitem In oItems

                If Me.oitem.id = Me.Tag Then
                    ObjectMass(Me.Tag) = CDbl(txtMass.Text)

                    oitem.description = txtItemDescription.Text
                    oitem.detaileddescription = txtDetailedDescription.Text
                    oitem.userlabel = txtUserLabel.Text
                    oitem.type = ComboType.Text

                    oitem.length = txtItemLength.Text
                    oitem.width = txtItemWidth.Text
                    oitem.elevation = txtItemElevation.Text
                    oitem.height = txtItemHeight.Text
                    oitem.critflux = txtCritFlux.Text
                    oitem.critfluxauto = txtCritFluxauto.Text
                    oitem.ftplimitpilot = txtFTPlimitpilot.Text
                    oitem.ftplimitauto = txtFTPlimitauto.Text
                    oitem.ftpindexpilot = txtFTPindexpilot.Text
                    oitem.ftpindexauto = txtFTPindexauto.Text
                    oitem.prob = txtItemProb.Text
                    oitem.mass = txtMass.Text

                    'oitem.hoc = txtValue8.Text
                    'oitem.soot = txtSootValue.Text
                    'oitem.co2 = txtCO2Value.Text

                    'clean up hrr data, strip out spaces
                    Dim hrrdata As String = Time___Heat_ReleaseTextBox.Text
                    hrrdata = hrrdata.Replace(Chr(9), ",") 'horizontal tab
                    hrrdata = hrrdata.Replace(Chr(10), ",") 'line feed
                    hrrdata = hrrdata.Replace(Chr(13), "")
                    hrrdata = hrrdata.Replace(" ", "")

                    'clean up mlr data, strip out spaces
                    Dim mlrdata As String = txtFreeBurnMLR.Text
                    mlrdata = mlrdata.Replace(Chr(9), ",") 'horizontal tab
                    mlrdata = mlrdata.Replace(Chr(10), ",") 'line feed
                    mlrdata = mlrdata.Replace(Chr(13), "")
                    mlrdata = mlrdata.Replace(" ", "")

                    oitem.hrr = hrrdata
                    oitem.mlrfreeburn = mlrdata
                    oitem.ignitiontime = txtIgnitionTime.Text
                    oitem.maxnumitem = txtMaxNumItem.Text
                    oitem.xleft = txtXleft.Text
                    oitem.ybottom = txtYbottom.Text
                    oitem.radlossfrac = txtRadLossFracItem.Text
                    oitem.constantA = txtConstA.Text
                    oitem.constantB = txtConstB.Text
                    'oitem.LHoG = txtLHOG.Text
                    oitem.HRRUA = txtHRRUA.Text

                    If optwind.Checked = True Then
                        oitem.windeffect = 2
                    Else
                        oitem.windeffect = 1
                    End If
                    If optPoolFire.Checked = True Then
                        oitem.pyrolysisoption = 1
                    ElseIf optWoodCrib.Checked = True Then
                        oitem.pyrolysisoption = 2
                    ElseIf optFreeBurn.Checked = True Then
                        oitem.pyrolysisoption = 0
                    End If

                    For Each x In oitemdistributions
                        If x.id = oitem.id Then
                            If x.varname = "heat of combustion" Then
                                x.varvalue = txtValue8.Text
                            End If
                            If x.varname = "soot yield" Then
                                x.varvalue = txtSootValue.Text
                            End If
                            If x.varname = "co2 yield" Then
                                x.varvalue = txtCO2Value.Text
                            End If
                            If x.varname = "Latent Heat of Gasification" Then
                                x.varvalue = txtLHOG.Text
                            End If
                            If x.varname = "Radiant Loss Fraction" Then
                                x.varvalue = txtRadLossFracItem.Text
                            End If
                            If x.varname = "HRRUA" Then
                                x.varvalue = txtHRRUA.Text
                            End If
                        End If
                    Next

                End If
            Next

            '  End If


            If oitem IsNot Nothing Then
                ItemDB.SaveItemsv2(oItems, oitemdistributions)
                frmItemList.FillItemList()
            End If

            Me.Close()
        End If
    End Sub
    Private Function IsValidData() As Boolean

        Return Validator.IsPresent(txtValue8) AndAlso
               Validator.IsPresent(txtCO2Value) AndAlso
               Validator.IsPresent(txtItemWidth) AndAlso
               Validator.IsPresent(txtItemHeight) AndAlso
               Validator.IsPresent(txtItemElevation) AndAlso
               Validator.IsPresent(txtCOValue) AndAlso
               Validator.IsPresent(txtItemLength) AndAlso
               Validator.IsPresent(txtSootValue) AndAlso
               Validator.IsPresent(txtItemDescription) AndAlso
               Validator.IsPresent(txtItemProb) AndAlso
               Validator.IsPresent(txtCritFlux) AndAlso
               Validator.IsPresent(txtCritFluxauto) AndAlso
               Validator.IsPresent(txtFTPlimitpilot) AndAlso
               Validator.IsPresent(txtFTPlimitauto) AndAlso
               Validator.IsPresent(txtFTPindexpilot) AndAlso
               Validator.IsPresent(txtFTPindexauto) AndAlso
               Validator.IsPresent(txtIgnitionTime) AndAlso
               Validator.IsPresent(txtMaxNumItem) AndAlso
               Validator.IsPresent(txtXleft) AndAlso
               Validator.IsPresent(txtYbottom) AndAlso
                Validator.IsPresent(txtLHOG) AndAlso
                 Validator.IsPresent(txtHRRUA) AndAlso
            Validator.IsPresent(txtRadLossFracItem) AndAlso
         Validator.IsPresent(txtMass)

    End Function

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim i As Long, id As Integer, j As Integer
        Dim NumberPoints As Integer
        Dim title As String
        Dim count As Integer
        Dim linetext As String = ""
        Dim NumberDataPoints As Integer
        Dim chdata(,) As Object
        Dim ydata() As Double
        Dim x() As Double
        Dim y() As Double

        If FuelResponseEffects = True Then
            title = "Free burn mass loss rate (kg/s)"
            linetext = Me.txtFreeBurnMLR.Text
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.000"
        Else
            title = "Rate of heat release (kW)"
            linetext = Me.Time___Heat_ReleaseTextBox.Text
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
        End If

        linetext = CStr(linetext.Trim)
        linetext = linetext.Replace(Chr(9), ",")
        linetext = linetext.Replace(Chr(10), ",")
        linetext = linetext.Replace(Chr(13), "")
        linetext = linetext.Replace(" ", "")
        linetext = linetext.Replace(",,", ",")

        'remove the comma off the end
        If linetext.Length > 0 Then
            If linetext.Chars(linetext.Length - 1) = "," Then linetext = linetext.Remove(linetext.Length - 1, 1)
        Else
            Exit Sub
        End If

        For i = 0 To linetext.Length - 1
            If linetext.Chars(i) = ","c Or linetext.Chars(i) = "."c Then
            Else
                If IsNumeric(linetext.Chars(i)) Then
                Else
                    MsgBox("Invalid data")
                    Exit Sub
                End If
            End If
        Next

        count = 0

        Dim str_data As String() = linetext.Split(CChar(","))
        i = str_data.Count - 1

        If i < 1 Then Exit Sub

        j = 0
        NumberDataPoints = 0
        Do While (count < i)

            ReDim Preserve x(0 To NumberDataPoints)
            ReDim Preserve y(0 To NumberDataPoints)

            count = count + 2
            If IsNumeric(str_data(count - 2)) And IsNumeric(str_data(count - 1)) Then
                x(NumberDataPoints) = CDbl(str_data(count - 2).Normalize) 'time
                y(NumberDataPoints) = CDbl(str_data(count - 1).Normalize) 'hrr

            End If
            j = j + 1
            NumberDataPoints = NumberDataPoints + 1
        Loop

        NumberPoints = NumberDataPoints

        ReDim chdata(0 To NumberDataPoints, 0 To 1)
        ReDim ydata(0 To NumberDataPoints)

        frmPlot.Chart1.Series.Clear()

        frmPlot.Chart1.Series.Add("Fire " & id)

        frmPlot.Chart1.Series("Fire " & id).ChartType = SeriesChartType.FastLine

        For i = 1 To NumberPoints
            chdata(i - 1, 0) = x(i - 1) 'time
            chdata(i - 1, 1) = y(i - 1) 'hrr/mlr
            ydata(i - 1) = chdata(i - 1, 1) 'hrr/mlr
            frmPlot.Chart1.Series(("Fire " & id)).Points.AddXY(chdata(i - 1, 0), ydata(i - 1))
        Next

        ' Next room
        frmPlot.Chart1.BackColor = Color.AliceBlue

        frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
        frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
        frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = title

        frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
        frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
        frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
        frmPlot.Chart1.Legends.Clear()
        frmPlot.Chart1.Titles("Title1").Text = txtItemDescription.Text
        frmPlot.Chart1.Visible = True
        frmPlot.Show()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub


    Private Sub cmdMassCalc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMassCalc.Click

        Dim i As Long, j As Integer
        Dim count As Integer
        Dim linetext As String = ""
        Dim NumberDataPoints As Integer
        Dim x() As Double
        Dim y() As Double

        If FuelResponseEffects = False Then
            linetext = Me.Time___Heat_ReleaseTextBox.Text
        Else
            linetext = Me.txtFreeBurnMLR.Text
        End If
        linetext = CStr(linetext.Trim)
        linetext = linetext.Replace(Chr(9), ",")
        linetext = linetext.Replace(Chr(10), ",")
        linetext = linetext.Replace(Chr(13), "")
        linetext = linetext.Replace(" ", "")
        linetext = linetext.Replace(",,", ",")

        'remove the comma off the end
        If linetext.Length > 0 Then
            If linetext.Chars(linetext.Length - 1) = "," Then linetext = linetext.Remove(linetext.Length - 1, 1)
        Else
            Exit Sub
        End If

        For i = 0 To linetext.Length - 1
            If linetext.Chars(i) = ","c Or linetext.Chars(i) = "."c Then
            Else
                If IsNumeric(linetext.Chars(i)) Then
                Else
                    MsgBox("Invalid RHR versus time data")
                    Exit Sub
                End If
            End If
        Next

        count = 0

        Dim str_data As String() = linetext.Split(CChar(","))
        i = str_data.Count - 1

        If i < 1 Then Exit Sub

        j = 0
        NumberDataPoints = 0
        Do While (count < i)

            ReDim Preserve x(0 To NumberDataPoints)
            ReDim Preserve y(0 To NumberDataPoints)

            count = count + 2
            If IsNumeric(str_data(count - 2)) And IsNumeric(str_data(count - 1)) Then
                x(NumberDataPoints) = CDbl(str_data(count - 2).Normalize) 'time
                y(NumberDataPoints) = CDbl(str_data(count - 1).Normalize) 'hrr
            End If
            j = j + 1
            NumberDataPoints = NumberDataPoints + 1
        Loop

        'trapezoidal rule
        Dim sum As Double = 0
        For i = 0 To NumberDataPoints - 2
            sum = sum + ((y(i) + y(i + 1)) / 2 * (x(i + 1) - x(i)))
        Next

        'convert kJ to MJ 
        'sum = sum / 1000

        Dim hc As Double = CDbl(txtValue8.Text)

        Dim mass As Double = 0
        If FuelResponseEffects = False Then
            mass = sum / hc / 1000 'kg
        Else
            mass = sum
        End If

        txtMass.Text = CStr(Format(mass, "0.000"))


    End Sub
    Private Sub cmdDist_radlossfrac_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim oitems As List(Of oItem)
        Dim oitemdistributions As List(Of oDistribution)

        oitems = ItemDB.GetItemsv2()
        oitemdistributions = ItemDB.GetItemDistributions()
        itemid = Me.Tag - 1
        Dim oitem As oItem = oitems(itemid)
        Dim paramdist As String
        Dim param As String = "Radiant Loss Fraction"
        Dim units As String = "-"
        Dim instruction As String = "radiant loss fraction for item"

        Call frmDistParam.ShowItemDistributionForm(param, units, instruction, oitem, oitemdistributions, itemid, paramdist)

        If paramdist <> "None" Then
            txtRadLossFracItem.BackColor = Color.LightSalmon
        Else
            txtRadLossFracItem.BackColor = Color.White
        End If
        txtRadLossFracItem.Text = oitem.radlossfrac

    End Sub

    Private Sub cmdDist_hoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_hoc.Click

        Dim oitems As List(Of oItem)
        Dim oitemdistributions As List(Of oDistribution)

        oitems = ItemDB.GetItemsv2()
        oitemdistributions = ItemDB.GetItemDistributions()

        itemid = Me.Tag - 1
        Dim oitem As oItem = oitems(itemid)
        Dim paramdist As String
        Dim param As String = "heat of combustion"
        Dim units As String = "kJ/g"
        Dim instruction As String = "Heat of Combustion for Item"

        Call frmDistParam.ShowItemDistributionForm(param, units, instruction, oitem, oitemdistributions, itemid, paramdist)
        If paramdist <> "None" Then
            txtValue8.BackColor = Color.LightSalmon
        Else
            txtValue8.BackColor = Color.White
        End If
        txtValue8.Text = oitem.hoc

    End Sub

    Private Sub cmdDist_soot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_soot.Click

        Dim oitems As List(Of oItem)
        Dim oitemdistributions As List(Of oDistribution)

        oitems = ItemDB.GetItemsv2()
        oitemdistributions = ItemDB.GetItemDistributions()

        itemid = Me.Tag - 1
        Dim oitem As oItem = oitems(itemid)
        Dim paramdist As String
        Dim param As String = "soot yield"
        Dim units As String = "g/g"
        Dim instruction As String = "Soot Yield for Item"

        Call frmDistParam.ShowItemDistributionForm(param, units, instruction, oitem, oitemdistributions, itemid, paramdist)
        If paramdist <> "None" Then
            txtSootValue.BackColor = Color.LightSalmon
        Else
            txtSootValue.BackColor = Color.White
        End If
        txtSootValue.Text = oitem.soot

    End Sub

    Private Sub cmdDist_co2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_co2.Click
        Dim oitems As List(Of oItem)
        Dim oitemdistributions As List(Of oDistribution)

        oitems = ItemDB.GetItemsv2()
        oitemdistributions = ItemDB.GetItemDistributions()

        itemid = Me.Tag - 1
        Dim oitem As oItem = oitems(itemid)
        Dim paramdist As String
        Dim param As String = "co2 yield"
        Dim units As String = "g/g"
        Dim instruction As String = "CO2 Yield for Item"

        Call frmDistParam.ShowItemDistributionForm(param, units, instruction, oitem, oitemdistributions, itemid, paramdist)
        If paramdist <> "None" Then
            txtCO2Value.BackColor = Color.LightSalmon
        Else
            txtCO2Value.BackColor = Color.White
        End If
        txtCO2Value.Text = oitem.co2
    End Sub

    Private Sub cmdDist_lhog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_lhog.Click
        Dim oitems As List(Of oItem)
        Dim oitemdistributions As List(Of oDistribution)

        oitems = ItemDB.GetItemsv2()
        oitemdistributions = ItemDB.GetItemDistributions()

        itemid = Me.Tag - 1
        Dim oitem As oItem = oitems(itemid)
        Dim paramdist As String
        Dim param As String = "Latent Heat of Gasification"
        Dim units As String = "kJ/g"
        Dim instruction As String = "Latent Heat of Gasification for Item"

        Call frmDistParam.ShowItemDistributionForm(param, units, instruction, oitem, oitemdistributions, itemid, paramdist)
        If paramdist <> "None" Then
            txtLHOG.BackColor = Color.LightSalmon
        Else
            txtLHOG.BackColor = Color.White
        End If
        txtLHOG.Text = oitem.LHoG
    End Sub

    Private Sub txtHRRUA_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHRRUA.Validated
        ErrorProvider1.Clear()
    End Sub

    Private Sub txtHRRUA_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtHRRUA.Validating
        If IsNumeric(txtHRRUA.Text) Then
            If CSng(txtHRRUA.Text) >= 0 Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtHRRUA.Select(0, txtHRRUA.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtHRRUA, "Invalid Entry.Heat Release Rate per unit area must be > 0 kW/m2. ")

    End Sub

    Private Sub cmdDist_RLF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_RLF.Click
        Dim oitems As List(Of oItem)
        Dim oitemdistributions As List(Of oDistribution)

        oitems = ItemDB.GetItemsv2()
        oitemdistributions = ItemDB.GetItemDistributions()

        itemid = Me.Tag - 1
        Dim oitem As oItem = oitems(itemid)
        Dim paramdist As String
        Dim param As String = "Radiant Loss Fraction"
        Dim units As String = "-"
        Dim instruction As String = "Radiant Loss Fraction for Item"

        Call frmDistParam.ShowItemDistributionForm(param, units, instruction, oitem, oitemdistributions, itemid, paramdist)
        If paramdist <> "None" Then
            txtRadLossFracItem.BackColor = Color.LightSalmon
        Else
            txtRadLossFracItem.BackColor = Color.White
        End If
        txtRadLossFracItem.Text = oitem.radlossfrac
    End Sub

    Private Sub cmdDist_HRRUA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_HRRUA.Click

        Dim oitems As List(Of oItem)
        Dim oitemdistributions As List(Of oDistribution)

        oitems = ItemDB.GetItemsv2()
        oitemdistributions = ItemDB.GetItemDistributions()

        itemid = Me.Tag - 1
        Dim oitem As oItem = oitems(itemid)
        Dim paramdist As String
        Dim param As String = "HRRUA"
        Dim units As String = "kW/m2"
        Dim instruction As String = "Heat Release Rate per Unit Area"

        Call frmDistParam.ShowItemDistributionForm(param, units, instruction, oitem, oitemdistributions, itemid, paramdist)

        If paramdist <> "None" Then
            txtHRRUA.BackColor = Color.LightSalmon
        Else
            txtHRRUA.BackColor = Color.White
        End If

        txtHRRUA.Text = oitem.HRRUA

    End Sub

    Private Sub frmNewItem_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        If VM2 = True Then
            'NewItemForm.TableLayoutPanel2.Visible = False
            txtFreeBurnMLR.Visible = False
            Time___Heat_ReleaseTextBox.Visible = False
            Label14.Visible = False
            Label50.Visible = False

            txtItemLength.Enabled = False
            txtItemWidth.Enabled = False
            txtItemHeight.Enabled = False
            txtItemProb.Enabled = False
            txtLHOG.Enabled = False
            txtMaxNumItem.Enabled = False
        Else
            'NewItemForm.TableLayoutPanel2.Visible = True
            'txtFreeBurnMLR.Visible = True
            Time___Heat_ReleaseTextBox.Visible = True
            txtItemLength.Enabled = True
            txtItemWidth.Enabled = True
            txtItemHeight.Enabled = True
            txtItemProb.Enabled = True
            txtLHOG.Enabled = True
            txtMaxNumItem.Enabled = True
            Label14.Visible = True
            Label50.Visible = True
        End If

        If DevKey = True And FuelResponseEffects = True Then

            GroupBox_RoomEffects.Visible = True
            Button1.Visible = True
            txtFreeBurnMLR.Visible = True
            Time___Heat_ReleaseTextBox.Visible = False
            Label14.Text = "Mass Loss Rate"
            Label50.Text = "Time (s), MLR (kg/s)"
        ElseIf DevKey = True And FuelResponseEffects = False Then

            GroupBox_RoomEffects.Visible = False
            Button1.Visible = False
            txtFreeBurnMLR.Visible = False
            ' Time___Heat_ReleaseTextBox.Visible = True
            Label14.Text = "Heat Release Rate"
            Label50.Text = "Time (s), HRR (kW)"
        Else

            GroupBox_RoomEffects.Visible = False
            Button1.Visible = False
            txtFreeBurnMLR.Visible = False
            'Time___Heat_ReleaseTextBox.Visible = True
            Label14.Text = "Heat Release Rate"
            Label50.Text = "Time (s), HRR (kW)"
        End If

    End Sub


    Private Sub cmdUtiskul_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'oItems = ItemDB.GetItemsv2()
        'Dim oitem As oItem
        'oitemdistributions = ItemDB.GetItemDistributions
        Dim oitems As List(Of oItem)
        Dim oitemdistributions As List(Of oDistribution)
        oitems = ItemDB.GetItemsv2()
        oitemdistributions = ItemDB.GetItemDistributions()

        itemid = Me.Tag - 1
        Dim oitem As oItem = oitems(itemid)

        frmUtiskul.Show()
        frmUtiskul.Tag = itemid + 1
        frmUtiskul.txtFBMLR.Text = oitem.poolFBMLR
        frmUtiskul.txtDiameter.Text = oitem.pooldiameter
        frmUtiskul.txtDensity.Text = oitem.pooldensity
        frmUtiskul.txtRamp.Text = oitem.poolramp
        frmUtiskul.txtVolume.Text = oitem.poolvolume
        frmUtiskul.txtVtemp.Text = oitem.poolvaptemp

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Dim oitems As List(Of oItem)
        Dim oitemdistributions As List(Of oDistribution)
        oitems = ItemDB.GetItemsv2()
        oitemdistributions = ItemDB.GetItemDistributions()

        itemid = Me.Tag - 1
        Dim oitem As oItem = oitems(itemid)

        frmUtiskul.Show()
        frmUtiskul.Tag = itemid + 1
        frmUtiskul.txtFBMLR.Text = oitem.poolFBMLR
        frmUtiskul.txtDiameter.Text = oitem.pooldiameter
        frmUtiskul.txtDensity.Text = oitem.pooldensity
        frmUtiskul.txtRamp.Text = oitem.poolramp
        frmUtiskul.txtVolume.Text = oitem.poolvolume
        frmUtiskul.txtVtemp.Text = oitem.poolvaptemp
    End Sub

    Private Sub txtFreeBurnMLR_TextChanged(sender As Object, e As EventArgs) Handles txtFreeBurnMLR.TextChanged

    End Sub
End Class