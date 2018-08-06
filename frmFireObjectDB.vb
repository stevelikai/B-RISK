Imports System.Data.OleDb
Imports System.Collections.Generic
Public Class frmFireObjectDB

    Private Sub Fire_DataBindingNavigatorSaveItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Fire_DataBindingNavigatorSaveItem.Click

        'handle ADO.net errors
        Try
            Me.Fire_DataBindingSource.EndEdit()
            Me.TableAdapterManager.UpdateAll(Me.FireDataSet1)
        Catch ex As ArgumentException
            MessageBox.Show(ex.Message, "Argument Exception")
            Fire_DataBindingSource.CancelEdit()
        Catch ex As DBConcurrencyException
            MessageBox.Show("A concurrency error occurred. " & "Some rows were not updated.", "Concurrency error")
            Fire_DataTableAdapter.Fill(Me.FireDataSet1.Fire_Data)
        Catch ex As DataException
            MessageBox.Show(ex.Message, ex.GetType.ToString)
            Fire_DataBindingSource.CancelEdit()
        Catch ex As OleDbException
            MessageBox.Show("Database error # " & ex.Message & ": " & ex.GetType.ToString)
        End Try

    End Sub

    Private Sub frmFireObjectDB_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'capture OleDB exception
        Try
            Me.Fire_DataTableAdapter.Fill(Me.FireDataSet1.Fire_Data)
        Catch ex As OleDbException
            MessageBox.Show("Database error # " & ex.ErrorCode & ": " & ex.GetType.ToString)
        End Try

        If Me.Object_TypeComboBox.Text = "T Squared Fire" Then
            GroupBox3.Visible = True
        Else
            GroupBox3.Visible = False
        End If
        plot_graph()
    End Sub

    Private Sub FuelComboBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FuelComboBox.SelectedIndexChanged
        Me.Energy_YieldTextBox.Enabled = False
        Me.Soot_YieldTextBox.Enabled = False
        Me.HCN_YieldTextBox.Enabled = False
        Me.CO2_YieldTextBox.Enabled = False

        Select Case Me.FuelComboBox.Text

            Case "User Specified"
                Me.Energy_YieldTextBox.Enabled = True
                Me.Soot_YieldTextBox.Enabled = True
                Me.HCN_YieldTextBox.Enabled = True
                Me.CO2_YieldTextBox.Enabled = True
            Case "Ethanol"
                Me.Energy_YieldTextBox.Text = CStr(25.6) 'energy
                Me.Soot_YieldTextBox.Text = CStr(0.008) 'soot
                Me.HCN_YieldTextBox.Text = CStr(0) 'hcn
                Me.CO2_YieldTextBox.Text = CStr(1.51) 'co2
            Case "Heptane"
                Me.Energy_YieldTextBox.Text = CStr(44.6)  'energy
                Me.Soot_YieldTextBox.Text = CStr(0.037) 'soot
                Me.HCN_YieldTextBox.Text = CStr(0) 'hcn
                Me.CO2_YieldTextBox.Text = CStr(2.85) 'co2
            Case "Hexane"
                Me.Energy_YieldTextBox.Text = CStr(41.5)  'energy
                Me.Soot_YieldTextBox.Text = CStr(0.035) 'soot
                Me.HCN_YieldTextBox.Text = CStr(0) 'hcn
                Me.CO2_YieldTextBox.Text = CStr(1.87) 'co2
            Case "Methane"
                Me.Energy_YieldTextBox.Text = CStr(49.6)  'energy
                Me.Soot_YieldTextBox.Text = CStr(0.013) 'soot
                Me.HCN_YieldTextBox.Text = CStr(0) 'hcn
                Me.CO2_YieldTextBox.Text = CStr(1.2) 'co2
            Case "Propane"
                Me.Energy_YieldTextBox.Text = CStr(43.7)  'energy
                Me.Soot_YieldTextBox.Text = CStr(0.024) 'soot
                Me.HCN_YieldTextBox.Text = CStr(0) 'hcn
                Me.CO2_YieldTextBox.Text = CStr(2.34) 'co2
            Case "Polyurethane Foam (GM23)" 'gm23
                Me.Energy_YieldTextBox.Text = CStr(19)  'energy
                Me.Soot_YieldTextBox.Text = CStr(0.227) 'soot
                Me.HCN_YieldTextBox.Text = CStr(0.01) 'hcn
                Me.CO2_YieldTextBox.Text = CStr(1.92) 'co2
            Case "Polymethylmethacrylate"
                Me.Energy_YieldTextBox.Text = CStr(24.2)  'energy
                Me.Soot_YieldTextBox.Text = CStr(0.022) 'soot
                Me.HCN_YieldTextBox.Text = CStr(0) 'hcn
                Me.CO2_YieldTextBox.Text = CStr(1.69) 'co2
            Case "Wood"
                Me.Energy_YieldTextBox.Text = CStr(12.4)  'energy
                Me.Soot_YieldTextBox.Text = CStr(0.015) 'soot
                Me.HCN_YieldTextBox.Text = CStr(0) 'hcn
                Me.CO2_YieldTextBox.Text = CStr(1.19) 'co2
            Case Else
        End Select
    End Sub

    Private Sub Growth_RateComboBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Growth_RateComboBox.SelectedIndexChanged
        Call TSquared()
    End Sub
    Private Sub TSquared()
        Dim alpha, count As Single
        Dim data As String
        Dim hrr As Single
        Dim MaxT, T As Double
        On Error Resume Next
        If Me.Object_TypeComboBox.Text = "T Squared Fire" Then
            alpha = 0
            T = 0
            count = 0
            data = ""

            If Me.Growth_RateComboBox.Text = "Ultra Fast" Then
                alpha = 0.1876
            ElseIf Me.Growth_RateComboBox.Text = "Fast" Then
                alpha = 0.0469
            ElseIf Me.Growth_RateComboBox.Text = "Medium" Then
                alpha = 0.01172
            ElseIf Me.Growth_RateComboBox.Text = "Slow" Then
                alpha = 0.00293
            End If

            If IsNumeric(Me.Max_HRRTextBox.Text) Then
                count = CSng(Me.Max_HRRTextBox.Text)
            End If
            MaxT = System.Math.Sqrt(count / alpha)
            Do While T < MaxT
                hrr = CSng(alpha * T ^ 2)
                data = data & CStr(Format(T, "0.0")) & "," & CStr(Format(hrr, "0.0")) & vbCrLf
                T = T + (MaxT / 30)
            Loop
            hrr = CSng(alpha * MaxT ^ 2)
            data = data & CStr(Format(MaxT, "0.0")) & "," & CStr(Format(hrr, "0.0")) & vbCrLf
            Me.Time___Heat_ReleaseTextBox.Text = data
            Call plot_graph()
        End If
    End Sub

    Private Sub plot_graph()
        '*  ====================================================
        '*      Show a graph of heat release rate versus time for
        '*      the current object.
        '*  ====================================================

        Dim i As Integer, id As Integer, tim As Double, hrr As Double
        Dim NumberPoints As Integer, HRRData(,,) As Double, fname As String, strTempPath As String

        Dim tempfile As String
        Dim linetext As String = ""
        ReDim HRRData(0 To 2, 0 To 500, 0 To 1)

        linetext = Me.Time___Heat_ReleaseTextBox.Text
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
                    Exit Sub
                End If
            End If
        Next

        NumberPoints = 0

            Dim str_data As String() = linetext.Split(CChar(","))
            i = str_data.Count - 1

            If i < 1 Then Exit Sub

            Dim chdata(0 To CInt(str_data.Count / 2 - 1), 0 To 1) As Double
            Dim j As Integer
            j = 0
        Do While (NumberPoints <= i And NumberPoints < 1000 And j <= UBound(chdata))

            NumberPoints = NumberPoints + 2
            If IsNumeric(str_data(NumberPoints - 2)) Then
                HRRData(1, j + 1, 1) = CDbl(str_data(NumberPoints - 2).Normalize) 'time
                HRRData(2, j + 1, 1) = CDbl(str_data(NumberPoints - 1).Normalize) 'hrr
                chdata(j, 0) = CDbl(str_data(NumberPoints - 2).Normalize) 'time
                chdata(j, 1) = CDbl(str_data(NumberPoints - 1).Normalize) 'hrr
            End If
            j = j + 1
        Loop

        'AxMSChart1.chartType = MSChart20Lib.VtChChartType.VtChChartType2dXY
        '  AxMSChart1.Plot.UniformAxis = False
        On Error Resume Next
        ' AxMSChart1.ChartData = chdata

    End Sub

    Private Sub Object_TypeComboBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Object_TypeComboBox.SelectedIndexChanged
        If Me.Object_TypeComboBox.Text = "T Squared Fire" Then
            GroupBox3.Visible = True
        Else
            GroupBox3.Visible = False
        End If
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Fire_DataBindingSource.CancelEdit()
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub BindingNavigatorPositionItem_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles BindingNavigatorPositionItem.TextChanged
        Call plot_graph()
    End Sub

    Private Sub Max_HRRTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Max_HRRTextBox.TextChanged
        Call TSquared()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        '**********************************************************
        '*  Include the current object in the fire simulation
        '**********************************************************
        Dim HRRData(,,) As Double
        Dim id As Short
        Dim i As Integer

        Dim oitem = New oItem

        'add objects to list box
        NumberObjects = CShort(NumberObjects + 1)

        'call procedure to resize the vent arrays
        Resize_Objects()

        'add extra label to the listbox
        frmFire.lstObjectID.Items.Clear()
        For i = 1 To NumberObjects
            frmFire.lstObjectID.Items.Add("Fire " & CStr(i) & " of " & CStr(NumberObjects))
        Next i

        frmFire.lstObjectID.SelectedIndex = NumberObjects - 1

        id = NumberObjects

        If id = 0 Then Exit Sub

        'store the energy yield as a variable
        If IsNumeric(Me.Energy_YieldTextBox.Text) = True Then EnergyYield(id) = CSng(Me.Energy_YieldTextBox.Text)
        frmFire.txtEnergyYield.Text = CStr(EnergyYield(id))
        oitem.hoc = CStr(EnergyYield(id))

        'store the carbon dioxide yield as a variable
        If IsNumeric(Me.CO2_YieldTextBox.Text) = True Then CO2Yield(id) = CSng(Me.CO2_YieldTextBox.Text)
        frmFire.txtCO2Yield.Text = CStr(CO2Yield(id))
        oitem.co2 = CStr(CO2Yield(id))

        'store the height of the base of the fire as a variable
        If IsNumeric(Me.Fire_HeightTextBox.Text) = True Then FireHeight(id) = CSng(Me.Fire_HeightTextBox.Text)
        frmFire.txtFireHeight.Text = CStr(FireHeight(id))
        oitem.elevation = CStr(FireHeight(id))

        'store the soot yield as a variable
        If IsNumeric(Me.Soot_YieldTextBox.Text) = True Then SootYield(id) = CSng(Me.Soot_YieldTextBox.Text)
        frmFire.txtSootYield.Text = CStr(SootYield(id))
        oitem.soot = CStr(SootYield(id))

        'store the hcn yield as a variable
        If IsNumeric(Me.HCN_YieldTextBox.Text) = True Then HCNuserYield(id) = CSng(Me.HCN_YieldTextBox.Text)
        frmFire.txtHCNYield.Text = CStr(HCNuserYield(id))

        'store the object description
        ObjectDescription(id) = Me.DescriptionTextBox.Text
        frmFire.lblObjectDescription.Text = ObjectDescription(id)
        oitem.description = ObjectDescription(id)

        Dim linetext As String = ""
        Dim numberpoints As Integer
        ReDim HRRData(0 To 2, 0 To 500, 0 To 1)

        linetext = Me.Time___Heat_ReleaseTextBox.Text
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

        oitem.hrr = linetext

        For i = 0 To linetext.Length - 1
            If linetext.Chars(i) = ","c Or linetext.Chars(i) = "."c Or linetext.Chars(i) = "+"c Or linetext.Chars(i) = "E"c Then
            Else
                If IsNumeric(linetext.Chars(i)) Then
                Else
                    MsgBox("Invalid RHR versus time data")
                    Exit Sub
                End If
            End If
        Next

        NumberPoints = 0

        Dim str_data As String() = linetext.Split(CChar(","))
        i = str_data.Count - 1

        If i < 1 Then Exit Sub

        Dim j As Integer
        j = 0
        NumberDataPoints(id) = 0
        Do While (numberpoints < i And numberpoints < 1000)
            numberpoints = numberpoints + 2
            If IsNumeric(str_data(numberpoints - 2)) And IsNumeric(str_data(numberpoints - 1)) Then
                HeatReleaseData(1, NumberDataPoints(id) + 1, id) = CDbl(str_data(numberpoints - 2).Normalize) 'time
                HeatReleaseData(2, NumberDataPoints(id) + 1, id) = CDbl(str_data(numberpoints - 1).Normalize) 'hrr
            End If
            j = j + 1
            NumberDataPoints(id) = CShort(NumberDataPoints(id) + 1)
        Loop

99:

        'MDIFrmMain.mnuRun.Enabled = True
        frmFire.GroupBox1.Visible = True
        Me.Close()

        frmFire.Show()

        Exit Sub
    End Sub

    Private Sub ToolStripLabel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel1.Click
        Dim oitem As oItem
        Dim oItems As List(Of oItem)
        Dim odistribution As oDistribution
        Dim oitemdistributions As List(Of oDistribution)

        oItems = ItemDB.GetItemsv2()
        'Dim oitem As oItem
        oitemdistributions = ItemDB.GetItemDistributions
        ' Dim oitem As oItem = oItems(i)
        'oitemdistributions = ItemDB.GetItemDistributions

        'NewItemForm = New frmNewItem
        'NewItemForm.Text = "New Item"
        oitem = New oItem("0,0", 98.4, 0.338, 2, 684, 60, 0.101, 1, 250, DescriptionTextBox.Text, "OBJ", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 427, 481, 0, 0, 0, 0, 9.5, 22, 0, 0, 0.3, 0.3, 0, 0, 0, Object_TypeComboBox.Text, DescriptionTextBox.Text, 20, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, 1, RoomLength(fireroom) / 2, RoomWidth(fireroom) / 2, 0.35, 0, 0, 3, 0)

        oitem.id = oItems.Count + 1
        oitem.description = DescriptionTextBox.Text

        oitem.type = Object_TypeComboBox.Text
        oitem.elevation = Fire_HeightTextBox.Text
        'oitem.height = txtItemHeight.Text
        'oitem.critflux = txtCritFlux.Text
        'oitem.critfluxauto = txtCritFluxauto.Text
        'oitem.ftplimitpilot = txtFTPlimitpilot.Text
        'oitem.ftplimitauto = txtFTPlimitauto.Text
        'oitem.ftpindexpilot = txtFTPindexpilot.Text
        'oitem.ftpindexauto = txtFTPindexauto.Text
        'oitem.prob = txtItemProb.Text
        'oitem.mass = txtMass.Text
        'oitem.hoc = Energy_YieldTextBox.Text
        'oitem.soot = Soot_YieldTextBox.Text
        'oitem.co2 = CO2_YieldTextBox.Text

        oitem.hrr = Time___Heat_ReleaseTextBox.Text
        'oitem.ignitiontime = txtIgnitionTime.Text
        'oitem.maxnumitem = txtMaxNumItem.Text
        'oitem.xleft = txtXleft.Text
        'oitem.ybottom = txtYbottom.Text
        'oitem.radlossfrac = txtRadLossFracItem.Text
        If oitem IsNot Nothing Then
            oItems.Add(oitem)
        End If

        odistribution = New oDistribution("", "", "None", CSng(Energy_YieldTextBox.Text), 0, 0, 0, 0, 0, 0, 0)
        odistribution.id = oitem.id
        odistribution.varname = "heat of combustion"
        'NewItemForm.txtValue8.Text = odistribution.varvalue
        oitemdistributions.Add(odistribution)

        odistribution = New oDistribution("", "", "None", CSng(Soot_YieldTextBox.Text), 0, 0, 0, 0, 0, 0, 0)
        odistribution.id = oitem.id
        odistribution.varname = "soot yield"
        'NewItemForm.txtSootValue.Text = odistribution.varvalue
        oitemdistributions.Add(odistribution)

        odistribution = New oDistribution("", "", "None", CSng(CO2_YieldTextBox.Text), 0, 0, 0, 0, 0, 0, 0)
        odistribution.id = oitem.id
        odistribution.varname = "co2 yield"
        'NewItemForm.txtCO2Value.Text = odistribution.varvalue
        oitemdistributions.Add(odistribution)

        odistribution = New oDistribution("", "", "None", 3, 0, 0, 0, 0, 0, 0, 0)
        odistribution.id = oitem.id
        odistribution.varname = "Latent Heat of Gasification"
        'NewItemForm.txtCO2Value.Text = odistribution.varvalue
        oitemdistributions.Add(odistribution)

        odistribution = New oDistribution("", "", "None", 0.3, 0, 0, 0, 0, 0, 0, 0)
        odistribution.id = oitem.id
        odistribution.varname = "Radiant Loss Fraction"
        oitemdistributions.Add(odistribution)

        odistribution = New oDistribution("", "", "None", 250, 0, 0, 0, 0, 0, 0, 0)
        odistribution.id = oitem.id
        odistribution.varname = "HRRUA"
        oitemdistributions.Add(odistribution)

        If oItems IsNot Nothing Then
            ItemDB.SaveItemsv2(oItems, oitemdistributions)
            frmItemList.FillItemList()
        End If

        Me.Close()
        frmItemList.Show()
        frmItemList.BringToFront()
    End Sub
End Class