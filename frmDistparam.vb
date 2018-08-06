
Imports System.Collections.Generic

Public Class frmDistParam

    Public oDistribution As oDistribution
    Public oDistributions As List(Of oDistribution)
    Public oItems As List(Of oItem)
    Public oSprinklers As List(Of oSprinkler)
    Public oSmokeDets As List(Of oSmokeDet)
    Public oFans As List(Of oFan)
    Public ovents As List(Of oVent)
    Public ocvents As List(Of oCVent)
    Public oRooms As List(Of oRoom)
    Public itemid As Integer
    Public sprid As Integer
    Public sdid As Integer
    Public fanid As Integer

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_SaveDist.Click
        Dim paramdist As String = "None"
        'save data
        If IsValidData() Then

            If Me.Text = "New Distribution" Then

            ElseIf Me.Text = "Edit Distribution" Then

                oDistributions = DistributionClass.GetDistributions()

                For Each Me.oDistribution In oDistributions
                    
                    If Me.oDistribution.varname = Me.Tag Then

                        Select Case Me.Tag
                            Case "Relative Humidity"
                                'convert from % to -
                                oDistribution.varvalue = txtValue.Text / 100
                                oDistribution.variance = txtVariance.Text / 10000
                                oDistribution.mean = txtMean.Text / 100
                                oDistribution.ubound = txtUpperBound.Text / 100
                                oDistribution.lbound = txtLowerBound.Text / 100
                                oDistribution.distribution = lstDistribution.Text
                                oDistribution.mode = txtMode.Text / 100
                                oDistribution.alpha = txtAlpha.Text
                                oDistribution.beta = txtBeta.Text
                            Case "Interior Temperature"
                                'convert from C to K
                                oDistribution.varvalue = txtValue.Text + 273
                                oDistribution.variance = txtVariance.Text
                                oDistribution.mean = txtMean.Text + 273
                                oDistribution.ubound = txtUpperBound.Text + 273
                                oDistribution.lbound = txtLowerBound.Text + 273
                                oDistribution.distribution = lstDistribution.Text
                                oDistribution.mode = txtMode.Text + 273
                                oDistribution.alpha = txtAlpha.Text
                                oDistribution.beta = txtBeta.Text
                            Case "Exterior Temperature"
                                'convert from C to K
                                oDistribution.varvalue = txtValue.Text + 273
                                oDistribution.variance = txtVariance.Text
                                oDistribution.mean = txtMean.Text + 273
                                oDistribution.ubound = txtUpperBound.Text + 273
                                oDistribution.lbound = txtLowerBound.Text + 273
                                oDistribution.distribution = lstDistribution.Text
                                oDistribution.mode = txtMode.Text + 273
                                oDistribution.alpha = txtAlpha.Text
                                oDistribution.beta = txtBeta.Text

                            Case Else
                                oDistribution.varvalue = txtValue.Text
                                oDistribution.variance = txtVariance.Text
                                oDistribution.mean = txtMean.Text
                                oDistribution.ubound = txtUpperBound.Text
                                oDistribution.lbound = txtLowerBound.Text
                                oDistribution.distribution = lstDistribution.Text
                                oDistribution.mode = txtMode.Text
                                oDistribution.alpha = txtAlpha.Text
                                oDistribution.beta = txtBeta.Text

                        End Select
                        paramdist = oDistribution.distribution
                        Select Case Me.Tag
                            Case "Interior Temperature"
                                oDistribution.units = "K"
                                ' paramdist = oDistribution.distribution
                            Case "Exterior Temperature"
                                oDistribution.units = "K"
                                'paramdist = oDistribution.distribution
                            Case "Relative Humidity"
                                oDistribution.units = "-"
                                'paramdist = oDistribution.distribution
                            Case "Fire Load Energy Density"
                                oDistribution.units = "MJ/m2"
                                ' paramdist = oDistribution.distribution
                            Case "Heat of Combustion PFO"
                                oDistribution.units = "kJ/g"
                                'paramdist = oDistribution.distribution
                            Case "Soot Preflashover Yield"
                                oDistribution.units = "g/g"
                                ' paramdist = oDistribution.distribution
                            Case "CO Preflashover Yield"
                                oDistribution.units = "g/g"
                                'paramdist = oDistribution.distribution
                            Case "Sprinkler Reliability"
                                oDistribution.units = "-"
                            Case "Sprinkler Suppression Probability"
                                oDistribution.units = "-"
                            Case "Sprinkler Cooling Coefficient"
                                oDistribution.units = "-"
                            Case "Fuel Heat of Gasification"
                                oDistribution.units = "kJ/g"
                            Case "Alpha T"
                                oDistribution.units = "kW/s2"
                            Case "Peak HRR"
                                oDistribution.units = "kW"
                            Case "Smoke Detector Reliability"
                                oDistribution.units = "-"
                            Case "Mechanical Ventilation Reliability"
                                oDistribution.units = "-"
                        End Select

                    End If
                Next

                If oDistribution IsNot Nothing Then
                    DistributionClass.SaveDistributions(oDistributions)
                End If

                Select Case Me.Tag
                    Case "Interior Temperature"
                        frmOptions1.txtInteriorTemp.Text = txtValue.Text
                        If paramdist <> "None" Then
                            frmOptions1.txtInteriorTemp.BackColor = distbackcolor
                        Else
                            frmOptions1.txtInteriorTemp.BackColor = distnobackcolor
                        End If
                    Case "Exterior Temperature"
                        frmOptions1.txtExteriorTemp.Text = txtValue.Text
                        If paramdist <> "None" Then
                            frmOptions1.txtExteriorTemp.BackColor = distbackcolor
                        Else
                            frmOptions1.txtExteriorTemp.BackColor = distnobackcolor
                        End If
                    Case "Relative Humidity"
                        frmOptions1.txtRelativeHumidity.Text = txtValue.Text
                        If paramdist <> "None" Then
                            frmOptions1.txtRelativeHumidity.BackColor = distbackcolor
                        Else
                            frmOptions1.txtRelativeHumidity.BackColor = distnobackcolor
                        End If
                    Case "Fire Load Energy Density"
                        frmItemList.txtFLED.Text = txtValue.Text
                        If paramdist <> "None" Then
                            frmItemList.txtFLED.BackColor = distbackcolor
                        Else
                            frmItemList.txtFLED.BackColor = distnobackcolor
                        End If
                    Case "Heat of Combustion PFO"
                        frmOptions1.txtHOCFuel.Text = txtValue.Text
                        If paramdist <> "None" Then
                            frmOptions1.txtHOCFuel.BackColor = distbackcolor
                        Else
                            frmOptions1.txtHOCFuel.BackColor = distnobackcolor
                        End If
                    Case "Soot Preflashover Yield"
                        frmOptions1.txtpreSoot.Text = txtValue.Text
                        If paramdist <> "None" Then
                            frmOptions1.txtpreSoot.BackColor = distbackcolor
                        Else
                            frmOptions1.txtpreSoot.BackColor = distnobackcolor
                        End If
                    Case "CO Preflashover Yield"
                        frmOptions1.txtpreCO.Text = txtValue.Text
                        If paramdist <> "None" Then
                            frmOptions1.txtpreCO.BackColor = distbackcolor
                        Else
                            frmOptions1.txtpreCO.BackColor = distnobackcolor
                        End If
                    Case "Sprinkler Reliability"
                        frmSprinklerList.txtSprReliability.Text = txtValue.Text
                        If paramdist <> "None" Then
                            frmSprinklerList.txtSprReliability.BackColor = distbackcolor
                        Else
                            frmSprinklerList.txtSprReliability.BackColor = distnobackcolor
                        End If
                    Case "Sprinkler Suppression Probability"
                        frmSprinklerList.txtSprSuppressProb.Text = txtValue.Text
                        If paramdist <> "None" Then
                            frmSprinklerList.txtSprSuppressProb.BackColor = distbackcolor
                        Else
                            frmSprinklerList.txtSprSuppressProb.BackColor = distnobackcolor
                        End If
                    Case "Sprinkler Cooling Coefficient"
                        frmSprinklerList.txtSprCoolingCoeff.Text = txtValue.Text
                        If paramdist <> "None" Then
                            frmSprinklerList.txtSprCoolingCoeff.BackColor = distbackcolor
                        Else
                            frmSprinklerList.txtSprCoolingCoeff.BackColor = distnobackcolor
                        End If
                    Case "Alpha T"
                        frmPowerlaw.TxtAlphaT.Text = txtValue.Text
                        If paramdist <> "None" Then
                            frmPowerlaw.TxtAlphaT.BackColor = distbackcolor
                        Else
                            frmPowerlaw.TxtAlphaT.BackColor = distnobackcolor
                        End If
                    Case "Peak HRR"
                        frmPowerlaw.txtPeakHRR.Text = txtValue.Text
                        If paramdist <> "None" Then
                            frmPowerlaw.txtPeakHRR.BackColor = distbackcolor
                        Else
                            frmPowerlaw.txtPeakHRR.BackColor = distnobackcolor
                        End If
                    Case "Smoke Detector Reliability"
                        frmSmokeDetList.txtSmokeDetReliability.Text = txtValue.Text
                        If paramdist <> "None" Then
                            frmSmokeDetList.txtSmokeDetReliability.BackColor = distbackcolor
                        Else
                            frmSmokeDetList.txtSmokeDetReliability.BackColor = distnobackcolor
                        End If
                    Case "Mechanical Ventilation Reliability"
                        frmFanList.txtFanReliability.Text = txtValue.Text
                        If paramdist <> "None" Then
                            frmFanList.txtFanReliability.BackColor = distbackcolor
                        Else
                            frmFanList.txtFanReliability.BackColor = distnobackcolor
                        End If
                End Select

            ElseIf Me.Text = "Edit Item Distribution" Then
                itemid = frmItemList.lstItems.SelectedIndex
                oItems = ItemDB.GetItemsv2()
                oDistributions = ItemDB.GetItemDistributions()

                For Each Me.oDistribution In oDistributions
                    If oDistribution.varname = Me.Tag And oDistribution.id = itemid + 1 Then
                        oDistribution.varvalue = txtValue.Text
                        oDistribution.distribution = lstDistribution.Text
                        oDistribution.lbound = txtLowerBound.Text
                        oDistribution.ubound = txtUpperBound.Text
                        oDistribution.mean = txtMean.Text
                        oDistribution.variance = txtVariance.Text
                        oDistribution.mode = txtMode.Text
                        oDistribution.alpha = txtAlpha.Text
                        oDistribution.beta = txtBeta.Text

                        Select Case Me.Tag
                            Case "heat of combustion"
                                oItems(itemid).hoc = txtValue.Text
                            Case "soot"
                                oItems(itemid).soot = txtValue.Text
                            Case "co2"
                                oItems(itemid).co2 = txtValue.Text
                            Case "Latent Heat of Gasification"
                                oItems(itemid).LHoG = txtValue.Text
                            Case "Radiant Loss Fraction"
                                oItems(itemid).radlossfrac = txtValue.Text
                            Case "HRRUA"
                                oItems(itemid).HRRUA = txtValue.Text
                        End Select
                    End If
                Next

                ItemDB.SaveItemsv2(oItems, oDistributions)

            ElseIf Me.Text = "Edit Vent Distribution" Then
                ' ventid = frmVentList.lstVents.SelectedIndex

                ovents = VentDB.GetVents
                If ventid = -1 Then ventid = ovents.Count - 1
                oDistributions = VentDB.GetVentDistributions()

                For Each Me.oDistribution In oDistributions
                    If oDistribution.varname = Me.Tag And oDistribution.id = ventid + 1 Then
                        oDistribution.varvalue = txtValue.Text
                        oDistribution.distribution = lstDistribution.Text
                        oDistribution.lbound = txtLowerBound.Text
                        oDistribution.ubound = txtUpperBound.Text
                        oDistribution.mean = txtMean.Text
                        oDistribution.variance = txtVariance.Text
                        oDistribution.mode = txtMode.Text
                        oDistribution.alpha = txtAlpha.Text
                        oDistribution.beta = txtBeta.Text

                        Select Case Me.Tag
                            Case "width"
                                ovents(ventid).width = txtValue.Text
                                If lstDistribution.Text <> "None" Then
                                    frmNewVents.txtVentWidth.BackColor = distbackcolor
                                Else
                                    frmNewVents.txtVentWidth.BackColor = distnobackcolor
                                End If
                            Case "height"
                                ovents(ventid).height = txtValue.Text
                                If lstDistribution.Text <> "None" Then
                                    frmNewVents.txtVentHeight.BackColor = distbackcolor
                                Else
                                    frmNewVents.txtVentHeight.BackColor = distnobackcolor
                                End If
                            Case "prob"
                                ovents(ventid).probventclosed = txtValue.Text
                                If lstDistribution.Text <> "None" Then
                                    frmNewVents.txtVentProb.BackColor = distbackcolor
                                Else
                                    frmNewVents.txtVentProb.BackColor = distnobackcolor
                                End If
                            Case "HOreliability"
                                ovents(ventid).horeliability = txtValue.Text
                                If lstDistribution.Text <> "None" Then
                                    frmNewVents.txtHOreliability.BackColor = distbackcolor
                                Else
                                    frmNewVents.txtHOreliability.BackColor = distnobackcolor
                                End If
                            Case "integrity"
                                ovents(ventid).integrity = txtValue.Text
                                If lstDistribution.Text <> "None" Then
                                    frmFireResistance.txtFRintegrity.BackColor = distbackcolor
                                Else
                                    frmFireResistance.txtFRintegrity.BackColor = distnobackcolor
                                End If
                            Case "maxopening"
                                ovents(ventid).maxopening = txtValue.Text
                                If lstDistribution.Text <> "None" Then
                                    frmFireResistance.txtMaxOpening.BackColor = distbackcolor
                                Else
                                    frmFireResistance.txtMaxOpening.BackColor = distnobackcolor
                                End If
                            Case "maxopeningtime"
                                ovents(ventid).maxopeningtime = txtValue.Text
                                If lstDistribution.Text <> "None" Then
                                    frmFireResistance.txtMaxOpeningTime.BackColor = distbackcolor
                                Else
                                    frmFireResistance.txtMaxOpeningTime.BackColor = distnobackcolor
                                End If
                            Case "gastemp"
                                ovents(ventid).gastemp = txtValue.Text
                                If lstDistribution.Text <> "None" Then
                                    frmFireResistance.txtFRgastemp.BackColor = distbackcolor
                                Else
                                    frmFireResistance.txtFRgastemp.BackColor = distnobackcolor
                                End If
                        End Select
                    End If
                Next

                VentDB.SaveVents(ovents, oDistributions)

            ElseIf Me.Text = "Edit Ceiling Vent Distribution" Then

                ocvents = VentDB.GetCVents
                If ventid = -1 Then ventid = ocvents.Count - 1
                oDistributions = VentDB.GetCVentDistributions()

                For Each Me.oDistribution In oDistributions
                    If oDistribution.varname = Me.Tag And oDistribution.id = ventid + 1 Then
                        oDistribution.varvalue = txtValue.Text
                        oDistribution.distribution = lstDistribution.Text
                        oDistribution.lbound = txtLowerBound.Text
                        oDistribution.ubound = txtUpperBound.Text
                        oDistribution.mean = txtMean.Text
                        oDistribution.variance = txtVariance.Text
                        oDistribution.mode = txtMode.Text
                        oDistribution.alpha = txtAlpha.Text
                        oDistribution.beta = txtBeta.Text

                        Select Case Me.Tag
                            Case "area"
                                ocvents(ventid).area = txtValue.Text
                                If lstDistribution.Text <> "None" Then
                                    frmNewCVents.txtCVentArea.BackColor = distbackcolor
                                Else
                                    frmNewCVents.txtCVentArea.BackColor = distnobackcolor
                                End If
                            Case "integrity"
                                ocvents(ventid).integrity = txtValue.Text
                                If lstDistribution.Text <> "None" Then
                                    frmFireResistance.txtFRintegrity.BackColor = distbackcolor
                                Else
                                    frmFireResistance.txtFRintegrity.BackColor = distnobackcolor
                                End If
                            Case "maxopening"
                                ocvents(ventid).maxopening = txtValue.Text
                                If lstDistribution.Text <> "None" Then
                                    frmFireResistance.txtMaxOpening.BackColor = distbackcolor
                                Else
                                    frmFireResistance.txtMaxOpening.BackColor = distnobackcolor
                                End If
                            Case "maxopeningtime"
                                ocvents(ventid).maxopeningtime = txtValue.Text
                                If lstDistribution.Text <> "None" Then
                                    frmFireResistance.txtMaxOpeningTime.BackColor = distbackcolor
                                Else
                                    frmFireResistance.txtMaxOpeningTime.BackColor = distnobackcolor
                                End If
                            Case "gastemp"
                                ocvents(ventid).gastemp = txtValue.Text
                                If lstDistribution.Text <> "None" Then
                                    frmFireResistance.txtFRgastemp.BackColor = distbackcolor
                                Else
                                    frmFireResistance.txtFRgastemp.BackColor = distnobackcolor
                                End If
                        End Select
                    End If
                Next

                VentDB.SaveCVents(ocvents, oDistributions)

            ElseIf Me.Text = "Edit Sprinkler Distribution" Then
                sprid = frmSprinklerList.lstSprinklers.SelectedIndex
                oSprinklers = SprinklerDB.GetSprinklers2
                oDistributions = SprinklerDB.GetSprDistributions

                For Each Me.oDistribution In oDistributions
                    If oDistribution.varname = Me.Tag And oDistribution.id = sprid + 1 Then
                        oDistribution.varvalue = txtValue.Text
                        oDistribution.distribution = lstDistribution.Text
                        oDistribution.lbound = txtLowerBound.Text
                        oDistribution.ubound = txtUpperBound.Text
                        oDistribution.mean = txtMean.Text
                        oDistribution.variance = txtVariance.Text
                        oDistribution.mode = txtMode.Text
                        oDistribution.alpha = txtAlpha.Text
                        oDistribution.beta = txtBeta.Text

                    End If
                Next

                SprinklerDB.SaveSprinklers2(oSprinklers, oDistributions)

            ElseIf Me.Text = "Edit Smoke Detector Distribution" Then
                sdid = frmSmokeDetList.lstSmokeDet.SelectedIndex
                oSmokeDets = SmokeDetDB.GetSmokDets
                oDistributions = SmokeDetDB.GetSDDistributions

                For Each Me.oDistribution In oDistributions
                    If oDistribution.varname = Me.Tag And oDistribution.id = sdid + 1 Then
                        oDistribution.varvalue = txtValue.Text
                        oDistribution.distribution = lstDistribution.Text
                        oDistribution.lbound = txtLowerBound.Text
                        oDistribution.ubound = txtUpperBound.Text
                        oDistribution.mean = txtMean.Text
                        oDistribution.variance = txtVariance.Text
                        oDistribution.mode = txtMode.Text
                        oDistribution.alpha = txtAlpha.Text
                        oDistribution.beta = txtBeta.Text

                    End If
                Next

                SmokeDetDB.SaveSmokeDets(oSmokeDets, oDistributions)

            ElseIf Me.Text = "Edit Fan Distribution" Then
                fanid = frmFanList.lstFan.SelectedIndex
                oFans = FanDB.GetFans
                oDistributions = FanDB.GetFanDistributions

                For Each Me.oDistribution In oDistributions
                    If oDistribution.varname = Me.Tag And oDistribution.id = fanid + 1 Then
                        oDistribution.varvalue = txtValue.Text
                        oDistribution.distribution = lstDistribution.Text
                        oDistribution.lbound = txtLowerBound.Text
                        oDistribution.ubound = txtUpperBound.Text
                        oDistribution.mean = txtMean.Text
                        oDistribution.variance = txtVariance.Text
                        oDistribution.mode = txtMode.Text
                        oDistribution.alpha = txtAlpha.Text
                        oDistribution.beta = txtBeta.Text

                    End If
                Next

                FanDB.SaveFans(oFans, oDistributions)

            ElseIf Me.Text = "Edit Room Distribution" Then
                Dim roomid As Integer = 0
                roomid = frmRoomList.lstRooms.SelectedIndex
                oRooms = RoomDB.GetRooms
                oDistributions = RoomDB.GetRoomDistributions

                For Each Me.oDistribution In oDistributions
                    If oDistribution.varname = Me.Tag And oDistribution.id = roomid + 1 Then
                        oDistribution.varvalue = txtValue.Text
                        oDistribution.distribution = lstDistribution.Text
                        oDistribution.lbound = txtLowerBound.Text
                        oDistribution.ubound = txtUpperBound.Text
                        oDistribution.mean = txtMean.Text
                        oDistribution.variance = txtVariance.Text
                        oDistribution.mode = txtMode.Text
                        oDistribution.alpha = txtAlpha.Text
                        oDistribution.beta = txtBeta.Text

                    End If
                Next

                RoomDB.SaveRooms(oRooms, oDistributions)
            End If

            Me.Close()

        End If
    End Sub
    Private Function IsValidData() As Boolean
        Return Validator.IsPresent(txtValue) AndAlso _
               Validator.IsPresent(txtVariance) AndAlso _
               Validator.IsPresent(txtMean) AndAlso _
               Validator.IsPresent(txtLowerBound) AndAlso _
                Validator.IsPresent(txtAlpha) AndAlso _
                Validator.IsPresent(txtBeta) AndAlso _
                Validator.IsPresent(txtMode) AndAlso _
               Validator.IsPresent(txtUpperBound)
    End Function

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub lstDistribution_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstDistribution.SelectedIndexChanged

        If lstDistribution.Text = "None" Then

            txtValue.Enabled = True
            txtLowerBound.Enabled = False
            txtUpperBound.Enabled = False
            txtMean.Enabled = False
            txtVariance.Enabled = False
            txtMode.Enabled = False
            txtAlpha.Enabled = False
            txtBeta.Enabled = False

        ElseIf lstDistribution.Text = "Normal" Then
           
            txtValue.Enabled = False
            txtLowerBound.Enabled = True
            txtUpperBound.Enabled = True
            txtMean.Enabled = True
            txtVariance.Enabled = True
            txtMode.Enabled = False
            txtAlpha.Enabled = False
            txtBeta.Enabled = False

        ElseIf lstDistribution.Text = "Log Normal" Then

            txtValue.Enabled = False
            txtLowerBound.Enabled = False
            txtUpperBound.Enabled = False
            txtMean.Enabled = True
            txtVariance.Enabled = True
            txtMode.Enabled = False
            txtAlpha.Enabled = False
            txtBeta.Enabled = False

        ElseIf lstDistribution.Text = "Uniform" Then

            txtValue.Enabled = False
            txtLowerBound.Enabled = True
            txtUpperBound.Enabled = True
            txtMean.Enabled = False
            txtVariance.Enabled = False
            txtMode.Enabled = False
            txtAlpha.Enabled = False
            txtBeta.Enabled = False

        ElseIf lstDistribution.Text = "Triangular" Then

            txtValue.Enabled = False
            txtLowerBound.Enabled = True
            txtUpperBound.Enabled = True
            txtMean.Enabled = False
            txtVariance.Enabled = False
            txtMode.Enabled = True
            txtAlpha.Enabled = False
            txtBeta.Enabled = False

        ElseIf lstDistribution.Text = "Gamma" Then

            txtValue.Enabled = False
            txtLowerBound.Enabled = False
            txtUpperBound.Enabled = False
            txtMean.Enabled = False
            txtVariance.Enabled = False
            txtMode.Enabled = False
            txtAlpha.Enabled = True
            txtBeta.Enabled = True

        End If
    End Sub
    Public Sub ShowItemDistributionForm(ByVal param As String, ByVal units As String, ByVal instruction As String, ByRef oitem As Object, ByRef oitemdistributions As Object, ByRef itemid As Integer, ByRef paramdist As String)

        Dim NewDistributionForm As New frmDistParam
        NewDistributionForm.Text = "Edit Item Distribution"
        NewDistributionForm.LblInstruction.Text = instruction

        Try
            NewDistributionForm.lblDistName.Text = param.ToString & " (" & units.ToString & ")"

            For Each oitemdistribution In oitemdistributions
                If oitemdistribution.varname = param And oitemdistribution.id = itemid + 1 Then
                    NewDistributionForm.Tag = param
                    NewDistributionForm.lstDistribution.Text = oitemdistribution.distribution
                    NewDistributionForm.txtUpperBound.Text = oitemdistribution.ubound
                    NewDistributionForm.txtLowerBound.Text = oitemdistribution.lbound
                    NewDistributionForm.txtMean.Text = oitemdistribution.mean
                    NewDistributionForm.txtValue.Text = oitemdistribution.varvalue
                    NewDistributionForm.txtVariance.Text = oitemdistribution.variance
                    NewDistributionForm.txtMode.Text = oitemdistribution.mode
                    NewDistributionForm.txtAlpha.Text = oitemdistribution.alpha
                    NewDistributionForm.txtBeta.Text = oitemdistribution.beta

                    Exit For
                End If
            Next

            NewDistributionForm.ShowDialog()

            Dim oitems As List(Of oItem)
            oitems = ItemDB.GetItemsv2()
            oDistributions = ItemDB.GetItemDistributions()

            For Each Me.oDistribution In oDistributions
                If oDistribution.id = itemid + 1 Then

                    Select Case oDistribution.varname
                        Case "heat of combustion"
                            oitem.hoc = oDistribution.varvalue

                        Case "soot yield"
                            oitem.soot = oDistribution.varvalue

                        Case "co2 yield"
                            oitem.co2 = oDistribution.varvalue

                        Case "Latent Heat of Gasification"
                            oitem.lhog = oDistribution.varvalue

                        Case "Radiant Loss Fraction"
                            oitem.radlossfrac = oDistribution.varvalue

                        Case "HRRUA"
                            oitem.hrrua = oDistribution.varvalue

                    End Select


                End If

                If oDistribution.varname = param And oDistribution.id = itemid + 1 Then
                    paramdist = oDistribution.distribution
                End If

            Next

            ItemDB.SaveItemsv2(oitems, oDistributions)

            oitemdistributions = oDistributions

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "frmDistParam.ShowDistributionForm")
        End Try
    End Sub
    Public Sub ShowVentDistributionForm(ByVal param As String, ByVal units As String, ByVal instruction As String, ByRef ovent As Object, ByRef oventdistributions As Object, ByRef id As Integer, ByRef paramdist As String)
        ventid = id
        Dim NewDistributionForm As New frmDistParam
        NewDistributionForm.Text = "Edit Vent Distribution"
        NewDistributionForm.LblInstruction.Text = instruction

        Try
            NewDistributionForm.lblDistName.Text = param.ToString & " (" & units.ToString & ")"

            For Each oventdistribution In oventdistributions
                If oventdistribution.varname = param And oventdistribution.id = ventid + 1 Then
                    NewDistributionForm.Tag = param
                    NewDistributionForm.lstDistribution.SelectedItem = oventdistribution.distribution
                    NewDistributionForm.txtUpperBound.Text = oventdistribution.ubound
                    NewDistributionForm.txtLowerBound.Text = oventdistribution.lbound
                    NewDistributionForm.txtMean.Text = oventdistribution.mean
                    NewDistributionForm.txtValue.Text = oventdistribution.varvalue
                    NewDistributionForm.txtVariance.Text = oventdistribution.variance
                    NewDistributionForm.txtMode.Text = oventdistribution.mode
                    NewDistributionForm.txtAlpha.Text = oventdistribution.alpha
                    NewDistributionForm.txtBeta.Text = oventdistribution.beta

                    Exit For
                End If
            Next

            NewDistributionForm.ShowDialog()

            Dim ovents As List(Of oVent)
            ovents = VentDB.GetVents
            oDistributions = VentDB.GetVentDistributions()

            For Each Me.oDistribution In oDistributions
                If oDistribution.id = ventid + 1 Then

                    Select Case oDistribution.varname
                        Case "width"
                            ovent.width = oDistribution.varvalue
                        Case "height"
                            ovent.height = oDistribution.varvalue
                        Case "prob"
                            ovent.probventclosed = oDistribution.varvalue
                        Case "HOreliability"
                            ovent.horeliability = oDistribution.varvalue
                        Case "integrity"
                            ovent.integrity = oDistribution.varvalue
                        Case "maxopening"
                            ovent.maxopening = oDistribution.varvalue
                        Case "maxopeningtime"
                            ovent.maxopeningtime = oDistribution.varvalue
                        Case "gastemp"
                            ovent.gastemp = oDistribution.varvalue
                    End Select

                End If

                If oDistribution.varname = param And oDistribution.id = ventid + 1 Then
                    paramdist = oDistribution.distribution
                End If
            Next

            VentDB.SaveVents(ovents, oDistributions)
            oventdistributions = oDistributions

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "frmDistParam.ShowVentDistributionForm")
        End Try
    End Sub
    Public Sub ShowCVentDistributionForm(ByVal param As String, ByVal units As String, ByVal instruction As String, ByRef ocvent As Object, ByRef ocventdistributions As Object, ByRef id As Integer, ByRef paramdist As String)
        ventid = id
        Dim NewDistributionForm As New frmDistParam
        NewDistributionForm.Text = "Edit Ceiling Vent Distribution"
        NewDistributionForm.LblInstruction.Text = instruction

        Try
            NewDistributionForm.lblDistName.Text = param.ToString & " (" & units.ToString & ")"

            For Each ocventdistribution In ocventdistributions
                If ocventdistribution.varname = param And ocventdistribution.id = ventid + 1 Then
                    NewDistributionForm.Tag = param
                    NewDistributionForm.lstDistribution.SelectedItem = ocventdistribution.distribution
                    NewDistributionForm.txtUpperBound.Text = ocventdistribution.ubound
                    NewDistributionForm.txtLowerBound.Text = ocventdistribution.lbound
                    NewDistributionForm.txtMean.Text = ocventdistribution.mean
                    NewDistributionForm.txtValue.Text = ocventdistribution.varvalue
                    NewDistributionForm.txtVariance.Text = ocventdistribution.variance
                    NewDistributionForm.txtMode.Text = ocventdistribution.mode
                    NewDistributionForm.txtAlpha.Text = ocventdistribution.alpha
                    NewDistributionForm.txtBeta.Text = ocventdistribution.beta

                    Exit For
                End If
            Next

            NewDistributionForm.ShowDialog()

            Dim ocvents As List(Of oCVent)
            ocvents = VentDB.GetCVents
            oDistributions = VentDB.GetCVentDistributions()

            For Each Me.oDistribution In oDistributions
                If oDistribution.id = ventid + 1 Then

                    Select Case oDistribution.varname
                        Case "area"
                            ocvent.area = oDistribution.varvalue
                        Case "integrity"
                            ocvent.integrity = oDistribution.varvalue
                        Case "maxopening"
                            ocvent.maxopening = oDistribution.varvalue
                        Case "maxopeningtime"
                            ocvent.maxopeningtime = oDistribution.varvalue
                        Case "gastemp"
                            ocvent.gastemp = oDistribution.varvalue
                    End Select

                End If

                If oDistribution.varname = param And oDistribution.id = ventid + 1 Then
                    paramdist = oDistribution.distribution
                End If
            Next

            VentDB.SaveCVents(ocvents, oDistributions)
            ocventdistributions = oDistributions

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "frmDistParam.ShowCVentDistributionForm")
        End Try
    End Sub

    Public Sub ShowDistributionForm(ByVal param As String, ByVal units As String, ByVal instruction As String)

        Dim NewDistributionForm As New frmDistParam
        NewDistributionForm.Text = "Edit Distribution"
        NewDistributionForm.LblInstruction.Text = instruction
        Dim odistributions As List(Of oDistribution)

        Try

            odistributions = DistributionClass.GetDistributions()

            If odistributions.Count = 0 Then
                MsgBox("no distribution data")
                Exit Sub
            End If

            For Each Me.oDistribution In odistributions
                If oDistribution.varname = param Then
                    NewDistributionForm.oDistribution = oDistribution
                    NewDistributionForm.lblDistName.Text = param.ToString & " (" & units.ToString & ")"
                    NewDistributionForm.Tag = oDistribution.varname
                    NewDistributionForm.lstDistribution.Text = oDistribution.distribution

                    Select Case param
                        Case "Relative Humidity"
                            NewDistributionForm.txtUpperBound.Text = oDistribution.ubound * 100
                            NewDistributionForm.txtLowerBound.Text = oDistribution.lbound * 100
                            NewDistributionForm.txtMean.Text = oDistribution.mean * 100
                            NewDistributionForm.txtValue.Text = oDistribution.varvalue * 100
                            NewDistributionForm.txtVariance.Text = oDistribution.variance * 10000
                            NewDistributionForm.txtMode.Text = oDistribution.mode * 100
                            NewDistributionForm.txtAlpha.Text = oDistribution.alpha
                            NewDistributionForm.txtBeta.Text = oDistribution.beta
                        Case "Interior Temperature"
                            NewDistributionForm.txtUpperBound.Text = oDistribution.ubound - 273
                            NewDistributionForm.txtLowerBound.Text = oDistribution.lbound - 273
                            NewDistributionForm.txtMean.Text = oDistribution.mean - 273
                            NewDistributionForm.txtValue.Text = oDistribution.varvalue - 273
                            NewDistributionForm.txtVariance.Text = oDistribution.variance
                            NewDistributionForm.txtMode.Text = oDistribution.mode - 273
                            NewDistributionForm.txtAlpha.Text = oDistribution.alpha
                            NewDistributionForm.txtBeta.Text = oDistribution.beta
                        Case "Exterior Temperature"
                            NewDistributionForm.txtUpperBound.Text = oDistribution.ubound - 273
                            NewDistributionForm.txtLowerBound.Text = oDistribution.lbound - 273
                            NewDistributionForm.txtMean.Text = oDistribution.mean - 273
                            NewDistributionForm.txtValue.Text = oDistribution.varvalue - 273
                            NewDistributionForm.txtVariance.Text = oDistribution.variance
                            NewDistributionForm.txtMode.Text = oDistribution.mode - 273
                            NewDistributionForm.txtAlpha.Text = oDistribution.alpha
                            NewDistributionForm.txtBeta.Text = oDistribution.beta
                       
                        Case Else
                            NewDistributionForm.txtUpperBound.Text = oDistribution.ubound
                            NewDistributionForm.txtLowerBound.Text = oDistribution.lbound
                            NewDistributionForm.txtMean.Text = oDistribution.mean
                            NewDistributionForm.txtValue.Text = oDistribution.varvalue
                            NewDistributionForm.txtVariance.Text = oDistribution.variance
                            NewDistributionForm.txtMode.Text = oDistribution.mode
                            NewDistributionForm.txtAlpha.Text = oDistribution.alpha
                            NewDistributionForm.txtBeta.Text = oDistribution.beta
                    End Select

                 
                End If
            Next

            NewDistributionForm.ShowDialog()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "frmDistParam.ShowDistributionForm")
        End Try
    End Sub
    Public Sub ShowSprinklerDistributionForm(ByVal param As String, ByVal units As String, ByVal instruction As String, ByVal osprinklers As Object, ByRef odistributions As Object, ByRef sprid As Integer, ByRef paramdist As String)

        Dim NewDistributionForm As New frmDistParam
        NewDistributionForm.Text = "Edit Sprinkler Distribution"
        NewDistributionForm.LblInstruction.Text = instruction

        Try

            If Not IsNothing(odistributions) Then
                For Each Me.oDistribution In odistributions
                    If oDistribution.varname = param And oDistribution.id = sprid + 1 Then
                        NewDistributionForm.oDistribution = oDistribution
                        NewDistributionForm.lblDistName.Text = param.ToString & " (" & units.ToString & ")"
                        NewDistributionForm.Tag = oDistribution.varname
                        NewDistributionForm.lstDistribution.Text = oDistribution.distribution

                        NewDistributionForm.txtUpperBound.Text = oDistribution.ubound
                        NewDistributionForm.txtLowerBound.Text = oDistribution.lbound
                        NewDistributionForm.txtMean.Text = oDistribution.mean
                        NewDistributionForm.txtValue.Text = oDistribution.varvalue
                        NewDistributionForm.txtVariance.Text = oDistribution.variance
                        NewDistributionForm.txtMode.Text = oDistribution.mode
                        NewDistributionForm.txtAlpha.Text = oDistribution.alpha
                        NewDistributionForm.txtBeta.Text = oDistribution.beta

                    End If
                Next
            End If
            NewDistributionForm.ShowDialog()

            osprinklers = SprinklerDB.GetSprinklers2
            odistributions = SprinklerDB.GetSprDistributions

            For Each x In odistributions
                If x.varname = param And x.id = sprid + 1 Then
                    Select Case param
                        Case "rti"
                            paramdist = x.distribution
                            Exit For
                        Case "acttemp"
                            paramdist = x.distribution
                            Exit For
                        Case "cfactor"
                            paramdist = x.distribution
                            Exit For
                        Case "sprdensity"
                            paramdist = x.distribution
                            Exit For
                        Case "sprz"
                            paramdist = x.distribution
                            Exit For
                        Case "sprr"
                            paramdist = x.distribution
                            Exit For
                    End Select
                End If
            Next


        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "frmDistParam.ShowSprinklerDistributionForm")
        End Try
    End Sub
    Public Sub ShowSmokeDetDistributionForm(ByVal param As String, ByVal units As String, ByVal instruction As String, ByVal osmokedets As Object, ByRef odistributions As Object, ByRef sdid As Integer, ByRef paramdist As String)

        Dim NewDistributionForm As New frmDistParam
        NewDistributionForm.Text = "Edit Smoke Detector Distribution"
        NewDistributionForm.LblInstruction.Text = instruction

        Try

            If Not IsNothing(odistributions) Then
                For Each Me.oDistribution In odistributions
                    If oDistribution.varname = param And oDistribution.id = sdid + 1 Then
                        NewDistributionForm.oDistribution = oDistribution
                        NewDistributionForm.lblDistName.Text = param.ToString & " (" & units.ToString & ")"
                        NewDistributionForm.Tag = oDistribution.varname
                        NewDistributionForm.lstDistribution.Text = oDistribution.distribution

                        NewDistributionForm.txtUpperBound.Text = oDistribution.ubound
                        NewDistributionForm.txtLowerBound.Text = oDistribution.lbound
                        NewDistributionForm.txtMean.Text = oDistribution.mean
                        NewDistributionForm.txtValue.Text = oDistribution.varvalue
                        NewDistributionForm.txtVariance.Text = oDistribution.variance
                        NewDistributionForm.txtMode.Text = oDistribution.mode
                        NewDistributionForm.txtAlpha.Text = oDistribution.alpha
                        NewDistributionForm.txtBeta.Text = oDistribution.beta

                    End If
                Next
            End If
            NewDistributionForm.ShowDialog()

            osmokedets = SmokeDetDB.GetSmokDets
            odistributions = SmokeDetDB.GetSDDistributions

            For Each x In odistributions
                If x.varname = param And x.id = sdid + 1 Then
                    Select Case param
                        Case "OD"
                            paramdist = x.distribution
                            Exit For
                        Case "sdz"
                            paramdist = x.distribution
                            Exit For
                        Case "sdr"
                            paramdist = x.distribution
                            Exit For
                        Case "charlength"
                            paramdist = x.distribution
                            Exit For
                    End Select
                End If
            Next

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "frmDistParam.ShowSmokeDetDistributionForm")
        End Try
    End Sub
    Public Sub ShowFanDistributionForm(ByVal param As String, ByVal units As String, ByVal instruction As String, ByVal ofans As Object, ByRef odistributions As Object, ByRef id As Integer, ByRef paramdist As String)

        Dim NewDistributionForm As New frmDistParam
        NewDistributionForm.Text = "Edit Fan Distribution"
        NewDistributionForm.LblInstruction.Text = instruction

        Try

            If Not IsNothing(odistributions) Then
                For Each Me.oDistribution In odistributions
                    If oDistribution.varname = param And oDistribution.id = id Then
                        NewDistributionForm.oDistribution = oDistribution
                        NewDistributionForm.lblDistName.Text = param.ToString & " (" & units.ToString & ")"
                        NewDistributionForm.Tag = oDistribution.varname
                        NewDistributionForm.lstDistribution.Text = oDistribution.distribution

                        NewDistributionForm.txtUpperBound.Text = oDistribution.ubound
                        NewDistributionForm.txtLowerBound.Text = oDistribution.lbound
                        NewDistributionForm.txtMean.Text = oDistribution.mean
                        NewDistributionForm.txtValue.Text = oDistribution.varvalue
                        NewDistributionForm.txtVariance.Text = oDistribution.variance
                        NewDistributionForm.txtMode.Text = oDistribution.mode
                        NewDistributionForm.txtAlpha.Text = oDistribution.alpha
                        NewDistributionForm.txtBeta.Text = oDistribution.beta

                    End If
                Next
            End If
            NewDistributionForm.ShowDialog()

            ofans = FanDB.GetFans
            odistributions = FanDB.GetFanDistributions

            For Each x In odistributions
                If x.varname = param And x.id = id Then
                    Select Case param
                        Case "fanflowrate"
                            paramdist = x.distribution
                            Exit For
                        Case "fanstarttime"
                            paramdist = x.distribution
                            Exit For
                        Case "fanpressurelimit"
                            paramdist = x.distribution
                            Exit For
                        Case "fanreliability"
                            paramdist = x.distribution
                            Exit For
                    End Select
                End If
            Next

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "frmDistParam.ShowFanDistributionForm")
        End Try
    End Sub
    Public Sub ShowRoomDistributionForm(ByVal param As String, ByVal units As String, ByVal instruction As String, ByVal ofans As Object, ByRef odistributions As Object, ByRef id As Integer, ByRef paramdist As String)

        Dim NewDistributionForm As New frmDistParam
        NewDistributionForm.Text = "Edit Room Distribution"
        NewDistributionForm.LblInstruction.Text = instruction

        Try

            If Not IsNothing(odistributions) Then
                For Each Me.oDistribution In odistributions
                    If oDistribution.varname = param And oDistribution.id = id Then
                        NewDistributionForm.oDistribution = oDistribution
                        NewDistributionForm.lblDistName.Text = param.ToString & " (" & units.ToString & ")"
                        NewDistributionForm.Tag = oDistribution.varname
                        NewDistributionForm.lstDistribution.Text = oDistribution.distribution

                        NewDistributionForm.txtUpperBound.Text = oDistribution.ubound
                        NewDistributionForm.txtLowerBound.Text = oDistribution.lbound
                        NewDistributionForm.txtMean.Text = oDistribution.mean
                        NewDistributionForm.txtValue.Text = oDistribution.varvalue
                        NewDistributionForm.txtVariance.Text = oDistribution.variance
                        NewDistributionForm.txtMode.Text = oDistribution.mode
                        NewDistributionForm.txtAlpha.Text = oDistribution.alpha
                        NewDistributionForm.txtBeta.Text = oDistribution.beta

                    End If
                Next
            End If
            NewDistributionForm.ShowDialog()

            oRooms = RoomDB.GetRooms
            odistributions = RoomDB.GetRoomDistributions

            For Each x In odistributions
                If x.varname = param And x.id = id Then
                    Select Case param
                        Case "length"
                            paramdist = x.distribution
                            Exit For
                        Case "width"
                            paramdist = x.distribution
                            Exit For
                    End Select
                End If
            Next

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "frmDistParam.ShowRoomDistributionForm")
        End Try
    End Sub



End Class