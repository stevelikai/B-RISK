Imports System.Collections.Generic
Imports System.Data.OleDb
Public Class frmMatSelect
    Inherits System.Windows.Forms.Form
    Private Sub Table1BindingNavigatorSaveItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Validate()
        Me.Table1BindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(Me.ThermalDataSet)
    End Sub

    Private Sub frmMatSelect_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'capture OleDB exception
        Try
            Me.Table1TableAdapter.Fill(Me.ThermalDataSet.Table1)
        Catch ex As OleDbException
            MessageBox.Show("Database error # " & ex.ErrorCode & ": " & ex.GetType.ToString)
        End Try

        Dim room As Integer
        room = CInt(frmNewRoom.txtID.Text)

        Select Case thermal
            Case "wallsurface"
                'room = frmDescribeRoom._lstRoomID_1.SelectedIndex + 1
                Material_DescriptionListBox.Text = WallSurface(room)
            Case "wallsubstrate"
                'room = frmDescribeRoom._lstRoomID_1.SelectedIndex + 1
                Material_DescriptionListBox.Text = WallSubstrate(room)
            Case "ceilingsurface"
                'room = frmDescribeRoom._lstRoomID_2.SelectedIndex + 1
                Material_DescriptionListBox.Text = CeilingSurface(room)
            Case "ceilingsubstrate"
                'room = frmDescribeRoom._lstRoomID_2.SelectedIndex + 1
                Material_DescriptionListBox.Text = CeilingSubstrate(room)
            Case "floorsurface"
                'room = frmDescribeRoom._lstRoomID_3.SelectedIndex + 1
                Material_DescriptionListBox.Text = FloorSurface(room)
            Case "floorsubstrate"
                'room = frmDescribeRoom._lstRoomID_3.SelectedIndex + 1
                Material_DescriptionListBox.Text = FloorSubstrate(room)
        End Select
    End Sub

    '    Private Sub cmdSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelect.Click

    '        Dim room As Integer
    '        On Error GoTo handler

    '        Select Case thermal
    '            Case "wallsurface"
    '                room = frmDescribeRoom._lstRoomID_1.SelectedIndex + 1
    '                WallSurface(room) = Material_DescriptionListBox.Text
    '                frmDescribeRoom.lblWallSurface.Text = WallSurface(room)
    '                WallConductivity(room) = CDbl(Thermal_ConductivityTextBox.Text)
    '                WallSpecificHeat(room) = CDbl(Specific_HeatTextBox.Text)
    '                WallDensity(room) = CDbl(DensityTextBox.Text)
    '                Surface_Emissivity(2, room) = CDbl(EmissivityTextBox.Text)
    '                Surface_Emissivity(3, room) = CDbl(EmissivityTextBox.Text)
    '                WallTSMin(room) = CDbl(Min_Temp_for_SpreadTextBox.Text) + 273
    '                WallEffectiveHeatofCombustion(room) = CDbl(Heat_of_CombustionTextBox.Text)
    '                WallSootYield(room) = CDbl(Soot_YieldTextBox.Text)
    '                WallH2OYield(room) = CDbl(H2O_YieldTextBox.Text)
    '                WallCO2Yield(room) = CDbl(CO2_YieldTextBox.Text)
    '                WallFlameSpreadParameter(room) = CDbl(Flame_Spread_ParameterTextBox.Text)
    '                WallConeDataFile(room) = CStr(Cone_Data_FileTextBox.Text)
    '            Case "wallsubstrate"
    '                room = frmDescribeRoom._lstRoomID_1.SelectedIndex + 1
    '                WallSubstrate(room) = Material_DescriptionListBox.Text
    '                frmDescribeRoom.lblWallSubstrate.Text = WallSubstrate(room)
    '                WallSubConductivity(room) = CDbl(Thermal_ConductivityTextBox.Text)
    '                WallSubSpecificHeat(room) = CDbl(Specific_HeatTextBox.Text)
    '                WallSubDensity(room) = CDbl(DensityTextBox.Text)
    '            Case "ceilingsurface"
    '                room = frmDescribeRoom._lstRoomID_2.SelectedIndex + 1
    '                CeilingSurface(room) = Material_DescriptionListBox.Text
    '                frmDescribeRoom.lblCeilingSurface.Text = CeilingSurface(room)
    '                CeilingConductivity(room) = CDbl(Thermal_ConductivityTextBox.Text)
    '                CeilingSpecificHeat(room) = CDbl(Specific_HeatTextBox.Text)
    '                CeilingDensity(room) = CDbl(DensityTextBox.Text)
    '                Surface_Emissivity(1, room) = CDbl(EmissivityTextBox.Text)
    '                CeilingEffectiveHeatofCombustion(room) = CDbl(Heat_of_CombustionTextBox.Text)
    '                CeilingSootYield(room) = CDbl(Soot_YieldTextBox.Text)
    '                CeilingH2OYield(room) = CDbl(H2O_YieldTextBox.Text)
    '                CeilingCO2Yield(room) = CDbl(CO2_YieldTextBox.Text)
    '                CeilingConeDataFile(room) = CStr(Cone_Data_FileTextBox.Text)
    '            Case "ceilingsubstrate"
    '                room = frmDescribeRoom._lstRoomID_2.SelectedIndex + 1
    '                CeilingSubstrate(room) = Material_DescriptionListBox.Text
    '                frmDescribeRoom.lblCeilingSubstrate.Text = CeilingSubstrate(room)
    '                CeilingSubConductivity(room) = CDbl(Thermal_ConductivityTextBox.Text)
    '                CeilingSubSpecificHeat(room) = CDbl(Specific_HeatTextBox.Text)
    '                CeilingSubDensity(room) = CDbl(DensityTextBox.Text)
    '            Case "floorsurface"
    '                room = frmDescribeRoom._lstRoomID_3.SelectedIndex + 1
    '                FloorSurface(room) = Material_DescriptionListBox.Text
    '                frmDescribeRoom.lblFloorSurface.Text = FloorSurface(room)
    '                FloorConductivity(room) = CDbl(Thermal_ConductivityTextBox.Text)
    '                FloorSpecificHeat(room) = CDbl(Specific_HeatTextBox.Text)
    '                FloorDensity(room) = CDbl(DensityTextBox.Text)
    '                Surface_Emissivity(4, room) = CDbl(EmissivityTextBox.Text)
    '                FloorTSMin(room) = CDbl(Min_Temp_for_SpreadTextBox.Text) + 273
    '                FloorEffectiveHeatofCombustion(room) = CDbl(Heat_of_CombustionTextBox.Text)
    '                FloorSootYield(room) = CDbl(Soot_YieldTextBox.Text)
    '                FloorH2OYield(room) = CDbl(H2O_YieldTextBox.Text)
    '                FloorCO2Yield(room) = CDbl(CO2_YieldTextBox.Text)
    '                FloorFlameSpreadParameter(room) = CDbl(Flame_Spread_ParameterTextBox.Text)
    '                FloorConeDataFile(room) = CStr(Cone_Data_FileTextBox.Text)
    '            Case "floorsubstrate"
    '                room = frmDescribeRoom._lstRoomID_3.SelectedIndex + 1
    '                FloorSubstrate(room) = Material_DescriptionListBox.Text
    '                frmDescribeRoom.lblFloorSubstrate.Text = FloorSubstrate(room)
    '                FloorSubConductivity(room) = CDbl(Thermal_ConductivityTextBox.Text)
    '                FloorSubSpecificHeat(room) = CDbl(Specific_HeatTextBox.Text)
    '                FloorSubDensity(room) = CDbl(DensityTextBox.Text)
    '        End Select

    '        frmDescribeRoom.Show()
    '        Me.Close()

    '        Exit Sub

    'handler:
    '        MsgBox("There was a problem extracting data from the materials database.")

    '    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
        'frmDescribeRoom.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles cmdSelectMat.Click
        Dim room As Integer = CInt(Me.Tag)
        On Error GoTo handler

        Select Case thermal
            Case "wallsurface"
                WallSurface(room) = Material_DescriptionListBox.Text
                WallConductivity(room) = CDbl(Thermal_ConductivityTextBox.Text)
                WallSpecificHeat(room) = CDbl(Specific_HeatTextBox.Text)
                WallDensity(room) = CDbl(DensityTextBox.Text)
                Surface_Emissivity(2, room) = CDbl(EmissivityTextBox.Text)
                Surface_Emissivity(3, room) = CDbl(EmissivityTextBox.Text)
                WallTSMin(room) = CDbl(Min_Temp_for_SpreadTextBox.Text) + 273
                WallEffectiveHeatofCombustion(room) = CDbl(Heat_of_CombustionTextBox.Text)
                WallSootYield(room) = CDbl(Soot_YieldTextBox.Text)
                WallH2OYield(room) = CDbl(H2O_YieldTextBox.Text)
                WallCO2Yield(room) = CDbl(CO2_YieldTextBox.Text)
                WallFlameSpreadParameter(room) = CDbl(Flame_Spread_ParameterTextBox.Text)
                WallConeDataFile(room) = CStr(Cone_Data_FileTextBox.Text)
            Case "wallsubstrate"
                WallSubstrate(room) = Material_DescriptionListBox.Text
                WallSubConductivity(room) = CDbl(Thermal_ConductivityTextBox.Text)
                WallSubSpecificHeat(room) = CDbl(Specific_HeatTextBox.Text)
                WallSubDensity(room) = CDbl(DensityTextBox.Text)
            Case "ceilingsurface"
                CeilingSurface(room) = Material_DescriptionListBox.Text
                CeilingConductivity(room) = CDbl(Thermal_ConductivityTextBox.Text)
                CeilingSpecificHeat(room) = CDbl(Specific_HeatTextBox.Text)
                CeilingDensity(room) = CDbl(DensityTextBox.Text)
                Surface_Emissivity(1, room) = CDbl(EmissivityTextBox.Text)
                CeilingEffectiveHeatofCombustion(room) = CDbl(Heat_of_CombustionTextBox.Text)
                CeilingSootYield(room) = CDbl(Soot_YieldTextBox.Text)
                CeilingH2OYield(room) = CDbl(H2O_YieldTextBox.Text)
                CeilingCO2Yield(room) = CDbl(CO2_YieldTextBox.Text)
                CeilingConeDataFile(room) = CStr(Cone_Data_FileTextBox.Text)
            Case "ceilingsubstrate"
                CeilingSubstrate(room) = Material_DescriptionListBox.Text
                ' frmDescribeRoom.lblCeilingSubstrate.Text = CeilingSubstrate(room)
                CeilingSubConductivity(room) = CDbl(Thermal_ConductivityTextBox.Text)
                CeilingSubSpecificHeat(room) = CDbl(Specific_HeatTextBox.Text)
                CeilingSubDensity(room) = CDbl(DensityTextBox.Text)
            Case "floorsurface"
                FloorSurface(room) = Material_DescriptionListBox.Text
                FloorConductivity(room) = CDbl(Thermal_ConductivityTextBox.Text)
                FloorSpecificHeat(room) = CDbl(Specific_HeatTextBox.Text)
                FloorDensity(room) = CDbl(DensityTextBox.Text)
                Surface_Emissivity(4, room) = CDbl(EmissivityTextBox.Text)
                FloorTSMin(room) = CDbl(Min_Temp_for_SpreadTextBox.Text) + 273
                FloorEffectiveHeatofCombustion(room) = CDbl(Heat_of_CombustionTextBox.Text)
                FloorSootYield(room) = CDbl(Soot_YieldTextBox.Text)
                FloorH2OYield(room) = CDbl(H2O_YieldTextBox.Text)
                FloorCO2Yield(room) = CDbl(CO2_YieldTextBox.Text)
                FloorFlameSpreadParameter(room) = CDbl(Flame_Spread_ParameterTextBox.Text)
                FloorConeDataFile(room) = CStr(Cone_Data_FileTextBox.Text)
            Case "floorsubstrate"
                FloorSubstrate(room) = Material_DescriptionListBox.Text
                FloorSubConductivity(room) = CDbl(Thermal_ConductivityTextBox.Text)
                FloorSubSpecificHeat(room) = CDbl(Specific_HeatTextBox.Text)
                FloorSubDensity(room) = CDbl(DensityTextBox.Text)
        End Select

        Close()
        loadroomdata(room)

        Exit Sub

handler:
        MsgBox("There was a problem extracting data from the materials database.")
    End Sub

    Public Sub loadroomdata(ByVal room As Integer)
        Dim orooms As List(Of oRoom)
        Dim oroomdistributions As List(Of oDistribution)
        If room > 0 Then
            orooms = RoomDB.GetRooms()
            Dim oRoom As oRoom = orooms(room - 1)
            oroomdistributions = RoomDB.GetRoomDistributions

            Dim NewroomForm As New frmNewRoom

            NewroomForm.Text = "Edit Room"
            NewroomForm.oRoom = oRoom

            If NewroomForm.oRoom IsNot Nothing Then
                NewroomForm.Tag = oRoom.num
                NewroomForm.txtID.Text = oRoom.num
                NewroomForm.txtWidth.Text = oRoom.width
                NewroomForm.txtLength.Text = oRoom.length
                NewroomForm.txtMinHeight.Text = oRoom.minheight
                NewroomForm.txtMaxHeight.Text = oRoom.maxheight
                NewroomForm.txtElevation.Text = oRoom.elevation
                NewroomForm.TextAbsX.Text = oRoom.absx
                NewroomForm.TextAbsY.Text = oRoom.absy
                NewroomForm.txtDescription.Text = oRoom.description

                NewroomForm.txtWallSurfaceThickness.Text = WallThickness(oRoom.num)
                NewroomForm.lblWallSurface.Text = WallSurface(oRoom.num)
                NewroomForm.txtWallSubstrateThickness.Text = WallSubThickness(oRoom.num)
                NewroomForm.lblWallSubstrate.Text = WallSubstrate(oRoom.num)
                NewroomForm.txtCeilingSurfaceThickness.Text = CeilingThickness(oRoom.num)
                NewroomForm.lblCeilingSurface.Text = CeilingSurface(oRoom.num)
                NewroomForm.txtCeilingSubstrateThickness.Text = CeilingSubThickness(oRoom.num)
                NewroomForm.lblCeilingSubstrate.Text = CeilingSubstrate(oRoom.num)
                NewroomForm.txtFloorSurfaceThickness.Text = FloorThickness(oRoom.num)
                NewroomForm.lblFloorSurface.Text = FloorSurface(oRoom.num)
                NewroomForm.txtFloorSubstrateThickness.Text = FloorSubThickness(oRoom.num)
                NewroomForm.lblFloorSubstrate.Text = FloorSubstrate(oRoom.num)

                For Each x In oroomdistributions
                    If x.id = oRoom.num Then
                        If x.varname = "width" Then
                            NewroomForm.txtWidth.Text = x.varvalue
                            If x.distribution <> "None" Then
                                NewroomForm.txtWidth.BackColor = distbackcolor
                            Else
                                NewroomForm.txtWidth.BackColor = distnobackcolor
                            End If
                        End If
                        If x.varname = "length" Then
                            NewroomForm.txtLength.Text = x.varvalue
                            If x.distribution <> "None" Then
                                NewroomForm.txtLength.BackColor = distbackcolor
                            Else
                                NewroomForm.txtLength.BackColor = distnobackcolor
                            End If
                        End If
                    End If
                Next

                NewroomForm.ShowDialog()

            End If

        End If
    End Sub

End Class