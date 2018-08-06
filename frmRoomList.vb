Imports System
Imports System.Collections.Generic
Public Class frmRoomList
    Inherits System.Windows.Forms.Form
    Public NewRoomForm As frmNewRoom
    Dim oRooms As List(Of oRoom)
    Dim oroomdistributions As List(Of oDistribution)

    Private Sub frmRoomList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim sep As String = "   "
        Dim text As String
        oRooms = RoomDB.GetRooms

        Me.FillRoomList(oRooms)
        text = String.Format("{0,-4} {1,-20} {2,10} {3,10} {4,10}", "Room", "Description", "Length (m)", "Width (m)", "Height (m)")

        ListBox1.Items.Add(text)
        AddOwnedForm(frmNewRoom)

        CenterToParent()
    End Sub
    Public Sub FillRoomList(ByVal oRooms)
        lstRooms.Items.Clear()
        Dim j As Integer = 0

        For Each p As oRoom In oRooms
            j = j + 1
            lstRooms.Items.Add(p.GetDisplayText(vbTab))
        Next

    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdAddRoom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddRoom.Click
        Try
            If NumberRooms = MaxNumberRooms Then
                MsgBox("You have exceeded the maximum number of permitted rooms", MsgBoxStyle.Exclamation)
                Exit Sub
            End If

            oRooms = RoomDB.GetRooms
            Dim oroom As oRoom
            oroomdistributions = RoomDB.GetRoomDistributions

            NewRoomForm = New frmNewRoom
            NewRoomForm.Text = "New Room"
            oroom = New oRoom(1, "Room A", False, 0, 0, 0, 2.4, 2.4, 3.6, 2.4)

            oroom.num = oRooms.Count + 1
            NewRoomForm.Tag = oRooms.Count + 1

            If oroom IsNot Nothing Then
                oRooms.Add(oroom)

                NewRoomForm.txtID.Text = oroom.num
                NewRoomForm.txtWidth.Text = oroom.width
                NewRoomForm.txtLength.Text = oroom.length
                NewRoomForm.txtMinHeight.Text = oroom.minheight
                NewRoomForm.txtMaxHeight.Text = oroom.maxheight
                NewRoomForm.txtElevation.Text = oroom.elevation
                NewRoomForm.TextAbsX.Text = oroom.absx
                NewRoomForm.TextAbsY.Text = oroom.absy
                NewRoomForm.txtDescription.Text = oroom.description

                'ReDim HaveWallSubstrate(0 To oRooms.Count)
                'ReDim HaveCeilingSubstrate(0 To oRooms.Count)
                'ReDim HaveFloorSubstrate(0 To oRooms.Count)

                Dim oDistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
                oDistribution.id = oroom.num
                oDistribution.varname = "width"
                oDistribution.varvalue = 2.4
                NewRoomForm.txtWidth.Text = oDistribution.varvalue
                oroomdistributions.Add(oDistribution)

                oDistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
                oDistribution.id = oroom.num
                oDistribution.varname = "length"
                oDistribution.varvalue = 3.6
                NewRoomForm.txtLength.Text = oDistribution.varvalue
                oroomdistributions.Add(oDistribution)
                RoomDB.SaveRooms(oRooms, oroomdistributions)

                Resize_Rooms()

                TwoZones(oRooms.Count) = True

                WallThickness(oRooms.Count) = WallThickness(1)
                WallSurface(oRooms.Count) = WallSurface(1)
                WallSubThickness(oRooms.Count) = WallSubThickness(1)
                WallSubstrate(oRooms.Count) = WallSubstrate(1)
                WallConductivity(oRooms.Count) = WallConductivity(1)
                WallDensity(oRooms.Count) = WallDensity(1)
                WallSpecificHeat(oRooms.Count) = WallSpecificHeat(1)
                Surface_Emissivity(1, oRooms.Count) = Surface_Emissivity(1, 1)
                Surface_Emissivity(2, oRooms.Count) = Surface_Emissivity(2, 1)
                Surface_Emissivity(3, oRooms.Count) = Surface_Emissivity(3, 1)
                Surface_Emissivity(4, oRooms.Count) = Surface_Emissivity(4, 1)

                CeilingThickness(oRooms.Count) = CeilingThickness(1)
                CeilingSurface(oRooms.Count) = CeilingSurface(1)
                CeilingSubThickness(oRooms.Count) = CeilingSubThickness(1)
                CeilingSubstrate(oRooms.Count) = CeilingSubstrate(1)
                CeilingConductivity(oRooms.Count) = CeilingConductivity(1)
                CeilingDensity(oRooms.Count) = CeilingDensity(1)
                CeilingSpecificHeat(oRooms.Count) = CeilingSpecificHeat(1)

                FloorThickness(oRooms.Count) = FloorThickness(1)
                FloorSurface(oRooms.Count) = FloorSurface(1)
                FloorSubThickness(oRooms.Count) = FloorSubThickness(1)
                FloorSubstrate(NumberRooms) = FloorSubstrate(1)
                FloorConductivity(oRooms.Count) = FloorConductivity(1)
                FloorDensity(oRooms.Count) = FloorDensity(1)
                FloorSpecificHeat(oRooms.Count) = FloorSpecificHeat(1)

                NewRoomForm.txtWallSurfaceThickness.Text = WallThickness(1)
                NewRoomForm.lblWallSurface.Text = WallSurface(1)
                NewRoomForm.txtWallSubstrateThickness.Text = WallSubThickness(1)
                NewRoomForm.lblWallSubstrate.Text = WallSubstrate(1)

                NewRoomForm.txtCeilingSurfaceThickness.Text = CeilingThickness(1)
                NewRoomForm.lblCeilingSurface.Text = CeilingSurface(1)
                NewRoomForm.txtCeilingSubstrateThickness.Text = CeilingSubThickness(1)
                NewRoomForm.lblCeilingSubstrate.Text = CeilingSubstrate(1)

                NewRoomForm.txtFloorSurfaceThickness.Text = FloorThickness(1)
                NewRoomForm.lblFloorSurface.Text = FloorSurface(1)
                NewRoomForm.txtFloorSubstrateThickness.Text = FloorSubThickness(1)
                NewRoomForm.lblFloorSubstrate.Text = FloorSubstrate(1)
            End If

            'change the outside room id for any existing vents
            Dim oVents As List(Of oVent)
            Dim oVentdistributions As List(Of oDistribution)
            Dim oVent As oVent
            oVents = VentDB.GetVents
            If oVents.Count > 0 Then
                oVentdistributions = VentDB.GetVentDistributions
                For Each oVent In oVents
                    If oVent.toroom = NumberRooms Then
                        oVent.toroom = NumberRooms + 1
                    End If
                Next
                VentDB.SaveVents(oVents, oVentdistributions)
            End If

            Me.FillRoomList(oRooms)
            lstRooms.SelectedIndex = oroom.num - 1
            NewRoomForm.ShowDialog()

            'NumberRooms = oRooms.Count
            'Resize_Rooms()
            'fillroomarrays(oRooms)
            'addroomcleanup(oroom)

            Exit Sub

        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub cmdEditRoom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEditRoom.Click

        Dim i As Integer = lstRooms.SelectedIndex
        If i <> -1 Then
            oRooms = RoomDB.GetRooms()
            Dim oRoom As oRoom = oRooms(i)
            oroomdistributions = RoomDB.GetRoomDistributions

            Dim NewroomForm As New frmNewRoom

            NewroomForm.Text = "Edit Room"
            NewroomForm.oRoom = oRoom

            If NewroomForm.oRoom IsNot Nothing Then
                NewroomForm.Tag = oRoom.num
                NewroomForm.txtID.Text = oRoom.num
                NewroomForm.txtWidth.Text = oroom.width
                NewroomForm.txtLength.Text = oroom.length
                NewroomForm.txtMinHeight.Text = oroom.minheight
                NewroomForm.txtMaxHeight.Text = oroom.maxheight
                NewroomForm.txtElevation.Text = oroom.elevation
                NewroomForm.TextAbsX.Text = oroom.absx
                NewroomForm.TextAbsY.Text = oroom.absy
                NewroomForm.txtDescription.Text = oroom.description

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
                    If x.id = oroom.num Then
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


            ' fillroomarrays(oRooms)

        End If
    End Sub

    Private Sub cmdCopyRoom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopyRoom.Click
        If NumberRooms = MaxNumberRooms Then
            MsgBox("You have exceeded the maximum number of permitted rooms", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Dim oRooms As List(Of oRoom)
        Dim oroomdistributions As List(Of oDistribution)
        Dim i As Integer = lstRooms.SelectedIndex
        Dim id As Integer
        Dim j As Integer = 0
        If i <> -1 Then
            oRooms = RoomDB.GetRooms
            Dim oRoom As oRoom = oRooms(i)
            id = oRoom.num
            oroomdistributions = RoomDB.GetRoomDistributions
            Dim odistribution As New oDistribution
            Dim NewRoomForm As New frmNewRoom

            NewRoomForm.Text = "Copy Room"
            NewRoomForm.oRoom = oRoom
            Dim counter As Integer
            If NewRoomForm.oRoom IsNot Nothing Then

                counter = 0

                For z = 0 To oroomdistributions.Count - 1
                    If oroomdistributions(z).id = id Then

                        Dim x As oDistribution = oroomdistributions(z)
                        Dim y As New oDistribution(x.varname, x.units, x.distribution, x.varvalue, x.mean, x.variance, x.lbound, x.ubound, x.mode, x.alpha, x.beta)
                        y.id = oRooms.Count + 1
                        oroomdistributions.Add(y)

                    End If
                Next
                NewRoomForm.oRoom.num = oRooms.Count + 1
                oRooms.Add(NewRoomForm.oroom)

                RoomDB.SaveRooms(oRooms, oroomdistributions)

                NumberRooms = oRooms.Count

                Me.FillRoomList(oRooms)

                Resize_Rooms()
                fillroomarrays(oRooms)
                addroomcleanup(oRoom)

                WallThickness(NumberRooms) = WallThickness(id)
                WallSurface(NumberRooms) = WallSurface(id)
                WallSubThickness(NumberRooms) = WallSubThickness(id)
                WallSubstrate(NumberRooms) = WallSubstrate(id)
                WallDensity(NumberRooms) = WallDensity(id)
                WallSpecificHeat(NumberRooms) = WallSpecificHeat(id)
                WallConductivity(NumberRooms) = WallConductivity(id)
                WallSubDensity(NumberRooms) = WallSubDensity(id)
                WallSubSpecificHeat(NumberRooms) = WallSubSpecificHeat(id)
                WallSubConductivity(NumberRooms) = WallSubConductivity(id)
                Surface_Emissivity(2, NumberRooms) = Surface_Emissivity(2, id)
                Surface_Emissivity(3, NumberRooms) = Surface_Emissivity(3, id)

                CeilingThickness(NumberRooms) = CeilingThickness(id)
                CeilingSurface(NumberRooms) = CeilingSurface(id)
                CeilingSubThickness(NumberRooms) = CeilingSubThickness(id)
                CeilingSubstrate(NumberRooms) = CeilingSubstrate(id)
                CeilingDensity(NumberRooms) = CeilingDensity(id)
                CeilingSpecificHeat(NumberRooms) = CeilingSpecificHeat(id)
                CeilingConductivity(NumberRooms) = CeilingConductivity(id)
                CeilingSubDensity(NumberRooms) = CeilingSubDensity(id)
                CeilingSubSpecificHeat(NumberRooms) = CeilingSubSpecificHeat(id)
                CeilingSubConductivity(NumberRooms) = CeilingSubConductivity(id)
                Surface_Emissivity(1, NumberRooms) = Surface_Emissivity(1, id)

                FloorThickness(NumberRooms) = FloorThickness(id)
                FloorSurface(NumberRooms) = FloorSurface(id)
                FloorSubThickness(NumberRooms) = FloorSubThickness(id)
                FloorSubstrate(NumberRooms) = FloorSubstrate(id)
                FloorDensity(NumberRooms) = FloorDensity(id)
                FloorSpecificHeat(NumberRooms) = FloorSpecificHeat(id)
                FloorConductivity(NumberRooms) = FloorConductivity(id)
                FloorSubDensity(NumberRooms) = FloorSubDensity(id)
                FloorSubSpecificHeat(NumberRooms) = FloorSubSpecificHeat(id)
                FloorSubConductivity(NumberRooms) = FloorSubConductivity(id)
                Surface_Emissivity(4, NumberRooms) = Surface_Emissivity(4, id)

                TwoZones(NumberRooms) = TwoZones(id)
            End If



            'change the outside room id for any existing vents
            Dim oVents As List(Of oVent)
            Dim oVentdistributions As List(Of oDistribution)
            Dim oVent As oVent
            oVents = VentDB.GetVents
            If oVents.Count > 0 Then
                oVentdistributions = VentDB.GetVentDistributions
                For Each oVent In oVents
                    If oVent.toroom = NumberRooms Then
                        oVent.toroom = NumberRooms + 1
                    End If
                Next
                VentDB.SaveVents(oVents, oVentdistributions)
            End If

            Call frmMatSelect.loadroomdata(NumberRooms)

        End If
        Exit Sub
    End Sub


    Public Sub fillroomarrays(ByVal orooms)
        Try
            For Each oRoom In orooms
                Dim i = oRoom.num
                RoomHeight(i) = oRoom.maxheight
                RoomWidth(i) = oRoom.width
                RoomLength(i) = oRoom.length
                RoomDescription(i) = oRoom.description
                FloorElevation(i) = oRoom.elevation
                RoomAbsX(i) = oRoom.absx
                RoomAbsY(i) = oRoom.absy
                MinStudHeight(i) = oRoom.minheight
                CeilingSlope(i) = oRoom.cslope

            Next
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "frmRoomList.vb fillroomarrays")
        End Try
    End Sub
    Public Sub addroomcleanup(ByVal oRoom)

        'call procedure to resize the vent arrays
        Resize_Vents()
        Resize_CVents()

        'frmDescribeRoom._lstRoomID_1.Items.Clear()
        'frmDescribeRoom._lstRoomID_2.Items.Clear()
        'frmDescribeRoom._lstRoomID_3.Items.Clear()
        'For i = 1 To NumberRooms
        '    frmDescribeRoom._lstRoomID_1.Items.Add(CStr(i))
        '    frmDescribeRoom._lstRoomID_2.Items.Add(CStr(i))
        '    frmDescribeRoom._lstRoomID_3.Items.Add(CStr(i))
        'Next
        'frmDescribeRoom._lstRoomID_1.SelectedIndex = 0
        'frmDescribeRoom._lstRoomID_2.SelectedIndex = 0
        'frmDescribeRoom._lstRoomID_3.SelectedIndex = 0

        'frmDescribeRoom._lstRoomID_1.Items.Add(CStr(NumberRooms))
        'frmDescribeRoom._lstRoomID_2.Items.Add(CStr(NumberRooms))
        'frmDescribeRoom._lstRoomID_3.Items.Add(CStr(NumberRooms))

        'frmDescribeRoom.lstCVentIDlower.Items.Clear()
        'frmDescribeRoom.lstCVentIDUpper.Items.Clear()
        'frmDescribeRoom.lstCVentIDlower.Items.Add("outside")
        'frmDescribeRoom.lstCVentIDUpper.Items.Add("outside")
        'frmDescribeRoom.lstCVentIDlower.SelectedIndex = 0
        'frmDescribeRoom.lstCVentIDUpper.SelectedIndex = frmDescribeRoom.lstCVentIDUpper.Items.Count - 1

        Dim val As Integer = (NumberRooms - 1)
        frmOptions1.lstRoomZone.Items.Clear()
        If frmOptions1.lstRoomZone.Items.Count > 0 Then
            frmOptions1.lstRoomZone.SelectedIndex = val
        End If

        'if we have added a new room, then set defaults, otherwise don't change them
        If WallThickness(NumberRooms) = 0 Then
            TwoZones(NumberRooms) = True
            WallThickness(NumberRooms) = 100
            CeilingThickness(NumberRooms) = 100
            WallSubThickness(NumberRooms) = 100
            CeilingSubThickness(NumberRooms) = 100
            FloorSubThickness(NumberRooms) = 100
            FloorThickness(NumberRooms) = 100
            HaveWallSubstrate(NumberRooms) = False
            HaveFloorSubstrate(NumberRooms) = False
            HaveCeilingSubstrate(NumberRooms) = False
            WallConductivity(NumberRooms) = 1.2
            CeilingConductivity(NumberRooms) = 1.2
            FloorConductivity(NumberRooms) = 1.2
            WallSubConductivity(NumberRooms) = 1.2
            FloorSubConductivity(NumberRooms) = 1.2
            CeilingSubConductivity(NumberRooms) = 1.2
            FloorConductivity(NumberRooms) = 1.2
            WallDensity(NumberRooms) = 2300
            CeilingDensity(NumberRooms) = 2300
            WallSubDensity(NumberRooms) = 2300
            FloorSubDensity(NumberRooms) = 2300
            CeilingSubDensity(NumberRooms) = 2300
            FloorDensity(NumberRooms) = 2300
            WallSpecificHeat(NumberRooms) = 880
            CeilingSpecificHeat(NumberRooms) = 880
            FloorSpecificHeat(NumberRooms) = 880
            WallSubSpecificHeat(NumberRooms) = 880
            FloorSubSpecificHeat(NumberRooms) = 880
            CeilingSubSpecificHeat(NumberRooms) = 880
            WallSubstrate(NumberRooms) = "None"
            CeilingSubstrate(NumberRooms) = "None"
            FloorSubstrate(NumberRooms) = "None"
            WallSurface(NumberRooms) = "concrete"
            CeilingSurface(NumberRooms) = "concrete"
            FloorSurface(NumberRooms) = "concrete"
            For i = 1 To 4
                Surface_Emissivity(i, NumberRooms) = 0.5
            Next i
        End If

    End Sub

    Private Sub cmdRemoveRoom_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveRoom.Click

        Dim i As Integer = lstRooms.SelectedIndex
        oRooms = RoomDB.GetRooms()

        Dim j As Integer = 0
        If i <> -1 Then
            Dim oroom As oRoom = oRooms(i)
            Dim oroomdistributions As List(Of oDistribution)

            Dim message As String = "Are you sure you want to delete " _
                & "Room " & CStr(i + 1) & "?"
            Dim button As DialogResult = MessageBox.Show(message,
                "Confirm Delete", MessageBoxButtons.YesNo)
            If button = Windows.Forms.DialogResult.Yes Then
                oroomdistributions = RoomDB.GetRoomDistributions()

here:
                For Each oDistribution In oroomdistributions
                    If oDistribution.id = oroom.num Then
                        oroomdistributions.Remove(oDistribution)
                        GoTo here
                    End If
                Next

                Dim id As Integer = oroom.num

                oRooms.Remove(oroom)

                'sort and reindex
                Dim count As Integer = 1
                For Each oroom In oRooms

                    For Each oDistribution In oroomdistributions
                        If oDistribution.id = oroom.num Then
                            oDistribution.id = count
                        End If
                    Next

                    oroom.num = count
                    count = count + 1
                Next

                NumberRooms = oRooms.Count

                RoomDB.SaveRooms(oRooms, oroomdistributions)
                Me.FillRoomList(oRooms)


                'cleanup
                If fireroom = id Then
                    If fireroom <= 1 Then
                    Else
                        fireroom = fireroom - 1
                    End If
                End If



                If id < NumberRooms + 1 Then
                    For i = id To NumberRooms
                        'RoomWidth(i) = RoomWidth(i + 1)
                        'RoomLength(i) = RoomLength(i + 1)
                        'RoomDescription(i) = RoomDescription(i + 1)
                        'FloorElevation(i) = FloorElevation(i + 1)
                        'RoomAbsX(i) = RoomAbsX(i + 1)
                        'RoomAbsY(i) = RoomAbsY(i + 1)
                        'RoomHeight(i) = RoomHeight(i + 1)
                        'MinStudHeight(i) = MinStudHeight(i + 1)
                        'CeilingSlope(i) = CeilingSlope(i + 1)
                        TwoZones(i) = TwoZones(i + 1)
                        WallThickness(i) = WallThickness(i + 1)
                        WallSubThickness(i) = WallSubThickness(i + 1)
                        CeilingThickness(i) = CeilingThickness(i + 1)
                        CeilingSubThickness(i) = CeilingSubThickness(i + 1)
                        FloorSubThickness(i) = FloorSubThickness(i + 1)
                        FloorThickness(i) = FloorThickness(i + 1)
                        HaveCeilingSubstrate(i) = HaveCeilingSubstrate(i + 1)
                        HaveWallSubstrate(i) = HaveWallSubstrate(i + 1)
                        HaveFloorSubstrate(i) = HaveFloorSubstrate(i + 1)
                        'HaveSD(i) = HaveSD(i + 1)
                        'SDinside(i) = SDinside(i + 1)
                        'SpecifyOD(i) = HaveSD(i + 1)
                        WallSurface(i) = WallSurface(i + 1)
                        CeilingSurface(i) = CeilingSurface(i + 1)
                        WallSubstrate(i) = WallSubstrate(i + 1)
                        CeilingSubstrate(i) = CeilingSubstrate(i + 1)
                        FloorSubstrate(i) = FloorSubstrate(i + 1)
                        FloorSurface(i) = FloorSurface(i + 1)
                        WallDensity(i) = WallDensity(i + 1)
                        WallSubDensity(i) = WallSubDensity(i + 1)
                        FloorSubDensity(i) = FloorSubDensity(i + 1)
                        CeilingDensity(i) = CeilingDensity(i + 1)
                        CeilingSubDensity(i) = CeilingSubDensity(i + 1)
                        WallConductivity(i) = WallConductivity(i + 1)
                        WallSubConductivity(i) = WallSubConductivity(i + 1)
                        FloorSubConductivity(i) = FloorSubConductivity(i + 1)
                        CeilingConductivity(i) = CeilingConductivity(i + 1)
                        CeilingSubConductivity(i) = CeilingSubConductivity(i + 1)
                        WallSpecificHeat(i) = WallSpecificHeat(i + 1)
                        WallSubSpecificHeat(i) = WallSubSpecificHeat(i + 1)
                        FloorSubSpecificHeat(i) = FloorSubSpecificHeat(i + 1)
                        CeilingSpecificHeat(i) = CeilingSpecificHeat(i + 1)
                        CeilingSubSpecificHeat(i) = CeilingSubSpecificHeat(i + 1)
                    Next i
                End If

                'call procedure to resize the arrays
                Resize_Rooms()
                fillroomarrays(oRooms)

                'update list box
                'frmDescribeRoom.ComboBox_RoomID.Items.Clear()
                'frmDescribeRoom._lstRoomID_1.Items.Clear()
                'frmDescribeRoom._lstRoomID_2.Items.Clear()
                'frmDescribeRoom._lstRoomID_3.Items.Clear()
                'frmDescribeRoom.lstCVentIDlower.Items.Clear()
                'frmDescribeRoom.lstCVentIDUpper.Items.Clear()
                frmOptions1.lstRoomZone.Items.Clear()
                'frmDescribeRoom.lstCVentID.Items.Clear()

                For i = 1 To NumberRooms
                    'frmDescribeRoom.ComboBox_RoomID.Items.Add(CStr(i))
                    'frmDescribeRoom._lstRoomID_1.Items.Add(CStr(i))
                    'frmDescribeRoom._lstRoomID_2.Items.Add(CStr(i))
                    'frmDescribeRoom._lstRoomID_3.Items.Add(CStr(i))
                    'frmDescribeRoom.lstCVentIDlower.Items.Add(CStr(i))
                    'frmDescribeRoom.lstCVentIDUpper.Items.Add(CStr(i))
                    frmOptions1.lstRoomZone.Items.Add(CStr(i))
                Next i
                'frmDescribeRoom._lstRoomID_1.SelectedIndex = 0
                'frmDescribeRoom._lstRoomID_2.SelectedIndex = 0
                'frmDescribeRoom._lstRoomID_3.SelectedIndex = 0

                'frmDescribeRoom.lstCVentIDlower.Items.Add("outside")
                'frmDescribeRoom.lstCVentIDUpper.Items.Add("outside")

                'lets delete any vent which connect to the deleted room 'id'
                'change any vents that connect to the outside
                ReDim NumberVents(0 To MaxNumberRooms + 1, 0 To MaxNumberRooms + 1)
                Dim ovents As List(Of oVent)
                Dim oventdistributions As List(Of oDistribution)
x1:
                ovents = VentDB.GetVents
                oventdistributions = VentDB.GetVentDistributions

                For Each oVent In ovents
                    Dim idr = oVent.fromroom
                    Dim idc = oVent.toroom
                    Dim idv = oVent.id

                    If idr = id Or idc = id Then
                        Call deletevent(idv - 1, ovents)
                        GoTo x1
                    End If

                Next

                For Each oVent In ovents
                    Dim idr = oVent.fromroom
                    Dim idc = oVent.toroom
                    Dim idv = oVent.id

                    If idr > id Then
                        oVent.fromroom = idr - 1
                    End If
                    If idc > id Then
                        oVent.toroom = idc - 1
                    End If
                Next
                VentDB.SaveVents(ovents, oventdistributions)

                Resize_CVents()

            End If
        End If
        Exit Sub
    End Sub
    Public Sub deletevent(ByVal i As Integer, ByVal ovents As Object)

        Dim j As Integer = 0
        If i <> -1 Then
            Dim ovent As oVent = ovents(i)
            Dim oventdistributions As List(Of oDistribution)

            oventdistributions = VentDB.GetVentDistributions()

here:
            For Each oDistribution In oventdistributions
                If oDistribution.id = ovent.id Then
                    oventdistributions.Remove(oDistribution)
                    GoTo here
                End If
            Next

            ovents.Remove(ovent)

            'sort and reindex
            Dim count As Integer = 1
            For Each ovent In ovents

                For Each oDistribution In oventdistributions
                    If oDistribution.id = ovent.id Then
                        oDistribution.id = count
                    End If
                Next

                ovent.id = count

                count = count + 1
            Next

            VentDB.SaveVents(ovents, oventdistributions)
            frmVentList.FillVentList()
        End If

    End Sub

End Class