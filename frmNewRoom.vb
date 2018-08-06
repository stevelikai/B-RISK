Imports System.Collections.Generic
Imports System.ComponentModel

Public Class frmNewRoom
    Public oRoom As oRoom
    Public oRooms As List(Of oRoom)
    Public oroomdistribution As oDistribution
    Public oroomdistributions As List(Of oDistribution)
    Private Function IsValidData() As Boolean
        Return Validator.IsPresent(txtWidth) AndAlso
               Validator.IsPresent(txtLength) AndAlso
               Validator.IsPresent(txtMinHeight) AndAlso
                Validator.IsPresent(txtMaxHeight) AndAlso
                 Validator.IsPresent(txtElevation) AndAlso
                 Validator.IsPresent(TextAbsX) AndAlso
                  Validator.IsPresent(TextAbsY)

    End Function

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        If IsValidData() Then
            oRooms = RoomDB.GetRooms
            oroomdistributions = RoomDB.GetRoomDistributions

            For Each Me.oRoom In oRooms

                If Me.oRoom.num = Me.txtID.Text Then
                    oRoom.width = txtWidth.Text
                    oRoom.length = txtLength.Text
                    oRoom.minheight = txtMinHeight.Text
                    oRoom.maxheight = txtMaxHeight.Text
                    oRoom.elevation = txtElevation.Text
                    oRoom.absx = TextAbsX.Text
                    oRoom.absy = TextAbsY.Text
                    oRoom.description = txtDescription.Text

                    For Each x In oroomdistributions
                        If x.id = Me.txtID.Text Then
                            If x.varname = "width" Then x.varvalue = txtWidth.Text
                            If x.varname = "length" Then x.varvalue = txtLength.Text
                        End If
                    Next

                    WallSubThickness(oRoom.num) = CDbl(txtWallSubstrateThickness.Text)
                    If WallSubThickness(oRoom.num) = 0 Then
                        HaveWallSubstrate(oRoom.num) = False
                        WallSubstrate(oRoom.num) = "None"
                    Else
                        HaveWallSubstrate(oRoom.num) = True
                        WallSubstrate(oRoom.num) = lblWallSubstrate.Text
                    End If
                    WallThickness(oRoom.num) = CDbl(txtWallSurfaceThickness.Text)
                    WallSurface(oRoom.num) = lblWallSurface.Text

                    CeilingSubThickness(oRoom.num) = CDbl(txtCeilingSubstrateThickness.Text)
                    If CeilingSubThickness(oRoom.num) = 0 Then
                        HaveCeilingSubstrate(oRoom.num) = False
                        CeilingSubstrate(oRoom.num) = "None"
                    Else
                        HaveCeilingSubstrate(oRoom.num) = True
                        CeilingSubstrate(oRoom.num) = lblCeilingSubstrate.Text
                    End If
                    CeilingThickness(oRoom.num) = CDbl(txtCeilingSurfaceThickness.Text)
                    CeilingSurface(oRoom.num) = lblCeilingSurface.Text

                    FloorSubThickness(oRoom.num) = CDbl(txtFloorSubstrateThickness.Text)
                    If FloorSubThickness(oRoom.num) = 0 Then
                        HaveFloorSubstrate(oRoom.num) = False
                        FloorSubstrate(oRoom.num) = "None"
                    Else
                        HaveFloorSubstrate(oRoom.num) = True
                        FloorSubstrate(oRoom.num) = lblFloorSubstrate.Text
                    End If
                    FloorThickness(oRoom.num) = CDbl(txtFloorSurfaceThickness.Text)
                    FloorSurface(oRoom.num) = lblFloorSurface.Text

                End If
            Next

            If oRoom IsNot Nothing Then
                RoomDB.SaveRooms(oRooms, oroomdistributions)
                frmRoomList.FillRoomList(oRooms)
                Resize_Rooms()
                frmRoomList.fillroomarrays(oRooms)
                frmRoomList.addroomcleanup(oRoom)
            End If


        End If
        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdDist_length_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_length.Click
        Dim oRooms As List(Of oRoom)
        Dim oroomdistributions As List(Of oDistribution)
        Dim roomid = Me.Tag
        Dim param As String = "length"
        Dim units As String = "(m)"
        Dim instruction As String = "Room Length"
        Dim paramdist As String

        Try

            oRooms = RoomDB.GetRooms()
            oroomdistributions = RoomDB.GetRoomDistributions()

            Call frmDistParam.ShowRoomDistributionForm(param, units, instruction, oRooms, oroomdistributions, roomid, paramdist)

            For Each odistribution In oroomdistributions
                If odistribution.id = roomid And odistribution.varname = param Then
                    txtLength.Text = odistribution.varvalue
                    If paramdist <> "None" Then
                        txtLength.BackColor = distbackcolor
                    Else
                        txtLength.BackColor = distnobackcolor
                    End If
                    Exit For
                End If
            Next

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "cmdDist_length_Click")
        End Try
    End Sub

    Private Sub cmdDist_width_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_width.Click
        Dim oRooms As List(Of oRoom)
        Dim oroomdistributions As List(Of oDistribution)
        Dim roomid = Me.Tag
        Dim param As String = "width"
        Dim units As String = "(m)"
        Dim instruction As String = "Room Width"
        Dim paramdist As String

        Try

            oRooms = RoomDB.GetRooms()
            oroomdistributions = RoomDB.GetRoomDistributions()

            Call frmDistParam.ShowRoomDistributionForm(param, units, instruction, oRooms, oroomdistributions, roomid, paramdist)

            For Each odistribution In oroomdistributions
                If odistribution.id = roomid And odistribution.varname = param Then
                    txtWidth.Text = odistribution.varvalue
                    If paramdist <> "None" Then
                        txtWidth.BackColor = distbackcolor
                    Else
                        txtWidth.BackColor = distnobackcolor
                    End If
                    Exit For
                End If
            Next

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "cmdDist_width_Click")
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles cmdPickWsurface.Click
        '====================================================
        '   Show the thermal database form
        '====================================================
        thermal = "wallsurface"
        frmMatSelect.Tag = Me.Tag
        Close()
        frmMatSelect.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles cmdPickWSubstrate.Click
        '====================================================
        '   Show the thermal database form
        '====================================================
        thermal = "wallsubstrate"
        frmMatSelect.Tag = Me.Tag
        Close()
        frmMatSelect.Show()
    End Sub

    Private Sub cmdPickCsurface_Click(sender As Object, e As EventArgs) Handles cmdPickCsurface.Click
        thermal = "ceilingsurface"
        frmMatSelect.Tag = Me.Tag
        Close()
        frmMatSelect.Show()
    End Sub

    Private Sub cmdPickCsubstrate_Click(sender As Object, e As EventArgs) Handles cmdPickCsubstrate.Click
        thermal = "ceilingsubstrate"
        frmMatSelect.Tag = Me.Tag
        Close()
        frmMatSelect.Show()
    End Sub

    Private Sub cmdPickFsurface_Click(sender As Object, e As EventArgs) Handles cmdPickFsurface.Click
        thermal = "floorsurface"
        frmMatSelect.Tag = Me.Tag
        Close()
        frmMatSelect.Show()
    End Sub

    Private Sub cmdPickFsubstrate_Click(sender As Object, e As EventArgs) Handles cmdPickFsubstrate.Click
        thermal = "floorsubstrate"
        frmMatSelect.Tag = Me.Tag
        Close()
        frmMatSelect.Show()
    End Sub

    Private Sub txtWallSurfaceThickness_Validated(sender As Object, e As EventArgs) Handles txtWallSurfaceThickness.Validated
        ErrorProvider1.Clear()
        WallThickness(txtID.Text) = txtWallSurfaceThickness.Text

    End Sub

    Private Sub txtWallSurfaceThickness_Validating(sender As Object, e As CancelEventArgs) Handles txtWallSurfaceThickness.Validating
        If IsNumeric(txtWallSurfaceThickness.Text) Then
            If (CDbl(txtWallSurfaceThickness.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtWallSurfaceThickness.Select(0, txtWallSurfaceThickness.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtWallSurfaceThickness, "Invalid Entry. Thickness must be greater than 0.")
    End Sub

    Private Sub txtCeilingSurfaceThickness_Validated(sender As Object, e As EventArgs) Handles txtCeilingSurfaceThickness.Validated
        ErrorProvider1.Clear()
        CeilingThickness(txtID.Text) = txtCeilingSurfaceThickness.Text
    End Sub

    Private Sub txtCeilingSurfaceThickness_Validating(sender As Object, e As CancelEventArgs) Handles txtCeilingSurfaceThickness.Validating
        If IsNumeric(txtCeilingSurfaceThickness.Text) Then
            If (CDbl(txtCeilingSurfaceThickness.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtCeilingSurfaceThickness.Select(0, txtCeilingSurfaceThickness.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtCeilingSurfaceThickness, "Invalid Entry. Thickness must be greater than 0.")
    End Sub

    Private Sub txtFloorSurfaceThickness_Validating(sender As Object, e As CancelEventArgs) Handles txtFloorSurfaceThickness.Validating
        If IsNumeric(txtFloorSurfaceThickness.Text) Then
            If (CDbl(txtFloorSurfaceThickness.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtFloorSurfaceThickness.Select(0, txtFloorSurfaceThickness.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtFloorSurfaceThickness, "Invalid Entry. Thickness must be greater than 0.")
    End Sub

    Private Sub txtFloorSurfaceThickness_Validated(sender As Object, e As EventArgs) Handles txtFloorSurfaceThickness.Validated
        ErrorProvider1.Clear()
        FloorThickness(txtID.Text) = txtFloorSurfaceThickness.Text
    End Sub

    Private Sub txtWallSubstrateThickness_Validated(sender As Object, e As EventArgs) Handles txtWallSubstrateThickness.Validated
        ErrorProvider1.Clear()
        WallSubThickness(txtID.Text) = txtWallSubstrateThickness.Text
    End Sub

    Private Sub txtWallSubstrateThickness_Validating(sender As Object, e As CancelEventArgs) Handles txtWallSubstrateThickness.Validating
        If IsNumeric(txtWallSubstrateThickness.Text) Then
            If (CDbl(txtWallSubstrateThickness.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtWallSubstrateThickness.Select(0, txtWallSubstrateThickness.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtWallSubstrateThickness, "Invalid Entry. Thickness must be greater than or equal to 0.")
    End Sub

    Private Sub txtCeilingSubstrateThickness_Validated(sender As Object, e As EventArgs) Handles txtCeilingSubstrateThickness.Validated
        ErrorProvider1.Clear()
        CeilingSubThickness(txtID.Text) = txtCeilingSubstrateThickness.Text
    End Sub

    Private Sub txtCeilingSubstrateThickness_Validating(sender As Object, e As CancelEventArgs) Handles txtCeilingSubstrateThickness.Validating
        If IsNumeric(txtCeilingSubstrateThickness.Text) Then
            If (CDbl(txtCeilingSubstrateThickness.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtCeilingSubstrateThickness.Select(0, txtCeilingSubstrateThickness.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtCeilingSubstrateThickness, "Invalid Entry. Thickness must be greater than or equal to 0.")
    End Sub

    Private Sub txtFloorSubstrateThickness_Validated(sender As Object, e As EventArgs) Handles txtFloorSubstrateThickness.Validated
        ErrorProvider1.Clear()
        FloorSubThickness(txtID.Text) = txtFloorSubstrateThickness.Text
    End Sub

    Private Sub txtFloorSubstrateThickness_Validating(sender As Object, e As CancelEventArgs) Handles txtFloorSubstrateThickness.Validating
        If IsNumeric(txtFloorSubstrateThickness.Text) Then
            If (CDbl(txtFloorSubstrateThickness.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtFloorSubstrateThickness.Select(0, txtFloorSubstrateThickness.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtFloorSubstrateThickness, "Invalid Entry. Thickness must be greater than or equal to 0.")
    End Sub

    Private Sub txtLength_TextChanged(sender As Object, e As EventArgs) Handles txtLength.TextChanged


    End Sub

    Private Sub txtLength_Validated(sender As Object, e As EventArgs) Handles txtLength.Validated
        ' If IsValidData() Then
        oRooms = RoomDB.GetRooms
            oroomdistributions = RoomDB.GetRoomDistributions

            For Each Me.oRoom In oRooms

                If Me.oRoom.num = Me.txtID.Text Then

                    oRoom.length = txtLength.Text

                    For Each x In oroomdistributions
                        If x.id = Me.txtID.Text Then
                            If x.varname = "length" Then x.varvalue = txtLength.Text
                        End If
                    Next
                End If
            Next
        'End If

        If oRoom IsNot Nothing Then
            RoomDB.SaveRooms(oRooms, oroomdistributions)
        End If

    End Sub



    Private Sub txtWidth_Validated(sender As Object, e As EventArgs) Handles txtWidth.Validated
        'If IsValidData() Then
        oRooms = RoomDB.GetRooms
            oroomdistributions = RoomDB.GetRoomDistributions

            For Each Me.oRoom In oRooms

                If Me.oRoom.num = Me.txtID.Text Then

                    oRoom.width = txtWidth.Text

                    For Each x In oroomdistributions
                        If x.id = Me.txtID.Text Then
                            If x.varname = "width" Then x.varvalue = txtWidth.Text
                        End If
                    Next
                End If
            Next
        'End If

        If oRoom IsNot Nothing Then
            RoomDB.SaveRooms(oRooms, oroomdistributions)
        End If
    End Sub

    Private Sub txtMinHeight_TextChanged(sender As Object, e As EventArgs) Handles txtMinHeight.TextChanged

    End Sub

    Private Sub txtMinHeight_Validated(sender As Object, e As EventArgs) Handles txtMinHeight.Validated
        'If IsValidData() Then
        oRooms = RoomDB.GetRooms
            oroomdistributions = RoomDB.GetRoomDistributions

            For Each Me.oRoom In oRooms

                If Me.oRoom.num = Me.txtID.Text Then

                    oRoom.minheight = txtMinHeight.Text

                End If
            Next
        'End If

        If oRoom IsNot Nothing Then
            RoomDB.SaveRooms(oRooms, oroomdistributions)
        End If
    End Sub



    Private Sub txtMaxHeight_Validated(sender As Object, e As EventArgs) Handles txtMaxHeight.Validated
        ' If IsValidData() Then
        oRooms = RoomDB.GetRooms
            oroomdistributions = RoomDB.GetRoomDistributions

            For Each Me.oRoom In oRooms

                If Me.oRoom.num = Me.txtID.Text Then

                    oRoom.maxheight = txtMaxHeight.Text

                End If
            Next
        ' End If

        If oRoom IsNot Nothing Then
            RoomDB.SaveRooms(oRooms, oroomdistributions)
        End If
    End Sub


    Private Sub txtElevation_Validated(sender As Object, e As EventArgs) Handles txtElevation.Validated
        'If IsValidData() Then
        oRooms = RoomDB.GetRooms
            oroomdistributions = RoomDB.GetRoomDistributions

            For Each Me.oRoom In oRooms

                If Me.oRoom.num = Me.txtID.Text Then

                    oRoom.elevation = txtElevation.Text

                End If
            Next
        'End If

        If oRoom IsNot Nothing Then
            RoomDB.SaveRooms(oRooms, oroomdistributions)
        End If
    End Sub


    Private Sub TextAbsY_Validated(sender As Object, e As EventArgs) Handles TextAbsY.Validated
        'If IsValidData() Then
        oRooms = RoomDB.GetRooms
            oroomdistributions = RoomDB.GetRoomDistributions

            For Each Me.oRoom In oRooms

                If Me.oRoom.num = Me.txtID.Text Then

                    oRoom.absy = TextAbsY.Text

                End If
            Next
        'End If

        If oRoom IsNot Nothing Then
            RoomDB.SaveRooms(oRooms, oroomdistributions)
        End If
    End Sub



    Private Sub txtDescription_Validated(sender As Object, e As EventArgs) Handles txtDescription.Validated
        'If IsValidData() Then
        oRooms = RoomDB.GetRooms
            oroomdistributions = RoomDB.GetRoomDistributions

            For Each Me.oRoom In oRooms

                If Me.oRoom.num = Me.txtID.Text Then

                    oRoom.description = txtDescription.Text

                End If
            Next
        'End If

        If oRoom IsNot Nothing Then
            RoomDB.SaveRooms(oRooms, oroomdistributions)
        End If
    End Sub



    Private Sub txtElevation_Validating(sender As Object, e As CancelEventArgs) Handles txtElevation.Validating
        If IsNumeric(txtElevation.Text) Then
            If (CDbl(txtElevation.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtElevation.Select(0, txtElevation.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtElevation, "Invalid Entry. Elevation must be greater or equal to 0. ")

    End Sub



    Private Sub TextAbsY_Validating(sender As Object, e As CancelEventArgs) Handles TextAbsY.Validating
        If IsNumeric(TextAbsY.Text) Then

            'okay
            Exit Sub

        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        TextAbsY.Select(0, TextAbsY.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(TextAbsY, "Invalid Entry. Must be a number. ")

    End Sub

    Private Sub txtWidth_Validating(sender As Object, e As CancelEventArgs) Handles txtWidth.Validating
        If IsNumeric(txtWidth.Text) Then
            If (CDbl(txtWidth.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtWidth.Select(0, txtWidth.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtWidth, "Invalid Entry. Must be greater than 0. ")
    End Sub

    Private Sub txtMinHeight_Validating(sender As Object, e As CancelEventArgs) Handles txtMinHeight.Validating
        If IsNumeric(txtMinHeight.Text) Then
            If (CDbl(txtMinHeight.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtMinHeight.Select(0, txtMinHeight.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtMinHeight, "Invalid Entry. Must be greater than 0. ")
    End Sub

    Private Sub txtMaxHeight_Validating(sender As Object, e As CancelEventArgs) Handles txtMaxHeight.Validating
        If IsNumeric(txtMaxHeight.Text) Then
            If (CDbl(txtMaxHeight.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtMaxHeight.Select(0, txtMaxHeight.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtMaxHeight, "Invalid Entry. Must be greater or equal to 0. ")

    End Sub

    Private Sub TextAbsX_Validating(sender As Object, e As CancelEventArgs) Handles TextAbsX.Validating
        If IsNumeric(TextAbsX.Text) Then
            'If (CDbl(TextAbsX.Text) > 0) Then
            'okay
            Exit Sub
            ' End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        TextAbsX.Select(0, TextAbsX.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(TextAbsX, "Invalid Entry. Must be a number. ")
    End Sub

    Private Sub TextAbsX_Validated(sender As Object, e As EventArgs) Handles TextAbsX.Validated
        'If IsValidData() Then
        oRooms = RoomDB.GetRooms
        oroomdistributions = RoomDB.GetRoomDistributions

        For Each Me.oRoom In oRooms

            If Me.oRoom.num = Me.txtID.Text Then

                oRoom.absx = TextAbsX.Text

            End If
        Next
        'End If

        If oRoom IsNot Nothing Then
            RoomDB.SaveRooms(oRooms, oroomdistributions)
        End If
    End Sub
End Class