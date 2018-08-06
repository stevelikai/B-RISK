Option Strict Off
Option Explicit On
Imports System.IO
Imports System.Collections.Generic

Friend Class frmDescribeRoom
    Inherits System.Windows.Forms.Form


    Private Sub chkAddSubstrate_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles chkAddSubstrate.CheckedChanged
        '===============================================================
        '   show extra information if substrate present
        '===============================================================
        Dim room As Short
        On Error Resume Next

        room = _lstRoomID_2.SelectedIndex + 1

        If chkAddSubstrate.Checked = True Then
            fraCeilingSubstrate.Visible = True
        Else
            fraCeilingSubstrate.Visible = False
        End If

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


    Private Sub txtCeilingSubThickness_TextChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles txtCeilingSubThickness.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================
        Dim room As Short
        On Error Resume Next

        room = _lstRoomID_2.SelectedIndex + 1

        If IsNumeric(txtCeilingSubThickness.Text) = True Then
            'store the ceiling thickness as a variable
            CeilingSubThickness(room) = CDbl(txtCeilingSubThickness.Text)
        End If
    End Sub

    Private Sub txtCeilingSubThickness_KeyPress(ByVal eventSender As Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCeilingSubThickness.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*  ===================================================
        '*      This event checks for valid input when the user
        '*      pressing the enter key and if valid stores the
        '*      data as a variable.
        '*  ===================================================
        Dim room As Short
        On Error Resume Next

        room = _lstRoomID_1.SelectedIndex + 1

        If IsNumeric(txtCeilingSubThickness.Text) = True Then
            If KeyAscii = 13 Then
                If CDbl(txtCeilingSubThickness.Text) > 0 Then
                    CeilingSubThickness(room) = CDbl(txtCeilingSubThickness.Text)
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If


    End Sub

    Private Sub txtCeilingSubThickness_Leave(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles txtCeilingSubThickness.Leave
        '*  =====================================================
        '*      This event checks for valid data when the control
        '*      loses the focus.
        '*  =====================================================
        On Error Resume Next

        If IsNumeric(txtCeilingSubThickness.Text) = True Then
            If CSng(txtCeilingSubThickness.Text) > 0 Then
            Else
                MsgBox("Invalid Dimension!")
                txtCeilingSubThickness.Focus()
            End If
        End If

    End Sub


    Private Sub txtFloorSubThickness_TextChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles txtFloorSubThickness.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================
        Dim room As Short
        On Error Resume Next

        room = _lstRoomID_3.SelectedIndex + 1

        If IsNumeric(txtFloorSubThickness.Text) = True Then
            'store the ceiling thickness as a variable
            FloorSubThickness(room) = CDbl(txtFloorSubThickness.Text)
        End If

    End Sub

    Private Sub txtFloorSubThickness_KeyPress(ByVal eventSender As Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFloorSubThickness.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*  ===================================================
        '*      This event checks for valid input when the user
        '*      pressing the enter key and if valid stores the
        '*      data as a variable.
        '*  ===================================================
        Dim room As Short
        On Error Resume Next

        room = _lstRoomID_1.SelectedIndex + 1
        If IsNumeric(txtFloorSubThickness.Text) = True Then
            If KeyAscii = 13 Then
                If CDbl(txtFloorSubThickness.Text) > 0 Then
                    FloorSubThickness(room) = CDbl(txtFloorSubThickness.Text)
                End If
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFloorSubThickness_Leave(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles txtFloorSubThickness.Leave
        '*  =====================================================
        '*      This event checks for valid data when the control
        '*      loses the focus.
        '*  =====================================================
        On Error Resume Next

        If IsNumeric(txtFloorSubThickness.Text) = True Then
            If CSng(txtFloorSubThickness.Text) > 0 Then
            Else
                MsgBox("Invalid Dimension!")
                txtFloorSubThickness.Focus()
            End If
        End If
    End Sub



    Private Sub txtWallSubThickness_TextChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles txtWallSubThickness.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================
        Dim room As Short

        room = _lstRoomID_1.SelectedIndex + 1
        If IsNumeric(txtWallSubThickness.Text) = True Then
            'store the wall thickness as a variable
            WallSubThickness(room) = CDbl(txtWallSubThickness.Text)
        End If
    End Sub

    Private Sub txtWallSubThickness_KeyPress(ByVal eventSender As Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtWallSubThickness.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*  ===================================================
        '*      This event checks for valid input when the user
        '*      pressing the enter key and if valid stores the
        '*      data as a variable.
        '*  ===================================================
        Dim room As Short

        room = _lstRoomID_1.SelectedIndex + 1
        If IsNumeric(txtWallSubThickness.Text) = True Then
            If KeyAscii = 13 Then
                If CDbl(txtWallSubThickness.Text) > 0 Then
                    WallSubThickness(room) = CDbl(txtWallSubThickness.Text)
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtWallSubThickness_Leave(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles txtWallSubThickness.Leave
        '*  =====================================================
        '*      This event checks for valid data when the control
        '*      loses the focus.
        '*  =====================================================
        If IsNumeric(txtWallSubThickness.Text) = True Then
            If CSng(txtWallSubThickness.Text) > 0 Then
            Else
                MsgBox("Invalid Dimension!")
                txtWallSubThickness.Focus()
            End If
        Else
            MsgBox("Invalid Dimension!")
            txtWallSubThickness.Focus()
        End If
    End Sub

    Private Sub txtCeilingThickness_TextChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles txtCeilingThickness.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================
        Dim id As Short

        'identify which room is selected
        id = _lstRoomID_2.SelectedIndex + 1
        If IsNumeric(txtCeilingThickness.Text) = True Then
            'store the ceiling thickness as a variable
            CeilingThickness(id) = CDbl(txtCeilingThickness.Text)
        End If
    End Sub

    Private Sub txtCeilingThickness_KeyPress(ByVal eventSender As Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCeilingThickness.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*  ===================================================
        '*      This event checks for valid input when the user
        '*      pressing the enter key and if valid stores the
        '*      data as a variable.
        '*  ===================================================
        Dim id As Short

        'identify which room is selected
        id = _lstRoomID_2.SelectedIndex + 1
        If IsNumeric(txtCeilingThickness.Text) = True Then
            If KeyAscii = 13 Then
                If CDbl(txtCeilingThickness.Text) > 0 Then
                    CeilingThickness(id) = CDbl(txtCeilingThickness.Text)
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCeilingThickness_Leave(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles txtCeilingThickness.Leave
        '*  =====================================================
        '*      This event checks for valid data when the control
        '*      loses the focus.
        '*  =====================================================
        If IsNumeric(txtCeilingThickness.Text) = True Then
            If CSng(txtCeilingThickness.Text) > 0 Then
            Else
                MsgBox("Invalid Dimension!")
                txtCeilingThickness.Focus()
            End If
        Else
            MsgBox("Invalid Dimension!")
            txtCeilingThickness.Focus()
        End If
    End Sub

    Private Sub txtFloorThickness_TextChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles txtFloorThickness.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================
        Dim id As Short

        'identify which room is selected
        id = _lstRoomID_3.SelectedIndex + 1

        If IsNumeric(txtFloorThickness.Text) = True Then
            'store the floor thickness as a variable
            FloorThickness(id) = CDbl(txtFloorThickness.Text)
        End If
    End Sub

    Private Sub txtFloorThickness_KeyPress(ByVal eventSender As Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFloorThickness.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*  ===================================================
        '*      This event checks for valid input when the user
        '*      pressing the enter key and if valid stores the
        '*      data as a variable.
        '*  ==================================================
        Dim id As Short

        'identify which room is selected
        id = _lstRoomID_3.SelectedIndex + 1
        If IsNumeric(txtFloorThickness.Text) = True Then
            If KeyAscii = 13 Then
                If CDbl(txtFloorThickness.Text) > 0 Then
                    FloorThickness(id) = CDbl(txtFloorThickness.Text)
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFloorThickness_Leave(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles txtFloorThickness.Leave
        '*  =====================================================
        '*      This event checks for valid data when the control
        '*      loses the focus.
        '*  =====================================================
        If IsNumeric(txtFloorThickness.Text) = True Then
            If CSng(txtFloorThickness.Text) > 0 Then
            Else
                MsgBox("Invalid Dimension!")
                txtFloorThickness.Focus()
            End If
        Else
            MsgBox("Invalid Dimension!")
            txtFloorThickness.Focus()
        End If
    End Sub

    Private Sub txtWallThickness_TextChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles txtWallThickness.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================
        Dim id As Short

        'identify which room is selected
        id = _lstRoomID_1.SelectedIndex + 1
        If IsNumeric(txtWallThickness.Text) = True Then
            'store the wall thickness as a variable
            WallThickness(id) = CDbl(txtWallThickness.Text)
        End If
    End Sub

    Private Sub txtWallThickness_KeyPress(ByVal eventSender As Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtWallThickness.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*  ===================================================
        '*      This event checks for valid input when the user
        '*      pressing the enter key and if valid stores the
        '*      data as a variable.
        '*  ===================================================
        Dim id As Short

        'identify which room is selected
        id = _lstRoomID_1.SelectedIndex + 1
        If IsNumeric(txtWallThickness.Text) = True Then
            If KeyAscii = 13 Then
                If CDbl(txtWallThickness.Text) > 0 Then
                    WallThickness(id) = CDbl(txtWallThickness.Text)
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtWallThickness_Leave(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles txtWallThickness.Leave
        '*  =====================================================
        '*      This event checks for valid data when the control
        '*      loses the focus.
        '*  =====================================================
        If IsNumeric(txtWallThickness.Text) = True Then
            If CSng(txtWallThickness.Text) > 0 Then
            Else
                MsgBox("Invalid Dimension!")
                txtWallThickness.Focus()
            End If
        Else
            MsgBox("Invalid Dimension!")
            txtWallThickness.Focus()
        End If
    End Sub

    Private Sub cmdPickCsubstrate_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles cmdPickCsubstrate.Click
        '====================================================
        '   Show the thermal database form
        '====================================================
        thermal = "ceilingsubstrate"
        Me.Hide()
        frmMatSelect.Show()
    End Sub

    Private Sub cmdPickCsurface_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles cmdPickCsurface.Click
        '====================================================
        '   Show the thermal database form
        '====================================================
        thermal = "ceilingsurface"
        Me.Hide()
        frmMatSelect.Show()

    End Sub

    Private Sub cmdPickFloor_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles cmdPickFloor.Click
        '====================================================
        '   Show the thermal database form
        '====================================================
        thermal = "floorsurface"
        Me.Hide()
        frmMatSelect.Show()

    End Sub

    Private Sub cmdPickSubstrate_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles cmdPickSubstrate.Click
        '====================================================
        '   Show the thermal database form
        '====================================================
        thermal = "wallsubstrate"
        Me.Hide()
        frmMatSelect.Show()

    End Sub

    Private Sub cmdPickSurface_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles cmdPickSurface.Click
        '====================================================
        '   Show the thermal database form
        '====================================================
        thermal = "wallsurface"
        Me.Hide()
        frmMatSelect.Show()

    End Sub



    Private Sub frmDescribeRoom_Load(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles MyBase.Load
        '*  =============================================
        '*      This event centres the form on the screen
        '*      when the form first loads.
        '*  =============================================

        Dim room1, i, room2 As Short

        Centre_Form(Me)

        Me.AddOwnedForm(frmMatSelect)

        room1 = 1
        room2 = NumberRooms + 1 'outside

        'add rooms to list box
        _lstRoomID_1.Items.Clear()
        _lstRoomID_2.Items.Clear()
        _lstRoomID_3.Items.Clear()

        If NumberRooms > 0 Then
            For i = 1 To NumberRooms
                _lstRoomID_1.Items.Add(CStr(i))
                _lstRoomID_2.Items.Add(CStr(i))
                _lstRoomID_3.Items.Add(CStr(i))
            Next i
        End If

        _lstRoomID_1.SelectedIndex = 0
        _lstRoomID_2.SelectedIndex = 0
        _lstRoomID_3.SelectedIndex = 0


    End Sub

    Private Sub cmdOK_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles cmdOK.Click
        Dim index As Short = cmdOK.GetIndex(eventSender)
        '*  ====================================================
        '*      This event makes the form
        '*      invisible.
        '*  ====================================================

        Call Check_Vent()
        Hide()

    End Sub


    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub



    Private Sub _lstRoomID_1_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles _lstRoomID_1.SelectedIndexChanged
        '*  ====================================================
        '*      This event updates the contents of the text boxes
        '*      when the user clicks on the room listbox.
        '*  ===================================================

        Dim id As Integer

        'identify current room selected
        id = _lstRoomID_1.SelectedIndex + 1

        'update textboxes 
        txtWallThickness.Text = Format$(WallThickness(id), "0.0")
        txtWallSubThickness.Text = Format$(WallSubThickness(id), "0.0")
        lblWallSurface.Text = WallSurface(id)
        lblWallSubstrate.Text = WallSubstrate(id)

        If HaveWallSubstrate(id) = True Then
            chkAddWallSubstrate.Checked = True
            fraWallSubstrate.Visible = True
        Else
            chkAddWallSubstrate.Checked = False
            fraWallSubstrate.Visible = False
        End If

    End Sub

    Private Sub _lstRoomID_2_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles _lstRoomID_2.SelectedIndexChanged
        '*  ====================================================
        '*      This event updates the contents of the text boxes
        '*      when the user clicks on the room listbox.
        '*  ===================================================

        Dim id As Integer

        'identify current room selected
        id = _lstRoomID_2.SelectedIndex + 1

        'update textboxes
        txtCeilingThickness.Text = Format$(CeilingThickness(id), "0.0")
        txtCeilingSubThickness.Text = Format$(CeilingSubThickness(id), "0.0")
        lblCeilingSurface.Text = CeilingSurface(id)
        lblCeilingSubstrate.Text = CeilingSubstrate(id)

        If HaveCeilingSubstrate(id) = True Then
            chkAddSubstrate.Checked = True
            fraCeilingSubstrate.Visible = True
        Else
            chkAddSubstrate.Checked = False
            fraCeilingSubstrate.Visible = False
        End If

    End Sub

    Private Sub _lstRoomID_3_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles _lstRoomID_3.SelectedIndexChanged
        '*  ====================================================
        '*      This event updates the contents of the text boxes
        '*      when the user clicks on the room listbox.
        '*  ===================================================

        Dim id As Integer

        'identify current room selected
        id = _lstRoomID_3.SelectedIndex + 1

        'update textboxes
        txtFloorSubThickness.Text = Format$(FloorSubThickness(id), "0.0")
        txtFloorThickness.Text = Format$(FloorThickness(id), "0.0")
        lblFloorSubstrate.Text = FloorSubstrate(id)
        lblFloorSurface.Text = FloorSurface(id)

        If HaveFloorSubstrate(id) = True Then
            chkAddFloorSubstrate.Checked = True
            fraFloorSubstrate.Visible = True
        Else
            chkAddFloorSubstrate.Checked = False
            fraFloorSubstrate.Visible = False
        End If

    End Sub


    Private Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs)

        MDIFrmMain.ViewToolStripMenuItem.PerformClick()

    End Sub


    Private Sub chkAddWallSubstrate_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles chkAddWallSubstrate.CheckedChanged
        '===============================================================
        '   show extra information if substrate present
        '===============================================================
        Dim room As Short
        On Error Resume Next

        room = _lstRoomID_1.SelectedIndex + 1

        If chkAddWallSubstrate.Checked = True Then
            fraWallSubstrate.Visible = True

        Else
            fraWallSubstrate.Visible = False

        End If
    End Sub

    Private Sub chkAddWallSubstrate_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkAddWallSubstrate.Click
        '===============================================================
        '   show extra information if substrate present
        '===============================================================
        Dim room As Short
        On Error Resume Next

        'room = ComboBox_RoomID.SelectedIndex + 1
        room = _lstRoomID_1.SelectedIndex + 1

        If chkAddWallSubstrate.Checked = True Then
            fraWallSubstrate.Visible = True
            HaveWallSubstrate(room) = True
        Else
            fraWallSubstrate.Visible = False
            HaveWallSubstrate(room) = False
        End If
    End Sub

    Private Sub chkAddSubstrate_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkAddSubstrate.Click
        '===============================================================
        '   show extra information if substrate present
        '===============================================================
        Dim room As Short
        On Error Resume Next

        'room = ComboBox_RoomID.SelectedIndex + 1

        room = _lstRoomID_2.SelectedIndex + 1

        If chkAddSubstrate.Checked = True Then
            fraCeilingSubstrate.Visible = True
            HaveCeilingSubstrate(room) = True
        Else
            fraCeilingSubstrate.Visible = False
            HaveCeilingSubstrate(room) = False
        End If
    End Sub

    Private Sub chkAddFloorSubstrate_Click(ByVal sender As Object, ByVal e As EventArgs) Handles chkAddFloorSubstrate.Click
        '===============================================================
        '   show extra information if substrate present
        '   last revised: 9/10/98
        '===============================================================
        Dim room As Short
        On Error Resume Next

        'room = ComboBox_RoomID.SelectedIndex + 1
        room = _lstRoomID_3.SelectedIndex + 1

        If chkAddFloorSubstrate.Checked = True Then
            fraFloorSubstrate.Visible = True
            HaveFloorSubstrate(room) = True
        Else
            fraFloorSubstrate.Visible = False
            HaveFloorSubstrate(room) = False
        End If
    End Sub


    Private Sub cmdPickFsubstrate_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdPickFsubstrate.Click
        '====================================================
        '   Show the thermal database form
        '====================================================
        thermal = "floorsubstrate"
        Me.Hide()
        frmMatSelect.Show()
    End Sub
End Class