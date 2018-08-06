Option Strict Off
Option Explicit On
Friend Class frmExtract
	Inherits System.Windows.Forms.Form
	
    Private Sub chkUseFanCurve_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkUseFanCurve.CheckStateChanged
        Dim id As Short
        id = lstRoomID.SelectedIndex + 1
        If id = 0 Then id = 1
        If chkUseFanCurve.CheckState = System.Windows.Forms.CheckState.Checked Then
            UseFanCurve(id) = True
        Else
            UseFanCurve(id) = False
        End If
    End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click

        Dim id As Short
		On Error Resume Next
		id = lstRoomID.SelectedIndex + 1
		ExtractRate(id) = CSng(txtCeilingExtractRate.Text)
		NumberFans(id) = CShort(txtNumberFans.Text)
		ExtractStartTime(id) = CSng(txtOperationTime.Text)
		FanElevation(id) = CDbl(VB6.Format(CDbl(txtFanElevation.Text), "0.00"))
		MaxPressure(id) = CSng(txtMaxPressure.Text)
		Extract(id) = Me.optExtract.Checked
		UseFanCurve(id) = Me.chkUseFanCurve.CheckState
		
		If Me.optFanOn.Checked = True Then
			fanon(id) = True
		Else
			fanon(id) = False
		End If
		If Me.optFanAuto.Checked = True Then
			FanAutoStart(id) = True
		Else
			FanAutoStart(id) = False
		End If
		
        Me.Hide()

	End Sub
	
	Private Sub frmExtract_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'*  =============================================
		'*      This event centres the form on the screen
		'*      when the form first loads.
		'*  =============================================
		
		Dim i As Short
		'add rooms to list box
		If NumberRooms > 0 Then
			lstRoomID.Items.Clear()
			For i = 1 To NumberRooms
				lstRoomID.Items.Add(CStr(i))
			Next i
		End If
		lstRoomID.SelectedIndex = 0
		
		If fanon(lstRoomID.SelectedIndex + 1) = True Then
			Me.optFanOn.Checked = True
		Else
			Me.optFanOn.Checked = False
		End If
		
    End Sub
	

    Private Sub lstRoomID_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstRoomID.SelectedIndexChanged
        Dim id As Short
        id = lstRoomID.SelectedIndex + 1
        txtCeilingExtractRate.Text = CStr(ExtractRate(id))
        txtNumberFans.Text = CStr(NumberFans(id))
        txtOperationTime.Text = CStr(ExtractStartTime(id))
        txtFanElevation.Text = CStr(FanElevation(id))
        txtMaxPressure.Text = CStr(MaxPressure(id))
        Me.optExtract.Checked = Extract(id)
        If UseFanCurve(id) = True Then
            Me.chkUseFanCurve.CheckState = System.Windows.Forms.CheckState.Checked
        Else
            Me.chkUseFanCurve.CheckState = System.Windows.Forms.CheckState.Unchecked
        End If
        Me.optPressurise.Checked = Not Extract(id)

        If fanon(id) = True Then
            Me.optFanOn.Checked = True
        Else
            Me.optFanOff.Checked = True
        End If

    End Sub
	
    Private Sub optExtract_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optExtract.CheckedChanged
        If eventSender.Checked Then
            Dim id As Short
            id = lstRoomID.SelectedIndex + 1
            If id = 0 Then id = 1
            If optExtract.Checked = True Then
                Extract(id) = True
            Else
                Extract(id) = False
            End If
        End If
    End Sub
	
    Private Sub optFanAuto_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optFanAuto.CheckedChanged
        If eventSender.Checked Then
            Dim id As Short
            id = lstRoomID.SelectedIndex + 1
            If optFanAuto.Checked = True Then
                FanAutoStart(id) = True
            End If
        End If
    End Sub
	
    Private Sub optFanManual_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optFanManual.CheckedChanged
        If eventSender.Checked Then
            Dim id As Short
            id = lstRoomID.SelectedIndex + 1
            If optFanManual.Checked = True Then
                FanAutoStart(id) = False
            End If
        End If
    End Sub
	
    Private Sub optFanOff_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optFanOff.CheckedChanged
        If eventSender.Checked Then
            Dim id As Short
            id = lstRoomID.SelectedIndex + 1
            If optFanOn.Checked = True Then
                fanon(id) = True
            Else
                fanon(id) = False
            End If
        End If
    End Sub
	
    Private Sub optFanOn_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optFanOn.CheckedChanged
        If eventSender.Checked Then
            Dim id As Short
            id = lstRoomID.SelectedIndex + 1
            If optFanOn.Checked = True Then
                fanon(id) = True
            Else
                fanon(id) = False
            End If
        End If
    End Sub
	
    Private Sub optPressurise_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPressurise.CheckedChanged
        If eventSender.Checked Then
            Dim id As Short
            id = lstRoomID.SelectedIndex + 1
            If id = 0 Then id = 1
            If optPressurise.Checked = True Then
                Extract(id) = False
            Else
                Extract(id) = True
            End If
        End If
    End Sub
	
    Private Sub txtCeilingExtractRate_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCeilingExtractRate.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================
        Dim id As Short
        id = lstRoomID.SelectedIndex + 1
        If id = 0 Then id = 1
        If IsNumeric(txtCeilingExtractRate.Text) = True Then
            'store the energy yield as a variable
            ExtractRate(id) = CSng(txtCeilingExtractRate.Text)
        End If
    End Sub
	
	Private Sub txtCeilingExtractRate_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCeilingExtractRate.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'*  ===================================================
		'*      This event checks for valid input when the user
		'*      pressing the enter key and if valid stores the
		'*      data as a variable.
		'*  ===================================================
		
		Dim a As Single
		Dim id As Short
		If IsNumeric(txtCeilingExtractRate.Text) = True Then
			a = CSng(txtCeilingExtractRate.Text)
			id = lstRoomID.SelectedIndex + 1
			If KeyAscii = 13 Then
				If a >= 0 Then
					ExtractRate(id) = a
				End If
			End If
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtCeilingExtractRate_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCeilingExtractRate.Leave
		'*  =====================================================
		'*      This event checks for valid data when the control
		'*      loses the focus.
		'*  =====================================================
		
		Dim a As Single
		If IsNumeric(txtCeilingExtractRate.Text) = True Then
			a = CSng(txtCeilingExtractRate.Text)
			If a >= 0 Then
			Else
				MsgBox("Invalid Dimension!")
				txtCeilingExtractRate.Focus()
			End If
		Else
			MsgBox("Invalid Dimension!")
			txtCeilingExtractRate.Focus()
		End If
	End Sub
	
    Private Sub txtFanElevation_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFanElevation.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================
        Dim id As Short
        id = lstRoomID.SelectedIndex + 1
        If id = 0 Then id = 1
        If IsNumeric(txtFanElevation.Text) = True Then
            'store the energy yield as a variable
            FanElevation(id) = CDbl(VB6.Format(CDbl(txtFanElevation.Text), "0.00"))
        End If
    End Sub
	
	Private Sub txtFanElevation_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFanElevation.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'*  ===================================================
		'*      This event checks for valid input when the user
		'*      pressing the enter key and if valid stores the
		'*      data as a variable.
		'*  ===================================================
		
		Dim a As Single
		Dim id As Short
		id = lstRoomID.SelectedIndex
		If id = 0 Then GoTo EventExitSub
		If IsNumeric(txtFanElevation.Text) = True Then
			a = CDbl(txtFanElevation.Text)
			If KeyAscii = 13 Then
				If a >= 0 And a <= RoomHeight(id + 1) Then
					FanElevation(id) = a
				End If
			End If
		End If
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtFanElevation_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFanElevation.Leave
		'*  =====================================================
		'*      This event checks for valid data when the control
		'*      loses the focus.
		'*  =====================================================
		
		Dim a As Single
		Dim id As Short
		id = lstRoomID.SelectedIndex
		If id < 0 Then Exit Sub
		If IsNumeric(txtFanElevation.Text) = True Then
			a = CDbl(txtFanElevation.Text)
			If a >= 0 And a <= RoomHeight(id + 1) Then
			Else
				MsgBox("Invalid Dimension! Fan elevation must be greater or equal to 0 and less than or equal to the ceiling height in the room.")
				txtFanElevation.Focus()
				txtFanElevation.Text = CStr(RoomHeight(id + 1))
			End If
		Else
			MsgBox("Invalid Dimension! Fan elevation must be greater or equal to 0 and less than or equal to the ceiling height in the room.")
			txtFanElevation.Focus()
		End If
	End Sub
	
    Private Sub txtMaxPressure_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMaxPressure.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================
        Dim id As Short
        id = lstRoomID.SelectedIndex + 1
        If id = 0 Then id = 1
        'store the energy yield as a variable
        If IsNumeric(txtMaxPressure.Text) = True Then
            MaxPressure(id) = CSng(txtMaxPressure.Text)
        End If
    End Sub
	
	Private Sub txtMaxPressure_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxPressure.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'*  ===================================================
		'*      This event checks for valid input when the user
		'*      pressing the enter key and if valid stores the
		'*      data as a variable.
		'*  ===================================================
		
		Dim a As Single
		Dim id As Short
		If IsNumeric(txtMaxPressure.Text) = True Then
			a = CSng(txtMaxPressure.Text)
			id = lstRoomID.SelectedIndex + 1
			If KeyAscii = 13 Then
				If a > 0 Then
					MaxPressure(id) = a
				End If
			End If
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtMaxPressure_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMaxPressure.Leave
		'*  =====================================================
		'*      This event checks for valid data when the control
		'*      loses the focus.
		'*  =====================================================
		
		Dim a As Single
		If IsNumeric(txtMaxPressure.Text) = True Then
			a = CSng(txtMaxPressure.Text)
			If a > 0 Then
			Else
				MsgBox("Invalid Dimension! Pressure difference must be greater than zero.")
				txtMaxPressure.Focus()
			End If
		Else
			MsgBox("Invalid Dimension!")
			txtMaxPressure.Focus()
		End If
	End Sub
	
    Private Sub txtNumberFans_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNumberFans.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================
        Dim id As Short
        id = lstRoomID.SelectedIndex + 1
        If id = 0 Then id = 1
        If IsNumeric(txtNumberFans.Text) = True Then
            'store the energy yield as a variable
            NumberFans(id) = CShort(txtNumberFans.Text)
        End If
    End Sub
	
	Private Sub txtNumberFans_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtNumberFans.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'*  ===================================================
		'*      This event checks for valid input when the user
		'*      pressing the enter key and if valid stores the
		'*      data as a variable.
		'*  ===================================================
		
		Dim a As Single
		Dim id As Short
		If IsNumeric(txtNumberFans.Text) = True Then
			a = CShort(txtNumberFans.Text)
			id = lstRoomID.SelectedIndex + 1
			If KeyAscii = 13 Then
				If a >= 1 Then
					NumberFans(id) = a
				End If
			End If
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtNumberFans_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNumberFans.Leave
		'*  =====================================================
		'*      This event checks for valid data when the control
		'*      loses the focus.
		'*  =====================================================
		
		Dim a As Single
		If IsNumeric(txtNumberFans.Text) = True Then
			a = CShort(txtNumberFans.Text)
			If a >= 1 Then
			Else
				MsgBox("There must be at least 1 fan")
				txtNumberFans.Focus()
			End If
		Else
			MsgBox("Invalid Entry")
			txtNumberFans.Focus()
		End If
		
	End Sub
	
    Private Sub txtOperationTime_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOperationTime.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================
        Dim id As Short
        id = lstRoomID.SelectedIndex + 1
        If id = 0 Then id = 1
        'store the energy yield as a variable
        If IsNumeric(txtOperationTime.Text) = True Then
            ExtractStartTime(id) = CSng(txtOperationTime.Text)
        End If
    End Sub
	
	Private Sub txtOperationTime_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtOperationTime.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'*  ===================================================
		'*      This event checks for valid input when the user
		'*      pressing the enter key and if valid stores the
		'*      data as a variable.
		'*  ===================================================
		
		Dim a As Single
		Dim id As Short
		id = lstRoomID.SelectedIndex
		If IsNumeric(txtOperationTime.Text) = True Then
			a = CSng(txtOperationTime.Text)
			If KeyAscii = 13 Then
				If a >= 0 Then
					ExtractStartTime(id) = a
				End If
			End If
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtOperationTime_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOperationTime.Leave
		'*  =====================================================
		'*      This event checks for valid data when the control
		'*      loses the focus.
		'*  =====================================================
		
		Dim a As Single
		If IsNumeric(txtOperationTime.Text) = True Then
			a = CSng(txtOperationTime.Text)
			If a >= 0 Then
			Else
				MsgBox("Invalid Dimension!")
				txtOperationTime.Focus()
			End If
		Else
			MsgBox("Invalid Dimension!")
			txtOperationTime.Focus()
		End If
	End Sub

End Class