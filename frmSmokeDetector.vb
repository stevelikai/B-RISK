Option Strict Off
Option Explicit On
Friend Class frmSmokeDetector
	Inherits System.Windows.Forms.Form
	
	'UPGRADE_WARNING: Event chkSmokeDetector.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkSmokeDetector_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkSmokeDetector.CheckStateChanged
		Dim room As Short
		
		room = lstRoomSD.SelectedIndex + 1
		
		If chkSmokeDetector.CheckState = System.Windows.Forms.CheckState.Checked Then
			HaveSD(room) = True
			Frame2.Visible = True
			Frame3.Visible = True
			Frame4.Visible = True
		Else
			HaveSD(room) = False
			Frame2.Visible = False
			Frame3.Visible = False
			Frame4.Visible = False
		End If
		
	End Sub
	
	'UPGRADE_WARNING: Event chkUseODinside.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkUseODinside_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkUseODinside.CheckStateChanged
		Dim room As Short
		
		room = lstRoomSD.SelectedIndex + 1
		
		If chkUseODinside.CheckState = System.Windows.Forms.CheckState.Checked Then
			SDinside(room) = True
		Else
			SDinside(room) = False
		End If
	End Sub
	
	Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click

		'identify current room selected
		Dim id As Short
		id = lstRoomSD.SelectedIndex + 1
		
		If optSpecifyOD.Checked = True Then
			DetSensitivity(id) = 100 * (1 - 10 ^ (-SmokeOD(id) / 1000 * 25.4 * 12))
		ElseIf optDetSens.Checked = True Then 
			DetSensitivity(id) = CSng(txtDetSensitivity.Text)
			SmokeOD(id) = -System.Math.Log(-(DetSensitivity(id) / 100 - 1)) / System.Math.Log(10) / 12 / 25.4 * 1000
		End If
        Me.Hide()
		
	End Sub
	
	Private Sub frmSmokeDetector_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'***********************************
		'*  =============================================
		'*      This event centres the form on the screen
		'*      when the form first loads.
		'*  =============================================
		Dim room As Integer
		'call procedure to centre form
		Centre_Form(Me)
		
		For room = 1 To NumberRooms
			lstRoomSD.Items.Add(CStr(room))
		Next room
		lstRoomSD.SelectedIndex = 0
		Me.txtOpticalDensity.Text = VB6.Format(SmokeOD(1), "0.000")
		Me.txtSDdelay.Text = VB6.Format(SDdelay(1), "0.0")
		Me.txtDetSensitivity.Text = VB6.Format(DetSensitivity(1), "0.0")
		Me.txtSDRadialDist.Text = VB6.Format(SDRadialDist(1), "0.000")
		Me.txtSDdepth.Text = VB6.Format(SDdepth(1), "0.000")
		
		If SpecifyOD(1) = True Then
			Me.optSpecifyOD.Checked = True
			optSDnormal.Enabled = False
			optSDhigh.Enabled = False
			optSDvhigh.Enabled = False
			txtOpticalDensity.Enabled = True
			txtDetSensitivity.Enabled = True
		Else
			Me.optAS1603.Checked = True
			optSDnormal.Enabled = True
			optSDhigh.Enabled = True
			optSDvhigh.Enabled = True
			txtOpticalDensity.Enabled = False
			txtDetSensitivity.Enabled = False
			If SmokeOD(1) = SDNormalDefault Then
				Me.optSDnormal.Checked = True
				Me.txtDetSensitivity.Text = CStr(6.6)
			ElseIf SmokeOD(1) = SDhighDefault Then 
				Me.optSDhigh.Checked = True
				Me.txtDetSensitivity.Text = CStr(3.8)
			ElseIf SmokeOD(1) = SDvhighDefault Then 
				Me.optSDvhigh.Checked = True
				Me.txtDetSensitivity.Text = CStr(0.9)
			Else
				Me.optSDnormal.Checked = True
				SmokeOD(1) = SDNormalDefault
				txtOpticalDensity.Text = VB6.Format(SDNormalDefault, "0.000")
				Me.txtDetSensitivity.Text = CStr(6.6)
			End If
		End If
	End Sub
	

	'UPGRADE_WARNING: Event lstRoomSD.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstRoomSD_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstRoomSD.SelectedIndexChanged
        Dim id As Short
		
		'identify current room selected
		id = lstRoomSD.SelectedIndex + 1
		
		'update textboxes
		txtOpticalDensity.Text = VB6.Format(SmokeOD(id), "0.000")
		txtSDdelay.Text = VB6.Format(SDdelay(id), "0.0")
		txtDetSensitivity.Text = VB6.Format(DetSensitivity(id), "0.0")
		txtSDRadialDist.Text = VB6.Format(SDRadialDist(id), "0.0")
		txtSDdepth.Text = VB6.Format(SDdepth(id), "0.000")
		
		If HaveSD(id) = True Then
			Me.chkSmokeDetector.CheckState = System.Windows.Forms.CheckState.Checked
			Me.Frame2.Visible = True
			Me.Frame3.Visible = True
			Me.Frame4.Visible = True
		Else
			Me.chkSmokeDetector.CheckState = System.Windows.Forms.CheckState.Unchecked
			Me.Frame2.Visible = False
			Me.Frame3.Visible = False
			Me.Frame4.Visible = False
		End If
		If SDinside(id) = True Then
			Me.chkUseODinside.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			Me.chkUseODinside.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If SpecifyOD(id) = True Then
			Me.optSpecifyOD.Checked = True
			optSDnormal.Enabled = False
			optSDhigh.Enabled = False
			optSDvhigh.Enabled = False
			txtOpticalDensity.Enabled = True
			txtDetSensitivity.Enabled = True
		Else
			Me.optAS1603.Checked = True
			optSDnormal.Enabled = True
			optSDhigh.Enabled = True
			optSDvhigh.Enabled = True
			txtOpticalDensity.Enabled = False
			txtDetSensitivity.Enabled = False
			If SmokeOD(id) = SDNormalDefault Then
				Me.optSDnormal.Checked = True
				Me.txtDetSensitivity.Text = CStr(6.6)
			ElseIf SmokeOD(id) = SDhighDefault Then 
				Me.optSDhigh.Checked = True
				Me.txtDetSensitivity.Text = CStr(3.8)
			ElseIf SmokeOD(id) = SDvhighDefault Then 
				Me.optSDvhigh.Checked = True
				Me.txtDetSensitivity.Text = CStr(0.9)
			Else
				Me.optSDnormal.Checked = True
				SmokeOD(id) = SDNormalDefault
				txtOpticalDensity.Text = VB6.Format(SDNormalDefault, "0.000")
				Me.txtDetSensitivity.Text = CStr(6.6)
			End If
		End If
	End Sub
	
	
	'UPGRADE_WARNING: Event optAS1603.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optAS1603_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAS1603.CheckedChanged
		If eventSender.Checked Then
			Dim room As Byte
			room = Me.lstRoomSD.SelectedIndex + 1
			If optAS1603.Checked = True Then
				SpecifyOD(room) = False
				optSDnormal.Enabled = True
				optSDhigh.Enabled = True
				optSDvhigh.Enabled = True
				txtOpticalDensity.Enabled = False
				txtDetSensitivity.Enabled = False
			End If
			If SpecifyOD(room) = False Then
				If SmokeOD(room) = SDNormalDefault Then
					Me.optSDnormal.Checked = True
				ElseIf SmokeOD(room) = SDhighDefault Then 
					Me.optSDhigh.Checked = True
				ElseIf SmokeOD(room) = SDvhighDefault Then 
					Me.optSDvhigh.Checked = True
				Else
					Me.optSDnormal.Checked = True
					SmokeOD(room) = SDNormalDefault
					txtOpticalDensity.Text = VB6.Format(SDNormalDefault, "0.000")
					txtDetSensitivity.Text = CStr(6.6)
				End If
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optDetSens.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optDetSens_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDetSens.CheckedChanged
		If eventSender.Checked Then
			Dim room As Byte
			room = Me.lstRoomSD.SelectedIndex + 1
			If optDetSens.Checked = True Then
				SpecifyOD(room) = True
				optSDnormal.Enabled = False
				optSDhigh.Enabled = False
				optSDvhigh.Enabled = False
				txtOpticalDensity.Enabled = False
				txtDetSensitivity.Enabled = True
				If IsNumeric(txtDetSensitivity.Text) = True Then
					DetSensitivity(room) = CSng(txtDetSensitivity.Text)
				End If
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optSDhigh.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optSDhigh_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSDhigh.CheckedChanged
		If eventSender.Checked Then
			Dim room As Byte
			room = lstRoomSD.SelectedIndex + 1
			If optSDhigh.Checked = True Then
				txtOpticalDensity.Text = VB6.Format(SDhighDefault, "0.000")
				txtDetSensitivity.Text = CStr(3.8)
				SmokeOD(room) = SDhighDefault
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optSDnormal.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optSDnormal_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSDnormal.CheckedChanged
		If eventSender.Checked Then
			Dim room As Byte
			room = lstRoomSD.SelectedIndex + 1
			If optSDnormal.Checked = True Then
				txtOpticalDensity.Text = VB6.Format(SDNormalDefault, "0.000")
				SmokeOD(room) = SDNormalDefault
				txtDetSensitivity.Text = CStr(6.6)
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optSDvhigh.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optSDvhigh_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSDvhigh.CheckedChanged
		If eventSender.Checked Then
			Dim room As Byte
			room = lstRoomSD.SelectedIndex + 1
			If optSDvhigh.Checked = True Then
				txtOpticalDensity.Text = VB6.Format(SDvhighDefault, "0.000")
				SmokeOD(room) = SDvhighDefault
				txtDetSensitivity.Text = CStr(0.9)
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optSpecifyOD.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optSpecifyOD_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSpecifyOD.CheckedChanged
		If eventSender.Checked Then
			Dim room As Byte
			room = Me.lstRoomSD.SelectedIndex + 1
			If optSpecifyOD.Checked = True Then
				SpecifyOD(room) = True
				optSDnormal.Enabled = False
				optSDhigh.Enabled = False
				optSDvhigh.Enabled = False
				txtOpticalDensity.Enabled = True
				txtDetSensitivity.Enabled = False
				If IsNumeric(txtOpticalDensity.Text) = True Then
					SmokeOD(room) = CSng(txtOpticalDensity.Text)
				End If
			End If
		End If
	End Sub
	
	
	
	'UPGRADE_WARNING: Event txtDetSensitivity.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtDetSensitivity_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDetSensitivity.TextChanged
		'*  ===================================================
		'*      This event stores the data input as a variable
		'*      when the contents of the text box are changed.
		'*  ===================================================
		Dim room As Short
		room = lstRoomSD.SelectedIndex + 1
		If IsNumeric(txtDetSensitivity.Text) = True Then
			'store variable
			DetSensitivity(room) = CSng(txtDetSensitivity.Text)
		End If
		
	End Sub
	
	Private Sub txtDetSensitivity_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDetSensitivity.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim A As Single
		Dim room As Short
		room = lstRoomSD.SelectedIndex + 1
		If IsNumeric(txtDetSensitivity.Text) = True Then
			A = CSng(txtDetSensitivity.Text)
			
			If KeyAscii = 13 Then
				If A > 0 Then
					DetSensitivity(room) = A
					'convert sensitivity to OD at alarm
					If DetSensitivity(room) < 100 Then SmokeOD(room) = -System.Math.Log(-(DetSensitivity(room) / 100 - 1)) / System.Math.Log(10) / 12 / 25.4 * 1000
					txtOpticalDensity.Text = VB6.Format(SmokeOD(room), "0.000")
				Else
					txtDetSensitivity.Focus()
				End If
			End If
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtDetSensitivity_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDetSensitivity.Leave
		Dim A As Single
		If IsNumeric(txtDetSensitivity.Text) = True Then
			A = CSng(txtDetSensitivity.Text)
			
			If A > 0 Then
			Else
				MsgBox("Invalid Entry!")
				txtDetSensitivity.Focus()
			End If
		End If
		
	End Sub
	
	'UPGRADE_WARNING: Event txtOpticalDensity.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtOpticalDensity_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOpticalDensity.TextChanged
		'*  ===================================================
		'*      This event stores the data input as a variable
		'*      when the contents of the text box are changed.
		'*  ===================================================
		Dim room As Short
		room = lstRoomSD.SelectedIndex + 1
		If IsNumeric(txtOpticalDensity.Text) = True Then
			'store radial distance as a variable
			SmokeOD(room) = CSng(txtOpticalDensity.Text)
			
			
		End If
	End Sub
	
	Private Sub txtOpticalDensity_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtOpticalDensity.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'*  ===================================================
		'*      This event checks for valid input when the user
		'*      pressing the enter key and if valid stores the
		'*      data as a variable.
		'*  ===================================================
		
		Dim A As Single
		Dim room As Short
		room = lstRoomSD.SelectedIndex + 1
		If IsNumeric(txtOpticalDensity.Text) = True Then
			A = CSng(txtOpticalDensity.Text)
			
			If KeyAscii = 13 Then
				If A > 0 Then
					SmokeOD(room) = A
					
					DetSensitivity(room) = 100 * (1 - 10 ^ (-SmokeOD(room) / 1000 * 25.4 * 12))
					txtDetSensitivity.Text = VB6.Format(DetSensitivity(room), "0.0")
					
				Else
					txtOpticalDensity.Focus()
				End If
			End If
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtOpticalDensity_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOpticalDensity.Leave
		'*  =====================================================
		'*      This event checks for valid data when the control
		'*      loses the focus.
		'*  =====================================================
		
		Dim A As Single
		If IsNumeric(txtOpticalDensity.Text) = True Then
			A = CSng(txtOpticalDensity.Text)
			
			If A > 0 Then
			Else
				MsgBox("Invalid Entry!")
				txtOpticalDensity.Focus()
			End If
		End If
		
	End Sub
	
	'UPGRADE_WARNING: Event txtSDdelay.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSDdelay_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSDdelay.TextChanged
		'*  ===================================================
		'*      This event stores the data input as a variable
		'*      when the contents of the text box are changed.
		'*  ===================================================
		Dim room As Short
		room = lstRoomSD.SelectedIndex + 1
		If IsNumeric(txtSDdelay.Text) = True Then
			'store radial distance as a variable
			SDdelay(room) = CSng(txtSDdelay.Text)
		End If
	End Sub
	
	
	Private Sub txtSDdelay_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSDdelay.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'*  ===================================================
		'*      This event checks for valid input when the user
		'*      pressing the enter key and if valid stores the
		'*      data as a variable.
		'*  ===================================================
		
		Dim A As Single
		Dim room As Short
		room = lstRoomSD.SelectedIndex + 1
		If IsNumeric(txtSDdelay.Text) = True Then
			A = CSng(txtSDdelay.Text)
			
			If KeyAscii = 13 Then
				If A >= 0 Then
					SDdelay(room) = A
				Else
					txtSDdelay.Focus()
				End If
			End If
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	
	Private Sub txtSDdelay_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSDdelay.Leave
		'*  =====================================================
		'*      This event checks for valid data when the control
		'*      loses the focus.
		'*  =====================================================
		
		Dim A As Single
		If IsNumeric(txtSDdelay.Text) = True Then
			A = CSng(txtSDdelay.Text)
			
			If A >= 0 Then
			Else
				MsgBox("Invalid Entry!")
				txtSDdelay.Focus()
			End If
		End If
		
	End Sub
	
	
	
	
	'UPGRADE_WARNING: Event txtSDdepth.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSDdepth_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSDdepth.TextChanged
		'*  ===================================================
		'*      This event stores the data input as a variable
		'*      when the contents of the text box are changed.
		'*  ===================================================
		Dim room As Short
		room = lstRoomSD.SelectedIndex + 1
		If IsNumeric(txtSDdepth.Text) = True Then
			'store radial distance as a variable
			SDdepth(room) = CSng(txtSDdepth.Text)
		End If
		
	End Sub
	
	Private Sub txtSDdepth_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSDdepth.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'*  ===================================================
		'*      This event checks for valid input when the user
		'*      pressing the enter key and if valid stores the
		'*      data as a variable.
		'*  ===================================================
		
		Dim A As Single
		Dim room As Short
		room = lstRoomSD.SelectedIndex + 1
		If IsNumeric(txtSDdepth.Text) = True Then
			A = CSng(txtSDdepth.Text)
			
			If KeyAscii = 13 Then
				If A >= 0 Then
					SDdepth(room) = A
				Else
					txtSDdepth.Focus()
				End If
			End If
			
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtSDdepth_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSDdepth.Leave
		'*  =====================================================
		'*      This event checks for valid data when the control
		'*      loses the focus.
		'*  =====================================================
		
		Dim A As Single
		If IsNumeric(txtSDdepth.Text) = True Then
			A = CSng(txtSDdepth.Text)
			
			If A >= 0 Then
			Else
				MsgBox("Invalid Entry!")
				txtSDdepth.Focus()
			End If
		End If
		
	End Sub
	
	'UPGRADE_WARNING: Event txtSDRadialDist.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSDRadialDist_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSDRadialDist.TextChanged
		'*  ===================================================
		'*      This event stores the data input as a variable
		'*      when the contents of the text box are changed.
		'*  ===================================================
		Dim room As Short
		room = lstRoomSD.SelectedIndex + 1
		If IsNumeric(txtSDRadialDist.Text) = True Then
			'store radial distance as a variable
			SDRadialDist(room) = CSng(txtSDRadialDist.Text)
		End If
	End Sub
	
	Private Sub txtSDRadialDist_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSDRadialDist.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'*  ===================================================
		'*      This event checks for valid input when the user
		'*      pressing the enter key and if valid stores the
		'*      data as a variable.
		'*  ===================================================
		
		Dim A As Single
		Dim room As Short
		room = lstRoomSD.SelectedIndex + 1
		If IsNumeric(txtSDRadialDist.Text) = True Then
			A = CSng(txtSDRadialDist.Text)
			
			If KeyAscii = 13 Then
				If A >= 0 Then
					SDRadialDist(room) = A
				Else
					txtSDRadialDist.Focus()
				End If
			End If
			
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtSDRadialDist_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSDRadialDist.Leave
		'*  =====================================================
		'*      This event checks for valid data when the control
		'*      loses the focus.
		'*  =====================================================
		
		Dim A As Single
		If IsNumeric(txtSDRadialDist.Text) = True Then
			A = CSng(txtSDRadialDist.Text)
			
			If A >= 0 Then
			Else
				MsgBox("Invalid Entry!")
				txtSDRadialDist.Focus()
			End If
		End If
	End Sub
End Class