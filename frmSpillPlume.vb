Option Strict Off
Option Explicit On
Friend Class frmSpillPlume
	Inherits System.Windows.Forms.Form
	Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        ''======================================================================
        '' close the form
        ''======================================================================
        ''check for validity of glass fracture model

        'Dim idc, idr, idv As Short
        ''identify current room
        'idr = frmDescribeRoom.lstRoom1.SelectedIndex + 1
        'If idr = 0 Then
        '	idr = 1
        'End If

        ''identify the connecting room
        'idc = frmDescribeRoom.lstRoom2.SelectedIndex + 1 + idr
        'If idc <= idr Then
        '	idc = idr + 1
        'End If

        'idv = frmDescribeRoom.lstVentID.SelectedIndex + 1
        'If idv = 0 Then idv = 1
        'If idv > 0 Then
        '	If OptHarrison.Checked = True Then
        '		spillplumemodel(idr, idc, idv) = 1
        '		'spillplume(idr,idc,idv)
        '	Else
        '		spillplumemodel(idr, idc, idv) = 0
        '	End If
        '	If optSingleSided.Checked = True Then
        '		spillbalconyprojection(idr, idc, idv) = True
        '	Else
        '		spillbalconyprojection(idr, idc, idv) = False
        '	End If
        '	If optSpillPlumeBalc.Checked = True Then
        '		SpillPlumeBalc(idr, idc, idv) = True
        '	Else
        '		SpillPlumeBalc(idr, idc, idv) = False
        '	End If
        'End If

        'Me.Hide()
        '      'frmBlank.Show()
        'frmDescribeRoom.Show()
    End Sub

    Private Sub frmSpillPlume_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'Centre_Form(Me)

        'Dim idc, idr, idv As Short
        ''identify current room
        'idr = frmDescribeRoom.lstRoom1.SelectedIndex + 1
        'If idr = 0 Then
        '	idr = 1
        'End If

        ''identify the connecting room
        'idc = frmDescribeRoom.lstRoom2.SelectedIndex + 1 + idr
        'If idc <= idr Then
        '	idc = idr + 1
        'End If

        'idv = frmDescribeRoom.lstVentID.SelectedIndex + 1
        'If idv = 0 Then idv = 1
        'If idv > 0 Then
        '	If spillplumemodel(idr, idc, idv) = 1 Then
        '		OptHarrison.Checked = True
        '		'fraBSOptions.Visible = False
        '	Else
        '		OptBRE368.Checked = True
        '		'fraBSOptions.Visible = True
        '	End If
        '	If spillbalconyprojection(idr, idc, idv) = True Then
        '		optSingleSided.Checked = True
        '	Else
        '		optDoubleSided.Checked = True
        '	End If
        'End If

    End Sub
	
	Private Sub frmSpillPlume_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		Me.Close()
		eventArgs.Cancel = Cancel
	End Sub
	
	Private Sub frmSpillPlume_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Me.Close()
	End Sub
	
    Private Sub txtDownstand_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDownstand.TextChanged
        ''*  ===================================================
        ''*      This event stores the data input as a variable
        ''*      when the contents of the text box are changed.
        ''*  ===================================================

        'Dim idv, idr, idc As Short

        ''identify current room
        'idr = frmDescribeRoom.lstRoom1.SelectedIndex + 1
        'If idr = 0 Then
        '    idr = 1
        'End If

        ''identify the connecting room
        'idc = frmDescribeRoom.lstRoom2.SelectedIndex + 1 + idr
        'If idc <= idr Then
        '    idc = idr + 1
        'End If

        'idv = frmDescribeRoom.lstVentID.SelectedIndex + 1
        'If idv = 0 Then idv = 1
        'If idv > 0 Then
        '    'store the downstand depth as a variable
        '    If IsNumeric(txtDownstand.Text) = True Then
        '        Downstand(idr, idc, idv) = CSng(txtDownstand.Text)
        '        Downstand(idc, idr, idv) = CSng(txtDownstand.Text)
        '    End If
        'End If

    End Sub
	
	Private Sub txtDownstand_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDownstand.KeyPress
        'Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        'Dim idr, id, idc As Short

        ''identify current room
        'idr = frmDescribeRoom.lstRoom1.SelectedIndex + 1
        'If idr = 0 Then
        '	idr = 1
        '	frmDescribeRoom.lstRoom1.SelectedIndex = 0
        'End If

        ''identify the connecting room
        'idc = frmDescribeRoom.lstRoom2.SelectedIndex + 1 + idr
        'If idc <= idr Then
        '	idc = idr + 1
        '	frmDescribeRoom.lstRoom2.SelectedIndex = idr - frmDescribeRoom.lstRoom1.SelectedIndex
        'End If

        'id = frmDescribeRoom.lstVentID.SelectedIndex + 1

        'If KeyAscii = 13 Then
        '	If IsNumeric(txtDownstand.Text) = True Then
        '		If CSng(txtDownstand.Text) >= 0 Then
        '			Downstand(idr, idc, id) = CSng(txtDownstand.Text)
        '			Downstand(idc, idr, id) = CSng(txtDownstand.Text)
        '		End If
        '	End If
        'End If

        'eventArgs.KeyChar = Chr(KeyAscii)
        'If KeyAscii = 0 Then
        '	eventArgs.Handled = True
        'End If
	End Sub
	
	Private Sub txtDownstand_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDownstand.Leave
		'*  =====================================================
		'*      This event checks for valid data when the control
		'*      loses the focus.
		'*  =====================================================
		If IsNumeric(txtDownstand.Text) = True Then
			If CSng(txtDownstand.Text) >= 0 Then
			Else
				MsgBox("Invalid Value!")
				txtDownstand.Focus()
			End If
		Else
			MsgBox("Invalid Value!")
			txtDownstand.Focus()
		End If
		
	End Sub
End Class