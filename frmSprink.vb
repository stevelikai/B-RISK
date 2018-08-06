Option Strict Off
Option Explicit On
Public Class frmSprink
    Inherits System.Windows.Forms.Form

    Private Sub cboDetectorType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboDetectorType.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*      26 May 1998
        '*  ===================================================
        Dim room As Short
        room = lstRoomDetector.SelectedIndex + 1
        DetectorType(room) = cboDetectorType.SelectedIndex
        If DetectorType(room) = 0 Then
            FraSprOpt.Visible = False
        Else
            FraSprOpt.Visible = True
        End If
    End Sub

    Private Sub cboDetectorType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboDetectorType.SelectedIndexChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*      Revised 26 May 1998
        '*  ===================================================

        On Error Resume Next
        Dim room As Short

        room = lstRoomDetector.SelectedIndex + 1
        If room = 0 Then room = fireroom
        DetectorType(room) = cboDetectorType.SelectedIndex
        'Call cmdDefaults_Click
        If room = fireroom Then
            Me.txtRTI.Text = VB6.Format(RTI(room), "0.0")
            Me.txtCFactor.Text = VB6.Format(cfactor(room), "0.00")
            Me.txtRadialDistance.Text = VB6.Format(RadialDistance(room), "0.0")
            Me.txtLinkActuation.Text = VB6.Format(ActuationTemp(room) - 273, "0.0")
            Me.txtWaterSprayDensity.Text = VB6.Format(WaterSprayDensity(room) * 60, "0.0")
            Me.txtSprinkDistance.Text = VB6.Format(SprinkDistance(room), "0.000")

            Select Case DetectorType(room)
                Case 0 'none
                    FraSprOpt.Visible = False

                    txtCFactor.Visible = False
                    txtRTI.Visible = False
                    txtRadialDistance.Visible = False
                    txtLinkActuation.Visible = False
                    txtWaterSprayDensity.Visible = False
                    txtSprinkDistance.Visible = False
                    lblCFactor.Visible = False
                    lblRTI.Visible = False
                    lblradialDistance.Visible = False
                    lblLinkActuation.Visible = False
                    lblWaterSprayDensity.Visible = False
                    lblSprinkDistance.Visible = False
                Case 1 'heat detector
                    FraSprOpt.Visible = False

                    txtCFactor.Visible = False
                    txtRTI.Visible = True
                    txtRadialDistance.Visible = True
                    txtLinkActuation.Visible = True
                    txtWaterSprayDensity.Visible = False
                    txtSprinkDistance.Visible = True
                    lblCFactor.Visible = False
                    lblRTI.Visible = True
                    lblradialDistance.Visible = True
                    lblLinkActuation.Visible = True
                    lblWaterSprayDensity.Visible = False
                    lblSprinkDistance.Visible = True
                Case 2 'residential sprinkler
                    FraSprOpt.Visible = True

                    If optSuppression.Checked = True Then
                        txtWaterSprayDensity.Visible = True
                        lblWaterSprayDensity.Visible = True
                    Else
                        txtWaterSprayDensity.Visible = False
                        lblWaterSprayDensity.Visible = False
                    End If

                    txtCFactor.Visible = True
                    txtRTI.Visible = True
                    txtRadialDistance.Visible = True
                    txtLinkActuation.Visible = True

                    txtSprinkDistance.Visible = True
                    lblCFactor.Visible = True
                    lblRTI.Visible = True
                    lblradialDistance.Visible = True
                    lblLinkActuation.Visible = True

                    lblSprinkDistance.Visible = True
                Case 3 'commercial sprinkler
                    FraSprOpt.Visible = True

                    If optSuppression.Checked = True Then
                        txtWaterSprayDensity.Visible = True
                        lblWaterSprayDensity.Visible = True
                    Else
                        txtWaterSprayDensity.Visible = False
                        lblWaterSprayDensity.Visible = False
                    End If

                    txtCFactor.Visible = True
                    txtRTI.Visible = True
                    txtRadialDistance.Visible = True
                    txtLinkActuation.Visible = True

                    txtSprinkDistance.Visible = True
                    lblCFactor.Visible = True
                    lblRTI.Visible = True
                    lblradialDistance.Visible = True
                    lblLinkActuation.Visible = True

                    lblSprinkDistance.Visible = True

            End Select
        End If
    End Sub

    Private Sub cmdDefaults_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDefaults.Click
        '===================================================
        '   Returns the input text boxes to the default values
        '   based on the detector type selected
        '   26 May 1998
        '====================================================
        'Dim none As Variant
        'none = Array(0, 0, 3.2, 68, 0, True, False, False, False, False, False, False, False, False)
        Dim room As Short
        room = lstRoomDetector.SelectedIndex + 1
        If room = 0 Then room = fireroom
        If room = fireroom Then
            Select Case DetectorType(room)
                Case 0 'none
                    txtCFactor.Text = CStr(0)
                    txtRTI.Text = CStr(0)
                    txtRadialDistance.Text = CStr(0)
                    txtLinkActuation.Text = CStr(0)
                    txtWaterSprayDensity.Text = CStr(0)
                    txtSprinkDistance.Text = CStr(0)
                    optSprinkOff.Checked = True
                    txtCFactor.Visible = False
                    txtRTI.Visible = False
                    txtRadialDistance.Visible = False
                    txtLinkActuation.Visible = False
                    txtWaterSprayDensity.Visible = False
                    txtSprinkDistance.Visible = False
                Case 1 'heat detector
                    txtCFactor.Text = CStr(0)
                    txtRTI.Text = CStr(gcs_RTI_hd)
                    txtRadialDistance.Text = CStr(3.2)
                    txtLinkActuation.Text = CStr(57)
                    txtWaterSprayDensity.Text = CStr(0)
                    txtSprinkDistance.Text = CStr(gcs_sprdist)
                    optSuppression.Visible = False
                    txtCFactor.Visible = False
                    txtRTI.Visible = True
                    txtRadialDistance.Visible = True
                    txtLinkActuation.Visible = True
                    txtWaterSprayDensity.Visible = False
                    txtSprinkDistance.Visible = True
                Case 2 'residential sprinkler
                    txtCFactor.Text = CStr(gcs_cfactor)
                    txtRTI.Text = CStr(gcs_RTI_rspr)
                    txtRadialDistance.Text = CStr(4.3)
                    txtLinkActuation.Text = CStr(68)
                    txtWaterSprayDensity.Text = "5"
                    txtSprinkDistance.Text = CStr(gcs_sprdist)
                    optSprinkOff.Visible = True
                    optControl.Visible = True
                    optSuppression.Visible = True
                    optControl.Checked = True
                    txtCFactor.Visible = True
                    txtRTI.Visible = True
                    txtRadialDistance.Visible = True
                    txtLinkActuation.Visible = True
                    txtWaterSprayDensity.Visible = False
                    txtSprinkDistance.Visible = True
                Case 3 'commercial sprinkler
                    txtCFactor.Text = CStr(gcs_cfactor)
                    txtRTI.Text = CStr(gcs_RTI_cspr)
                    txtRadialDistance.Text = CStr(3.2)
                    txtLinkActuation.Text = CStr(74)
                    txtWaterSprayDensity.Text = "5" 'mm/min
                    txtSprinkDistance.Text = CStr(gcs_sprdist)
                    optSprinkOff.Visible = True
                    optControl.Visible = True
                    optSuppression.Visible = True
                    optControl.Checked = True
                    txtCFactor.Visible = True
                    txtRTI.Visible = True
                    txtRadialDistance.Visible = True
                    txtLinkActuation.Visible = True
                    txtWaterSprayDensity.Visible = False
                    txtSprinkDistance.Visible = True
                Case 4 'smoke detector
                    'txtCFactor.text = 0
                    'txtRTI.text = 1
                    'txtRadialDistance.text = 3.2
                    'txtLinkActuation.text = 13 + InteriorTemp - 273
                    'txtWaterSprayDensity.text = 0
                    'optSprinkOff.Value = True
                    'optSprinkOff.Visible = False
                    'optControl.Visible = False
                    'optSuppression.Visible = False
                    'txtCFactor.Visible = False
                    'txtRTI.Visible = True
                    'txtRadialDistance.Visible = True
                    'txtLinkActuation.Visible = True
                    'txtWaterSprayDensity.Visible = False
                    'txtSprinkDistance.Visible = True
            End Select
        End If
        cfactor(room) = CSng(txtCFactor.Text)
        RTI(room) = CSng(txtRTI.Text)
        ActuationTemp(room) = CSng(txtLinkActuation.Text) + 273
        WaterSprayDensity(room) = CSng(txtWaterSprayDensity.Text) / 60
        SprinkDistance(room) = CSng(txtSprinkDistance.Text)
        RadialDistance(room) = CSng(txtRadialDistance.Text)
    End Sub

    Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click

        ''set cancel to true
        ''UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
        'frmBlank.CMDialog1.CancelError = True
        'On Error GoTo errhandler
        ''set the help command property
        ''UPGRADE_ISSUE: MSComDlg.CommonDialog property CMDialog1.HelpCommand was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'frmBlank.CMDialog1.HelpCommand = MSComDlg.HelpConstants.cdlHelpForceFile
        ''frmBlank.CMDialog1.HelpCommand = cdlHelpContext
        ''specify the help file
        ''UPGRADE_ISSUE: MSComDlg.CommonDialog property CMDialog1.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'frmBlank.CMDialog1.HelpFile = "Branzfir.hlp"
        ''display the windows help engine
        ''UPGRADE_WARNING: The help CommonDialog is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E8158DFB-975E-4306-8E4F-43391F853639"'
        'frmBlank.CMDialog1.ShowHelp()
        'Exit Sub

errhandler:
        'user pressed cancel button
    End Sub

    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
        Dim room As Short
        If Me.cboDetectorType.SelectedIndex = 1 Or Me.cboDetectorType.SelectedIndex = 2 Or Me.cboDetectorType.SelectedIndex = 3 Then
            room = lstRoomDetector.SelectedIndex + 1

            If RTI(room) <= 0 Then
                MsgBox("RTI must be greater than 0", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
                Exit Sub
            End If
            If ActuationTemp(room) <= InteriorTemp Then
                MsgBox("Actuation temperature must be greater than the initial room temperature", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
                Exit Sub
            End If
        End If
        Me.Hide()

    End Sub



    Private Sub frmSprink_GotFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.GotFocus
        Dim room As Integer
        On Error Resume Next
        'call procedure to centre form
        Centre_Form(Me)

        'add objects to list box
        'cboDetectorType.AddItem "None"
        'cboDetectorType.AddItem "Heat Detector"
        'cboDetectorType.AddItem "Residential Sprinkler"
        'cboDetectorType.AddItem "Commercial Sprinkler"
        'cboDetectorType.AddItem "Smoke Detector"
        'cboDetectorType.ListIndex = DetectorType(fireroom)
        'For room = 1 To NumberRooms
        '    lstRoomDetector.AddItem CStr(room)
        'Next room
        'lstRoomDetector.ListIndex = fireroom - 1
        'cboDetectorType.Refresh
        'frmSprink.txtRTI.text = Format((RTI(fireroom)), "0.0")
        'frmSprink.txtCFactor.text = Format((cfactor(fireroom)), "0.00")
        'frmSprink.txtRadialDistance.text = Format((RadialDistance(fireroom)), "0.0")
        'frmSprink.txtLinkActuation.text = Format((ActuationTemp(fireroom) - 273), "0.0")
        'frmSprink.txtWaterSprayDensity.text = Format((WaterSprayDensity(fireroom) * 60), "0.0000")
        If Me.optSprinkOff.Checked = True Then
            Me.txtWaterSprayDensity.Visible = False
            Me.lblWaterSprayDensity.Visible = False
        ElseIf Me.optControl.Checked = True Then
            Me.txtWaterSprayDensity.Visible = False
            Me.lblWaterSprayDensity.Visible = False
        Else
            Me.txtWaterSprayDensity.Visible = True
            Me.lblWaterSprayDensity.Visible = True
        End If
        If cjModel = cjAlpert Then
            Me.txtSprinkDistance.Enabled = False
            Me.lblSprinkDistance.Enabled = False
        Else
            Me.txtSprinkDistance.Enabled = True
            Me.lblSprinkDistance.Enabled = True
        End If
        ''UPGRADE_ISSUE: ComboBox property cboDetectorType.index was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'If cboDetectorType.index = 0 Then FraSprOpt.Visible = False
        ''UPGRADE_ISSUE: ComboBox property cboDetectorType.index was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'If cboDetectorType.index = 1 Then FraSprOpt.Visible = False
        ''UPGRADE_ISSUE: ComboBox property cboDetectorType.index was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'If cboDetectorType.index = 2 Then FraSprOpt.Visible = True
        ''UPGRADE_ISSUE: ComboBox property cboDetectorType.index was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'If cboDetectorType.index = 3 Then FraSprOpt.Visible = True

    End Sub

    Private Sub frmSprink_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '***********************************
        '*  =============================================
        '*      This event centres the form on the screen
        '*      when the form first loads.
        '*  =============================================
        Dim room As Integer
        On Error Resume Next
        'call procedure to centre form
        Centre_Form(Me)

        'add objects to list box
        cboDetectorType.Items.Add("None")
        cboDetectorType.Items.Add("Heat Detector")
        cboDetectorType.Items.Add("Residential Sprinkler")
        cboDetectorType.Items.Add("Commercial Sprinkler")
        'cboDetectorType.AddItem "Smoke Detector"
        cboDetectorType.SelectedIndex = DetectorType(fireroom)
        For room = 1 To NumberRooms
            lstRoomDetector.Items.Add(CStr(room))
        Next room
        lstRoomDetector.SelectedIndex = fireroom - 1
        cboDetectorType.Refresh()
        Me.txtRTI.Text = VB6.Format(RTI(fireroom), "0.0")
        Me.txtCFactor.Text = VB6.Format(cfactor(fireroom), "0.00")
        Me.txtRadialDistance.Text = VB6.Format(RadialDistance(fireroom), "0.0")
        Me.txtLinkActuation.Text = VB6.Format(ActuationTemp(fireroom) - 273, "0.0")
        Me.txtWaterSprayDensity.Text = VB6.Format(WaterSprayDensity(fireroom) * 60, "0.0000")
        If Me.optSprinkOff.Checked = True Then
            Me.txtWaterSprayDensity.Visible = False
            Me.lblWaterSprayDensity.Visible = False
        ElseIf Me.optControl.Checked = True Then
            Me.txtWaterSprayDensity.Visible = False
            Me.lblWaterSprayDensity.Visible = False
        Else
            Me.txtWaterSprayDensity.Visible = True
            Me.lblWaterSprayDensity.Visible = True
        End If
        If cjModel = cjAlpert Then
            Me.txtSprinkDistance.Enabled = False
            Me.lblSprinkDistance.Enabled = False
        Else
            Me.txtSprinkDistance.Enabled = True
            Me.lblSprinkDistance.Enabled = True
        End If

        If cboDetectorType.SelectedIndex = 0 Then
            FraSprOpt.Visible = False
        ElseIf cboDetectorType.SelectedIndex = 1 Then
            FraSprOpt.Visible = False
        ElseIf cboDetectorType.SelectedIndex = 2 Then
            FraSprOpt.Visible = True
        ElseIf cboDetectorType.SelectedIndex = 3 Then
            FraSprOpt.Visible = True
        Else
        End If
    End Sub

    Private Sub frmSprink_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'Me.Close()
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmSprink_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'Me.Close()
    End Sub

    'UPGRADE_WARNING: Event lstRoomDetector.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub lstRoomDetector_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstRoomDetector.SelectedIndexChanged

        Dim room As Short
        room = lstRoomDetector.SelectedIndex + 1

        'If room = fireroom And cboDetectorType.ListIndex > 0 Then
        If room = fireroom And DetectorType(room) > 0 Then
            'If DetectorType(room) > 0 Then
            cmdDefaults.Visible = True
            'cboDetectorType.Visible = True
            'Frame2.Visible = True
            txtSprinkDistance.Text = CStr(SprinkDistance(room))
            txtRTI.Text = CStr(RTI(room))
            txtRadialDistance.Text = CStr(RadialDistance(room))
            txtLinkActuation.Text = CStr(ActuationTemp(room) - 273)
            txtCFactor.Text = CStr(cfactor(room))
            cboDetectorType.SelectedIndex = DetectorType(room)
            txtWaterSprayDensity.Text = CStr(WaterSprayDensity(room) * 60)

            Me.optControl.Checked = Sprinkler(1, room)
            Me.optSuppression.Checked = Sprinkler(3, room)
            Me.optSprinkOff.Checked = Sprinkler(2, room)
            txtSprinkDistance.Visible = True
            lblSprinkDistance.Visible = True
            txtRTI.Visible = True
            lblRTI.Visible = True
            txtRadialDistance.Visible = True
            lblradialDistance.Visible = True
            txtLinkActuation.Visible = True
            lblLinkActuation.Visible = True
            txtCFactor.Visible = True
            lblCFactor.Visible = True
            txtWaterSprayDensity.Visible = True
            lblWaterSprayDensity.Visible = True
            'If DetectorType(room) > 1 Then
            '    frmSprink.Frame2.Visible = True
            'Else
            '    frmSprink.Frame2.Visible = False
            'End If
        Else
            'cmdDefaults.Visible = False
            'Frame2.Visible = False
            'cboDetectorType.Visible = False
            txtSprinkDistance.Visible = False
            lblSprinkDistance.Visible = False
            txtRTI.Visible = False
            lblRTI.Visible = False
            txtRadialDistance.Visible = False
            lblradialDistance.Visible = False
            lblLinkActuation.Visible = False
            txtLinkActuation.Visible = False
            txtCFactor.Visible = False
            lblCFactor.Visible = False
            'cboDetectorType.ListIndex = DetectorType(room)
            txtWaterSprayDensity.Visible = False
            lblWaterSprayDensity.Visible = False

        End If

    End Sub

    'UPGRADE_WARNING: Event optControl.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optControl_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optControl.CheckedChanged
        If eventSender.Checked Then
            Dim room As Short
            room = lstRoomDetector.SelectedIndex + 1
            If optControl.Checked = True Then
                Sprinkler(1, room) = True
                Sprinkler(3, room) = False
                Sprinkler(2, room) = False
                Me.txtWaterSprayDensity.Visible = False
                Me.lblWaterSprayDensity.Visible = False
            End If
        End If
    End Sub

    'UPGRADE_WARNING: Event optSprinkOff.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optSprinkOff_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSprinkOff.CheckedChanged
        If eventSender.Checked Then
            Dim room As Short
            room = lstRoomDetector.SelectedIndex + 1
            If optSprinkOff.Checked = True Then
                Sprinkler(2, room) = True
                Sprinkler(1, room) = False
                Sprinkler(3, room) = False
                Me.txtWaterSprayDensity.Visible = False
                Me.lblWaterSprayDensity.Visible = False
            End If
        End If
    End Sub

    'UPGRADE_WARNING: Event optSuppression.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optSuppression_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSuppression.CheckedChanged
        If eventSender.Checked Then
            Dim room As Short
            room = lstRoomDetector.SelectedIndex + 1
            If optSuppression.Checked = True And DetectorType(room) > 0 Then
                Sprinkler(3, room) = True
                Sprinkler(1, room) = False
                Sprinkler(2, room) = False
                Me.txtWaterSprayDensity.Visible = True
                Me.lblWaterSprayDensity.Visible = True
            End If

        End If
    End Sub

    'UPGRADE_WARNING: Event txtCFactor.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtCFactor_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCFactor.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================
        Dim room As Short
        room = lstRoomDetector.SelectedIndex + 1
        If room = 0 Then Exit Sub
        If IsNumeric(txtCFactor.Text) = True Then
            'store sprinkler C Factor as a variable
            cfactor(room) = CSng(txtCFactor.Text)
        End If
    End Sub

    Private Sub txtCFactor_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCFactor.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*  ===================================================
        '*      This event checks for valid input when the user
        '*      pressing the enter key and if valid stores the
        '*      data as a variable.
        '*  ===================================================

        Dim a As Single
        Dim room As Short
        room = lstRoomDetector.SelectedIndex + 1
        If IsNumeric(txtCFactor.Text) = True Then
            a = CSng(txtCFactor.Text)

            If KeyAscii = 13 Then
                If a >= 0 Then
                    cfactor(room) = a
                Else
                    txtCFactor.Focus()
                End If
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCFactor_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCFactor.Leave
        '*  =====================================================
        '*      This event checks for valid data when the control
        '*      loses the focus.
        '*  =====================================================

        Dim a As Single
        If IsNumeric(txtCFactor.Text) = True Then
            a = CSng(txtCFactor.Text)

            If a >= 0 Then
            Else
                MsgBox("Invalid Entry!")
                txtCFactor.Focus()
            End If
        End If
    End Sub

    'UPGRADE_WARNING: Event txtLinkActuation.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtLinkActuation_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLinkActuation.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================
        Dim room As Short
        room = lstRoomDetector.SelectedIndex + 1
        If room = 0 Then Exit Sub
        If IsNumeric(txtLinkActuation.Text) = True Then
            'store detector actuation temperature as a variable
            ActuationTemp(room) = CSng(txtLinkActuation.Text) + 273
        End If
    End Sub

    Private Sub txtLinkActuation_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLinkActuation.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*  ===================================================
        '*      This event checks for valid input when the user
        '*      pressing the enter key and if valid stores the
        '*      data as a variable.
        '*  ===================================================

        Dim a As Single
        Dim room As Short
        room = lstRoomDetector.SelectedIndex + 1
        If IsNumeric(txtLinkActuation.Text) = True Then
            a = CSng(txtLinkActuation.Text)

            If KeyAscii = 13 Then
                If a >= InteriorTemp - 273 Then
                    ActuationTemp(room) = a + 273
                Else
                    txtLinkActuation.Focus()
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtLinkActuation_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLinkActuation.Leave
        '*  =====================================================
        '*      This event checks for valid data when the control
        '*      loses the focus.
        '*  =====================================================

        Dim a As Single
        If IsNumeric(txtLinkActuation.Text) = True Then
            a = CSng(txtLinkActuation.Text)

            If a >= InteriorTemp - 273 Then
            Else
                MsgBox("Invalid Entry!")
                txtLinkActuation.Focus()
            End If
        End If
    End Sub


    'UPGRADE_WARNING: Event txtRadialDistance.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtRadialDistance_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRadialDistance.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================
        Dim room As Short
        room = lstRoomDetector.SelectedIndex + 1
        If room = 0 Then Exit Sub
        If IsNumeric(txtRadialDistance.Text) = True Then
            'store radial distance as a variable
            RadialDistance(room) = CSng(txtRadialDistance.Text)
        End If
    End Sub

    Private Sub txtRadialDistance_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRadialDistance.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*  ===================================================
        '*      This event checks for valid input when the user
        '*      pressing the enter key and if valid stores the
        '*      data as a variable.
        '*  ===================================================

        Dim a As Single
        Dim room As Short
        room = lstRoomDetector.SelectedIndex + 1
        If IsNumeric(txtRadialDistance.Text) = True Then
            a = CSng(txtRadialDistance.Text)

            If KeyAscii = 13 Then
                If a >= 0 Then
                    RadialDistance(room) = a
                Else
                    txtRadialDistance.Focus()
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtRadialDistance_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRadialDistance.Leave
        '*  =====================================================
        '*      This event checks for valid data when the control
        '*      loses the focus.
        '*  =====================================================

        Dim a As Single
        If IsNumeric(txtRadialDistance.Text) = True Then
            a = CSng(txtRadialDistance.Text)

            If a >= 0 Then
            Else
                MsgBox("Invalid Entry!")
                txtRadialDistance.Focus()
            End If
        End If
    End Sub

    'UPGRADE_WARNING: Event txtRTI.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtRTI_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRTI.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================
        Dim room As Short
        room = lstRoomDetector.SelectedIndex + 1
        If room = 0 Then Exit Sub
        'store response time index as a variable
        If IsNumeric(txtRTI.Text) = True Then
            RTI(room) = CSng(txtRTI.Text)
        End If
    End Sub

    Private Sub txtRTI_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRTI.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*  ===================================================
        '*      This event checks for valid input when the user
        '*      pressing the enter key and if valid stores the
        '*      data as a variable.
        '*  ===================================================

        Dim a As Single
        Dim room As Short
        room = lstRoomDetector.SelectedIndex + 1
        If IsNumeric(txtRTI.Text) = True Then
            a = CSng(txtRTI.Text)

            If KeyAscii = 13 Then
                If a > 0 Then
                    RTI(room) = a
                Else
                    txtRTI.Focus()
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtRTI_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRTI.Leave
        '*  =====================================================
        '*      This event checks for valid data when the control
        '*      loses the focus.
        '*  =====================================================

        Dim a As Single
        If IsNumeric(txtRTI.Text) = True Then
            a = CSng(txtRTI.Text)

            If a > 0 Then
            Else
                MsgBox("Invalid Entry!")
                txtRTI.Focus()
            End If
        End If
    End Sub


    'UPGRADE_WARNING: Event txtSprinkDistance.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtSprinkDistance_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSprinkDistance.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================
        Dim room As Short

        room = lstRoomDetector.SelectedIndex + 1
        If room = 0 Then Exit Sub
        If IsNumeric(txtSprinkDistance.Text) = True Then
            SprinkDistance(room) = CSng(txtSprinkDistance.Text)
        End If
    End Sub

    Private Sub txtSprinkDistance_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSprinkDistance.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*  ===================================================
        '*      This event checks for valid input when the user
        '*      pressing the enter key and if valid stores the
        '*      data as a variable.
        '*  ===================================================

        Dim a As Single
        Dim room As Short
        room = lstRoomDetector.SelectedIndex + 1
        If room = 0 Then GoTo EventExitSub
        If IsNumeric(txtSprinkDistance.Text) = True Then
            a = CSng(txtSprinkDistance.Text)
        End If
        If KeyAscii = 13 Then
            If a >= 0 And a <= RoomHeight(room) Then
                SprinkDistance(room) = a
            Else
                txtSprinkDistance.Focus()
            End If
        End If

EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtSprinkDistance_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSprinkDistance.Leave
        '*  =====================================================
        '*      This event checks for valid data when the control
        '*      loses the focus.
        '*  =====================================================

        Dim a As Single
        Dim room As Short
        room = lstRoomDetector.SelectedIndex + 1
        If room = 0 Then Exit Sub
        If IsNumeric(txtSprinkDistance.Text) = True Then
            a = CSng(txtSprinkDistance.Text)

            If a >= 0 And a <= RoomHeight(room) Then
            Else
                MsgBox("Invalid Entry!")
                txtSprinkDistance.Focus()
            End If
        End If
    End Sub

    'UPGRADE_WARNING: Event txtWaterSprayDensity.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtWaterSprayDensity_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWaterSprayDensity.TextChanged
        '*  ===================================================
        '*      This event stores the data input as a variable
        '*      when the contents of the text box are changed.
        '*  ===================================================
        Dim room As Short
        room = lstRoomDetector.SelectedIndex + 1
        If room = 0 Then Exit Sub
        If IsNumeric(txtWaterSprayDensity.Text) = True Then
            'store sprinkler C Factor as a variable
            WaterSprayDensity(room) = CSng(txtWaterSprayDensity.Text) / 60
        End If
    End Sub

    Private Sub txtWaterSprayDensity_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtWaterSprayDensity.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*  ===================================================
        '*      This event checks for valid input when the user
        '*      pressing the enter key and if valid stores the
        '*      data as a variable.
        '*  ===================================================

        Dim a As Single
        Dim room As Short
        room = lstRoomDetector.SelectedIndex + 1
        If IsNumeric(txtWaterSprayDensity.Text) = True Then
            a = CSng(txtWaterSprayDensity.Text)

            If KeyAscii = 13 Then
                If a >= 0 Then
                    WaterSprayDensity(room) = a / 60
                Else
                    txtCFactor.Focus()
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtWaterSprayDensity_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWaterSprayDensity.Leave
        '*  =====================================================
        '*      This event checks for valid data when the control
        '*      loses the focus.
        '*  =====================================================

        Dim a As Single
        If IsNumeric(txtWaterSprayDensity.Text) = True Then
            a = CSng(txtWaterSprayDensity.Text)

            If a >= 0 Then
            Else
                MsgBox("Invalid Entry!")
                txtWaterSprayDensity.Focus()
            End If
        Else
            MsgBox("Invalid Entry!")
            txtWaterSprayDensity.Focus()
        End If
    End Sub

End Class