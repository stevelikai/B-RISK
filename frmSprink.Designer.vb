<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmSprink
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents lstRoomDetector As System.Windows.Forms.ComboBox
	Public WithEvents cmdDefaults As System.Windows.Forms.Button
	Public WithEvents optSprinkOff As System.Windows.Forms.RadioButton
	Public WithEvents optControl As System.Windows.Forms.RadioButton
	Public WithEvents optSuppression As System.Windows.Forms.RadioButton
	Public WithEvents FraSprOpt As System.Windows.Forms.GroupBox
	Public WithEvents cmdHelp As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents txtSprinkDistance As System.Windows.Forms.TextBox
	Public WithEvents cboDetectorType As System.Windows.Forms.ComboBox
	Public WithEvents txtWaterSprayDensity As System.Windows.Forms.TextBox
	Public WithEvents txtRTI As System.Windows.Forms.TextBox
	Public WithEvents txtCFactor As System.Windows.Forms.TextBox
	Public WithEvents txtRadialDistance As System.Windows.Forms.TextBox
	Public WithEvents txtLinkActuation As System.Windows.Forms.TextBox
	Public WithEvents lblSprinkDistance As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents lblWaterSprayDensity As System.Windows.Forms.Label
	Public WithEvents lblRTI As System.Windows.Forms.Label
	Public WithEvents lblCFactor As System.Windows.Forms.Label
	Public WithEvents lblradialDistance As System.Windows.Forms.Label
	Public WithEvents lblLinkActuation As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdDefaults = New System.Windows.Forms.Button
        Me.lstRoomDetector = New System.Windows.Forms.ComboBox
        Me.FraSprOpt = New System.Windows.Forms.GroupBox
        Me.optSprinkOff = New System.Windows.Forms.RadioButton
        Me.optControl = New System.Windows.Forms.RadioButton
        Me.optSuppression = New System.Windows.Forms.RadioButton
        Me.cmdHelp = New System.Windows.Forms.Button
        Me.cmdOK = New System.Windows.Forms.Button
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.txtSprinkDistance = New System.Windows.Forms.TextBox
        Me.cboDetectorType = New System.Windows.Forms.ComboBox
        Me.txtWaterSprayDensity = New System.Windows.Forms.TextBox
        Me.txtRTI = New System.Windows.Forms.TextBox
        Me.txtCFactor = New System.Windows.Forms.TextBox
        Me.txtRadialDistance = New System.Windows.Forms.TextBox
        Me.txtLinkActuation = New System.Windows.Forms.TextBox
        Me.lblSprinkDistance = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.lblWaterSprayDensity = New System.Windows.Forms.Label
        Me.lblRTI = New System.Windows.Forms.Label
        Me.lblCFactor = New System.Windows.Forms.Label
        Me.lblradialDistance = New System.Windows.Forms.Label
        Me.lblLinkActuation = New System.Windows.Forms.Label
        Me.FraSprOpt.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdDefaults
        '
        Me.cmdDefaults.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDefaults.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDefaults.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDefaults.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDefaults.Location = New System.Drawing.Point(371, 200)
        Me.cmdDefaults.Name = "cmdDefaults"
        Me.cmdDefaults.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDefaults.Size = New System.Drawing.Size(105, 25)
        Me.cmdDefaults.TabIndex = 11
        Me.cmdDefaults.Text = "Reset to Default"
        Me.ToolTip1.SetToolTip(Me.cmdDefaults, "Use default values for the currently selected detector type")
        Me.cmdDefaults.UseVisualStyleBackColor = False
        '
        'lstRoomDetector
        '
        Me.lstRoomDetector.BackColor = System.Drawing.SystemColors.Window
        Me.lstRoomDetector.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstRoomDetector.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstRoomDetector.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstRoomDetector.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstRoomDetector.Location = New System.Drawing.Point(80, 24)
        Me.lstRoomDetector.Name = "lstRoomDetector"
        Me.lstRoomDetector.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstRoomDetector.Size = New System.Drawing.Size(81, 22)
        Me.lstRoomDetector.TabIndex = 1
        '
        'FraSprOpt
        '
        Me.FraSprOpt.BackColor = System.Drawing.SystemColors.Control
        Me.FraSprOpt.Controls.Add(Me.optSprinkOff)
        Me.FraSprOpt.Controls.Add(Me.optControl)
        Me.FraSprOpt.Controls.Add(Me.optSuppression)
        Me.FraSprOpt.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FraSprOpt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.FraSprOpt.Location = New System.Drawing.Point(352, 8)
        Me.FraSprOpt.Name = "FraSprOpt"
        Me.FraSprOpt.Padding = New System.Windows.Forms.Padding(0)
        Me.FraSprOpt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.FraSprOpt.Size = New System.Drawing.Size(124, 137)
        Me.FraSprOpt.TabIndex = 19
        Me.FraSprOpt.TabStop = False
        Me.FraSprOpt.Text = "Sprinkler Options"
        '
        'optSprinkOff
        '
        Me.optSprinkOff.BackColor = System.Drawing.SystemColors.Control
        Me.optSprinkOff.Checked = True
        Me.optSprinkOff.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSprinkOff.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optSprinkOff.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optSprinkOff.Location = New System.Drawing.Point(16, 95)
        Me.optSprinkOff.Name = "optSprinkOff"
        Me.optSprinkOff.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSprinkOff.Size = New System.Drawing.Size(81, 34)
        Me.optSprinkOff.TabIndex = 10
        Me.optSprinkOff.TabStop = True
        Me.optSprinkOff.Text = "Off"
        Me.optSprinkOff.UseVisualStyleBackColor = False
        '
        'optControl
        '
        Me.optControl.BackColor = System.Drawing.SystemColors.Control
        Me.optControl.Cursor = System.Windows.Forms.Cursors.Default
        Me.optControl.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optControl.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optControl.Location = New System.Drawing.Point(16, 64)
        Me.optControl.Name = "optControl"
        Me.optControl.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optControl.Size = New System.Drawing.Size(73, 25)
        Me.optControl.TabIndex = 9
        Me.optControl.TabStop = True
        Me.optControl.Text = "Control"
        Me.optControl.UseVisualStyleBackColor = False
        '
        'optSuppression
        '
        Me.optSuppression.BackColor = System.Drawing.SystemColors.Control
        Me.optSuppression.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSuppression.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optSuppression.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optSuppression.Location = New System.Drawing.Point(16, 32)
        Me.optSuppression.Name = "optSuppression"
        Me.optSuppression.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSuppression.Size = New System.Drawing.Size(105, 26)
        Me.optSuppression.TabIndex = 8
        Me.optSuppression.TabStop = True
        Me.optSuppression.Text = "Suppression"
        Me.optSuppression.UseVisualStyleBackColor = False
        '
        'cmdHelp
        '
        Me.cmdHelp.BackColor = System.Drawing.SystemColors.Control
        Me.cmdHelp.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdHelp.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdHelp.Location = New System.Drawing.Point(411, 160)
        Me.cmdHelp.Name = "cmdHelp"
        Me.cmdHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdHelp.Size = New System.Drawing.Size(65, 25)
        Me.cmdHelp.TabIndex = 13
        Me.cmdHelp.Text = "Help"
        Me.cmdHelp.UseVisualStyleBackColor = False
        Me.cmdHelp.Visible = False
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Location = New System.Drawing.Point(411, 235)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(65, 25)
        Me.cmdOK.TabIndex = 12
        Me.cmdOK.Text = "OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.txtSprinkDistance)
        Me.Frame1.Controls.Add(Me.cboDetectorType)
        Me.Frame1.Controls.Add(Me.txtWaterSprayDensity)
        Me.Frame1.Controls.Add(Me.txtRTI)
        Me.Frame1.Controls.Add(Me.txtCFactor)
        Me.Frame1.Controls.Add(Me.txtRadialDistance)
        Me.Frame1.Controls.Add(Me.txtLinkActuation)
        Me.Frame1.Controls.Add(Me.lblSprinkDistance)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me.lblWaterSprayDensity)
        Me.Frame1.Controls.Add(Me.lblRTI)
        Me.Frame1.Controls.Add(Me.lblCFactor)
        Me.Frame1.Controls.Add(Me.lblradialDistance)
        Me.Frame1.Controls.Add(Me.lblLinkActuation)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(8, 8)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(337, 257)
        Me.Frame1.TabIndex = 0
        Me.Frame1.TabStop = False
        '
        'txtSprinkDistance
        '
        Me.txtSprinkDistance.AcceptsReturn = True
        Me.txtSprinkDistance.BackColor = System.Drawing.SystemColors.Window
        Me.txtSprinkDistance.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSprinkDistance.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSprinkDistance.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSprinkDistance.Location = New System.Drawing.Point(256, 216)
        Me.txtSprinkDistance.MaxLength = 0
        Me.txtSprinkDistance.Name = "txtSprinkDistance"
        Me.txtSprinkDistance.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSprinkDistance.Size = New System.Drawing.Size(57, 20)
        Me.txtSprinkDistance.TabIndex = 21
        Me.txtSprinkDistance.Text = "0.025"
        '
        'cboDetectorType
        '
        Me.cboDetectorType.BackColor = System.Drawing.SystemColors.Window
        Me.cboDetectorType.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboDetectorType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDetectorType.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDetectorType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboDetectorType.Location = New System.Drawing.Point(160, 16)
        Me.cboDetectorType.Name = "cboDetectorType"
        Me.cboDetectorType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboDetectorType.Size = New System.Drawing.Size(157, 22)
        Me.cboDetectorType.TabIndex = 2
        '
        'txtWaterSprayDensity
        '
        Me.txtWaterSprayDensity.AcceptsReturn = True
        Me.txtWaterSprayDensity.BackColor = System.Drawing.SystemColors.Window
        Me.txtWaterSprayDensity.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWaterSprayDensity.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWaterSprayDensity.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWaterSprayDensity.Location = New System.Drawing.Point(256, 56)
        Me.txtWaterSprayDensity.MaxLength = 0
        Me.txtWaterSprayDensity.Name = "txtWaterSprayDensity"
        Me.txtWaterSprayDensity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWaterSprayDensity.Size = New System.Drawing.Size(57, 20)
        Me.txtWaterSprayDensity.TabIndex = 3
        Me.txtWaterSprayDensity.Text = "0"
        '
        'txtRTI
        '
        Me.txtRTI.AcceptsReturn = True
        Me.txtRTI.BackColor = System.Drawing.Color.White
        Me.txtRTI.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRTI.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRTI.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRTI.Location = New System.Drawing.Point(256, 88)
        Me.txtRTI.MaxLength = 0
        Me.txtRTI.Name = "txtRTI"
        Me.txtRTI.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRTI.Size = New System.Drawing.Size(57, 20)
        Me.txtRTI.TabIndex = 4
        Me.txtRTI.Text = "165"
        '
        'txtCFactor
        '
        Me.txtCFactor.AcceptsReturn = True
        Me.txtCFactor.BackColor = System.Drawing.Color.White
        Me.txtCFactor.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCFactor.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCFactor.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCFactor.Location = New System.Drawing.Point(256, 120)
        Me.txtCFactor.MaxLength = 0
        Me.txtCFactor.Name = "txtCFactor"
        Me.txtCFactor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCFactor.Size = New System.Drawing.Size(57, 20)
        Me.txtCFactor.TabIndex = 5
        Me.txtCFactor.Text = "0"
        '
        'txtRadialDistance
        '
        Me.txtRadialDistance.AcceptsReturn = True
        Me.txtRadialDistance.BackColor = System.Drawing.Color.White
        Me.txtRadialDistance.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRadialDistance.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRadialDistance.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRadialDistance.Location = New System.Drawing.Point(256, 152)
        Me.txtRadialDistance.MaxLength = 0
        Me.txtRadialDistance.Name = "txtRadialDistance"
        Me.txtRadialDistance.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRadialDistance.Size = New System.Drawing.Size(57, 20)
        Me.txtRadialDistance.TabIndex = 6
        Me.txtRadialDistance.Text = "2.16"
        '
        'txtLinkActuation
        '
        Me.txtLinkActuation.AcceptsReturn = True
        Me.txtLinkActuation.BackColor = System.Drawing.Color.White
        Me.txtLinkActuation.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLinkActuation.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLinkActuation.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLinkActuation.Location = New System.Drawing.Point(256, 184)
        Me.txtLinkActuation.MaxLength = 0
        Me.txtLinkActuation.Name = "txtLinkActuation"
        Me.txtLinkActuation.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLinkActuation.Size = New System.Drawing.Size(57, 20)
        Me.txtLinkActuation.TabIndex = 7
        Me.txtLinkActuation.Text = "141"
        '
        'lblSprinkDistance
        '
        Me.lblSprinkDistance.BackColor = System.Drawing.SystemColors.Control
        Me.lblSprinkDistance.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSprinkDistance.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSprinkDistance.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSprinkDistance.Location = New System.Drawing.Point(96, 216)
        Me.lblSprinkDistance.Name = "lblSprinkDistance"
        Me.lblSprinkDistance.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSprinkDistance.Size = New System.Drawing.Size(145, 17)
        Me.lblSprinkDistance.TabIndex = 22
        Me.lblSprinkDistance.Text = "Distance Below Ceiling (m)"
        Me.lblSprinkDistance.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(8, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(49, 17)
        Me.Label2.TabIndex = 20
        Me.Label2.Text = "Room #"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblWaterSprayDensity
        '
        Me.lblWaterSprayDensity.BackColor = System.Drawing.SystemColors.Control
        Me.lblWaterSprayDensity.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblWaterSprayDensity.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWaterSprayDensity.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWaterSprayDensity.Location = New System.Drawing.Point(80, 56)
        Me.lblWaterSprayDensity.Name = "lblWaterSprayDensity"
        Me.lblWaterSprayDensity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblWaterSprayDensity.Size = New System.Drawing.Size(161, 17)
        Me.lblWaterSprayDensity.TabIndex = 18
        Me.lblWaterSprayDensity.Text = "Water Spray Density (mm/min)"
        Me.lblWaterSprayDensity.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblRTI
        '
        Me.lblRTI.BackColor = System.Drawing.Color.Transparent
        Me.lblRTI.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblRTI.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRTI.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblRTI.Location = New System.Drawing.Point(40, 88)
        Me.lblRTI.Name = "lblRTI"
        Me.lblRTI.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblRTI.Size = New System.Drawing.Size(201, 25)
        Me.lblRTI.TabIndex = 17
        Me.lblRTI.Text = "Response Time Index (m.s)^1/2"
        Me.lblRTI.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblCFactor
        '
        Me.lblCFactor.BackColor = System.Drawing.Color.Transparent
        Me.lblCFactor.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCFactor.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCFactor.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblCFactor.Location = New System.Drawing.Point(48, 120)
        Me.lblCFactor.Name = "lblCFactor"
        Me.lblCFactor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCFactor.Size = New System.Drawing.Size(193, 25)
        Me.lblCFactor.TabIndex = 16
        Me.lblCFactor.Text = "Sprinkler C-Factor (m/s)^1/2"
        Me.lblCFactor.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblradialDistance
        '
        Me.lblradialDistance.BackColor = System.Drawing.Color.Transparent
        Me.lblradialDistance.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblradialDistance.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblradialDistance.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblradialDistance.Location = New System.Drawing.Point(48, 152)
        Me.lblradialDistance.Name = "lblradialDistance"
        Me.lblradialDistance.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblradialDistance.Size = New System.Drawing.Size(193, 25)
        Me.lblradialDistance.TabIndex = 15
        Me.lblradialDistance.Text = "Radial Distance - Fire to Detector (m)"
        Me.lblradialDistance.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblLinkActuation
        '
        Me.lblLinkActuation.BackColor = System.Drawing.Color.Transparent
        Me.lblLinkActuation.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblLinkActuation.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLinkActuation.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblLinkActuation.Location = New System.Drawing.Point(16, 184)
        Me.lblLinkActuation.Name = "lblLinkActuation"
        Me.lblLinkActuation.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblLinkActuation.Size = New System.Drawing.Size(225, 33)
        Me.lblLinkActuation.TabIndex = 14
        Me.lblLinkActuation.Text = "Actuation Temperature (C)"
        Me.lblLinkActuation.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'frmSprink
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(488, 272)
        Me.Controls.Add(Me.lstRoomDetector)
        Me.Controls.Add(Me.cmdDefaults)
        Me.Controls.Add(Me.FraSprOpt)
        Me.Controls.Add(Me.cmdHelp)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(3, 18)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSprink"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detector / Sprinkler Settings"
        Me.FraSprOpt.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class