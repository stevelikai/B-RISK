<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmExtract
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
	Public WithEvents chkUseFanCurve As System.Windows.Forms.CheckBox
	Public WithEvents optPressurise As System.Windows.Forms.RadioButton
	Public WithEvents optExtract As System.Windows.Forms.RadioButton
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents lstRoomID As System.Windows.Forms.ComboBox
	Public WithEvents optFanAuto As System.Windows.Forms.RadioButton
	Public WithEvents optFanManual As System.Windows.Forms.RadioButton
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents optFanOff As System.Windows.Forms.RadioButton
	Public WithEvents optFanOn As System.Windows.Forms.RadioButton
	Public WithEvents fraCeilingFans As System.Windows.Forms.GroupBox
	Public WithEvents txtNumberFans As System.Windows.Forms.TextBox
	Public WithEvents txtFanElevation As System.Windows.Forms.TextBox
	Public WithEvents txtMaxPressure As System.Windows.Forms.TextBox
	Public WithEvents txtOperationTime As System.Windows.Forms.TextBox
	Public WithEvents txtCeilingExtractRate As System.Windows.Forms.TextBox
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents lblOperationTime As System.Windows.Forms.Label
	Public WithEvents lblCeilingExtractRate As System.Windows.Forms.Label
	Public WithEvents fraCeilingExtract As System.Windows.Forms.GroupBox
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.chkUseFanCurve = New System.Windows.Forms.CheckBox
        Me.Frame2 = New System.Windows.Forms.GroupBox
        Me.optPressurise = New System.Windows.Forms.RadioButton
        Me.optExtract = New System.Windows.Forms.RadioButton
        Me.lstRoomID = New System.Windows.Forms.ComboBox
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.optFanAuto = New System.Windows.Forms.RadioButton
        Me.optFanManual = New System.Windows.Forms.RadioButton
        Me.cmdOK = New System.Windows.Forms.Button
        Me.fraCeilingFans = New System.Windows.Forms.GroupBox
        Me.optFanOff = New System.Windows.Forms.RadioButton
        Me.optFanOn = New System.Windows.Forms.RadioButton
        Me.fraCeilingExtract = New System.Windows.Forms.GroupBox
        Me.txtNumberFans = New System.Windows.Forms.TextBox
        Me.txtFanElevation = New System.Windows.Forms.TextBox
        Me.txtMaxPressure = New System.Windows.Forms.TextBox
        Me.txtOperationTime = New System.Windows.Forms.TextBox
        Me.txtCeilingExtractRate = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.lblOperationTime = New System.Windows.Forms.Label
        Me.lblCeilingExtractRate = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.fraCeilingFans.SuspendLayout()
        Me.fraCeilingExtract.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkUseFanCurve
        '
        Me.chkUseFanCurve.BackColor = System.Drawing.SystemColors.Control
        Me.chkUseFanCurve.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkUseFanCurve.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkUseFanCurve.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkUseFanCurve.Location = New System.Drawing.Point(17, 88)
        Me.chkUseFanCurve.Name = "chkUseFanCurve"
        Me.chkUseFanCurve.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkUseFanCurve.Size = New System.Drawing.Size(117, 20)
        Me.chkUseFanCurve.TabIndex = 17
        Me.chkUseFanCurve.Text = "Use Fan Curve?"
        Me.chkUseFanCurve.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.chkUseFanCurve)
        Me.Frame2.Controls.Add(Me.optPressurise)
        Me.Frame2.Controls.Add(Me.optExtract)
        Me.Frame2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(238, 120)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(139, 121)
        Me.Frame2.TabIndex = 16
        Me.Frame2.TabStop = False
        '
        'optPressurise
        '
        Me.optPressurise.BackColor = System.Drawing.SystemColors.Control
        Me.optPressurise.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPressurise.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optPressurise.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPressurise.Location = New System.Drawing.Point(24, 48)
        Me.optPressurise.Name = "optPressurise"
        Me.optPressurise.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPressurise.Size = New System.Drawing.Size(97, 25)
        Me.optPressurise.TabIndex = 8
        Me.optPressurise.TabStop = True
        Me.optPressurise.Text = "Pressurise"
        Me.optPressurise.UseVisualStyleBackColor = False
        '
        'optExtract
        '
        Me.optExtract.BackColor = System.Drawing.SystemColors.Control
        Me.optExtract.Checked = True
        Me.optExtract.Cursor = System.Windows.Forms.Cursors.Default
        Me.optExtract.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optExtract.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optExtract.Location = New System.Drawing.Point(24, 24)
        Me.optExtract.Name = "optExtract"
        Me.optExtract.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optExtract.Size = New System.Drawing.Size(97, 17)
        Me.optExtract.TabIndex = 7
        Me.optExtract.TabStop = True
        Me.optExtract.Text = "Extract"
        Me.optExtract.UseVisualStyleBackColor = False
        '
        'lstRoomID
        '
        Me.lstRoomID.BackColor = System.Drawing.SystemColors.Window
        Me.lstRoomID.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstRoomID.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstRoomID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstRoomID.Location = New System.Drawing.Point(72, 16)
        Me.lstRoomID.Name = "lstRoomID"
        Me.lstRoomID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstRoomID.Size = New System.Drawing.Size(97, 22)
        Me.lstRoomID.TabIndex = 9
        Me.lstRoomID.Text = "Combo1"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.optFanAuto)
        Me.Frame1.Controls.Add(Me.optFanManual)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(192, 56)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(185, 57)
        Me.Frame1.TabIndex = 14
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Actuation"
        '
        'optFanAuto
        '
        Me.optFanAuto.BackColor = System.Drawing.SystemColors.Control
        Me.optFanAuto.Cursor = System.Windows.Forms.Cursors.Default
        Me.optFanAuto.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optFanAuto.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optFanAuto.Location = New System.Drawing.Point(99, 24)
        Me.optFanAuto.Name = "optFanAuto"
        Me.optFanAuto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optFanAuto.Size = New System.Drawing.Size(81, 17)
        Me.optFanAuto.TabIndex = 4
        Me.optFanAuto.TabStop = True
        Me.optFanAuto.Text = "Automatic"
        Me.optFanAuto.UseVisualStyleBackColor = False
        '
        'optFanManual
        '
        Me.optFanManual.BackColor = System.Drawing.SystemColors.Control
        Me.optFanManual.Checked = True
        Me.optFanManual.Cursor = System.Windows.Forms.Cursors.Default
        Me.optFanManual.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optFanManual.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optFanManual.Location = New System.Drawing.Point(16, 24)
        Me.optFanManual.Name = "optFanManual"
        Me.optFanManual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optFanManual.Size = New System.Drawing.Size(67, 17)
        Me.optFanManual.TabIndex = 3
        Me.optFanManual.TabStop = True
        Me.optFanManual.Text = "Manual"
        Me.optFanManual.UseVisualStyleBackColor = False
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Location = New System.Drawing.Point(312, 280)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(65, 25)
        Me.cmdOK.TabIndex = 10
        Me.cmdOK.Text = "OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'fraCeilingFans
        '
        Me.fraCeilingFans.BackColor = System.Drawing.SystemColors.Control
        Me.fraCeilingFans.Controls.Add(Me.optFanOff)
        Me.fraCeilingFans.Controls.Add(Me.optFanOn)
        Me.fraCeilingFans.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraCeilingFans.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraCeilingFans.Location = New System.Drawing.Point(8, 56)
        Me.fraCeilingFans.Name = "fraCeilingFans"
        Me.fraCeilingFans.Padding = New System.Windows.Forms.Padding(0)
        Me.fraCeilingFans.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraCeilingFans.Size = New System.Drawing.Size(169, 57)
        Me.fraCeilingFans.TabIndex = 13
        Me.fraCeilingFans.TabStop = False
        Me.fraCeilingFans.Text = "Extract Fan"
        '
        'optFanOff
        '
        Me.optFanOff.BackColor = System.Drawing.SystemColors.Control
        Me.optFanOff.Checked = True
        Me.optFanOff.Cursor = System.Windows.Forms.Cursors.Default
        Me.optFanOff.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optFanOff.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optFanOff.Location = New System.Drawing.Point(104, 24)
        Me.optFanOff.Name = "optFanOff"
        Me.optFanOff.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optFanOff.Size = New System.Drawing.Size(49, 17)
        Me.optFanOff.TabIndex = 2
        Me.optFanOff.TabStop = True
        Me.optFanOff.Text = "Off"
        Me.optFanOff.UseVisualStyleBackColor = False
        '
        'optFanOn
        '
        Me.optFanOn.BackColor = System.Drawing.SystemColors.Control
        Me.optFanOn.Cursor = System.Windows.Forms.Cursors.Default
        Me.optFanOn.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optFanOn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optFanOn.Location = New System.Drawing.Point(32, 24)
        Me.optFanOn.Name = "optFanOn"
        Me.optFanOn.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optFanOn.Size = New System.Drawing.Size(73, 17)
        Me.optFanOn.TabIndex = 1
        Me.optFanOn.TabStop = True
        Me.optFanOn.Text = "On"
        Me.optFanOn.UseVisualStyleBackColor = False
        '
        'fraCeilingExtract
        '
        Me.fraCeilingExtract.BackColor = System.Drawing.SystemColors.Control
        Me.fraCeilingExtract.Controls.Add(Me.txtNumberFans)
        Me.fraCeilingExtract.Controls.Add(Me.txtFanElevation)
        Me.fraCeilingExtract.Controls.Add(Me.txtMaxPressure)
        Me.fraCeilingExtract.Controls.Add(Me.txtOperationTime)
        Me.fraCeilingExtract.Controls.Add(Me.txtCeilingExtractRate)
        Me.fraCeilingExtract.Controls.Add(Me.Label5)
        Me.fraCeilingExtract.Controls.Add(Me.Label3)
        Me.fraCeilingExtract.Controls.Add(Me.Label2)
        Me.fraCeilingExtract.Controls.Add(Me.lblOperationTime)
        Me.fraCeilingExtract.Controls.Add(Me.lblCeilingExtractRate)
        Me.fraCeilingExtract.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraCeilingExtract.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraCeilingExtract.Location = New System.Drawing.Point(11, 120)
        Me.fraCeilingExtract.Name = "fraCeilingExtract"
        Me.fraCeilingExtract.Padding = New System.Windows.Forms.Padding(0)
        Me.fraCeilingExtract.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraCeilingExtract.Size = New System.Drawing.Size(221, 185)
        Me.fraCeilingExtract.TabIndex = 0
        Me.fraCeilingExtract.TabStop = False
        '
        'txtNumberFans
        '
        Me.txtNumberFans.AcceptsReturn = True
        Me.txtNumberFans.BackColor = System.Drawing.SystemColors.Window
        Me.txtNumberFans.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNumberFans.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNumberFans.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNumberFans.Location = New System.Drawing.Point(167, 55)
        Me.txtNumberFans.MaxLength = 0
        Me.txtNumberFans.Name = "txtNumberFans"
        Me.txtNumberFans.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNumberFans.Size = New System.Drawing.Size(40, 20)
        Me.txtNumberFans.TabIndex = 23
        Me.txtNumberFans.Text = "1"
        '
        'txtFanElevation
        '
        Me.txtFanElevation.AcceptsReturn = True
        Me.txtFanElevation.BackColor = System.Drawing.SystemColors.Window
        Me.txtFanElevation.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFanElevation.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFanElevation.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFanElevation.Location = New System.Drawing.Point(167, 151)
        Me.txtFanElevation.MaxLength = 0
        Me.txtFanElevation.Name = "txtFanElevation"
        Me.txtFanElevation.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFanElevation.Size = New System.Drawing.Size(40, 20)
        Me.txtFanElevation.TabIndex = 20
        Me.txtFanElevation.Text = "0"
        '
        'txtMaxPressure
        '
        Me.txtMaxPressure.AcceptsReturn = True
        Me.txtMaxPressure.BackColor = System.Drawing.SystemColors.Window
        Me.txtMaxPressure.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMaxPressure.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaxPressure.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMaxPressure.Location = New System.Drawing.Point(167, 119)
        Me.txtMaxPressure.MaxLength = 0
        Me.txtMaxPressure.Name = "txtMaxPressure"
        Me.txtMaxPressure.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMaxPressure.Size = New System.Drawing.Size(40, 20)
        Me.txtMaxPressure.TabIndex = 18
        Me.txtMaxPressure.Text = "50"
        '
        'txtOperationTime
        '
        Me.txtOperationTime.AcceptsReturn = True
        Me.txtOperationTime.BackColor = System.Drawing.SystemColors.Window
        Me.txtOperationTime.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtOperationTime.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOperationTime.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtOperationTime.Location = New System.Drawing.Point(167, 87)
        Me.txtOperationTime.MaxLength = 0
        Me.txtOperationTime.Name = "txtOperationTime"
        Me.txtOperationTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtOperationTime.Size = New System.Drawing.Size(40, 20)
        Me.txtOperationTime.TabIndex = 6
        Me.txtOperationTime.Text = "0"
        '
        'txtCeilingExtractRate
        '
        Me.txtCeilingExtractRate.AcceptsReturn = True
        Me.txtCeilingExtractRate.BackColor = System.Drawing.SystemColors.Window
        Me.txtCeilingExtractRate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCeilingExtractRate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCeilingExtractRate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCeilingExtractRate.Location = New System.Drawing.Point(167, 23)
        Me.txtCeilingExtractRate.MaxLength = 0
        Me.txtCeilingExtractRate.Name = "txtCeilingExtractRate"
        Me.txtCeilingExtractRate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCeilingExtractRate.Size = New System.Drawing.Size(40, 20)
        Me.txtCeilingExtractRate.TabIndex = 5
        Me.txtCeilingExtractRate.Text = "0"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(54, 55)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(97, 17)
        Me.Label5.TabIndex = 24
        Me.Label5.Text = "Number of fans"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(38, 151)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(113, 17)
        Me.Label3.TabIndex = 21
        Me.Label3.Text = "Fan Elevation (m)"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(3, 119)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(155, 25)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Cross-Fan Pressure Limit (Pa)"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblOperationTime
        '
        Me.lblOperationTime.BackColor = System.Drawing.SystemColors.Control
        Me.lblOperationTime.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblOperationTime.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOperationTime.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOperationTime.Location = New System.Drawing.Point(70, 87)
        Me.lblOperationTime.Name = "lblOperationTime"
        Me.lblOperationTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblOperationTime.Size = New System.Drawing.Size(81, 17)
        Me.lblOperationTime.TabIndex = 12
        Me.lblOperationTime.Text = "Start time (sec)"
        Me.lblOperationTime.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblCeilingExtractRate
        '
        Me.lblCeilingExtractRate.BackColor = System.Drawing.SystemColors.Control
        Me.lblCeilingExtractRate.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCeilingExtractRate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCeilingExtractRate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCeilingExtractRate.Location = New System.Drawing.Point(26, 23)
        Me.lblCeilingExtractRate.Name = "lblCeilingExtractRate"
        Me.lblCeilingExtractRate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCeilingExtractRate.Size = New System.Drawing.Size(130, 25)
        Me.lblCeilingExtractRate.TabIndex = 11
        Me.lblCeilingExtractRate.Text = "Flow Rate per fan (m3/s)"
        Me.lblCeilingExtractRate.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(175, 19)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(113, 22)
        Me.Label4.TabIndex = 22
        Me.Label4.Text = "(Select room first)"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(49, 17)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "Room #"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'frmExtract
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(396, 310)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.lstRoomID)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.fraCeilingFans)
        Me.Controls.Add(Me.fraCeilingExtract)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmExtract"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Mechanical Extract"
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.fraCeilingFans.ResumeLayout(False)
        Me.fraCeilingExtract.ResumeLayout(False)
        Me.fraCeilingExtract.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class