<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmSmokeDetector
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
	Public WithEvents optDetSens As System.Windows.Forms.RadioButton
	Public WithEvents optAS1603 As System.Windows.Forms.RadioButton
	Public WithEvents optSpecifyOD As System.Windows.Forms.RadioButton
	Public WithEvents Frame3 As System.Windows.Forms.GroupBox
	Public WithEvents txtDetSensitivity As System.Windows.Forms.TextBox
	Public WithEvents chkUseODinside As System.Windows.Forms.CheckBox
	Public WithEvents txtSDdelay As System.Windows.Forms.TextBox
	Public WithEvents optSDvhigh As System.Windows.Forms.RadioButton
	Public WithEvents optSDhigh As System.Windows.Forms.RadioButton
	Public WithEvents optSDnormal As System.Windows.Forms.RadioButton
	Public WithEvents txtOpticalDensity As System.Windows.Forms.TextBox
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents lblOpticalDensity As System.Windows.Forms.Label
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents cmdClose As System.Windows.Forms.Button
	Public WithEvents chkSmokeDetector As System.Windows.Forms.CheckBox
	Public WithEvents lstRoomSD As System.Windows.Forms.ComboBox
	Public WithEvents lblRoomSD As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents txtSDdepth As System.Windows.Forms.TextBox
	Public WithEvents txtSDRadialDist As System.Windows.Forms.TextBox
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Frame4 As System.Windows.Forms.GroupBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.chkUseODinside = New System.Windows.Forms.CheckBox
        Me.Frame3 = New System.Windows.Forms.GroupBox
        Me.optDetSens = New System.Windows.Forms.RadioButton
        Me.optAS1603 = New System.Windows.Forms.RadioButton
        Me.optSpecifyOD = New System.Windows.Forms.RadioButton
        Me.Frame2 = New System.Windows.Forms.GroupBox
        Me.txtDetSensitivity = New System.Windows.Forms.TextBox
        Me.txtSDdelay = New System.Windows.Forms.TextBox
        Me.optSDvhigh = New System.Windows.Forms.RadioButton
        Me.optSDhigh = New System.Windows.Forms.RadioButton
        Me.optSDnormal = New System.Windows.Forms.RadioButton
        Me.txtOpticalDensity = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblOpticalDensity = New System.Windows.Forms.Label
        Me.cmdClose = New System.Windows.Forms.Button
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.chkSmokeDetector = New System.Windows.Forms.CheckBox
        Me.lstRoomSD = New System.Windows.Forms.ComboBox
        Me.lblRoomSD = New System.Windows.Forms.Label
        Me.Frame4 = New System.Windows.Forms.GroupBox
        Me.txtSDdepth = New System.Windows.Forms.TextBox
        Me.txtSDRadialDist = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Frame3.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.Frame4.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkUseODinside
        '
        Me.chkUseODinside.BackColor = System.Drawing.SystemColors.Control
        Me.chkUseODinside.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkUseODinside.Checked = True
        Me.chkUseODinside.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkUseODinside.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkUseODinside.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkUseODinside.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkUseODinside.Location = New System.Drawing.Point(40, 152)
        Me.chkUseODinside.Name = "chkUseODinside"
        Me.chkUseODinside.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkUseODinside.Size = New System.Drawing.Size(265, 17)
        Me.chkUseODinside.TabIndex = 21
        Me.chkUseODinside.Text = "Sensitivity criteria apply to inside detector chamber"
        Me.ToolTip1.SetToolTip(Me.chkUseODinside, "leave unchecked to apply to outside detector")
        Me.chkUseODinside.UseVisualStyleBackColor = False
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.optDetSens)
        Me.Frame3.Controls.Add(Me.optAS1603)
        Me.Frame3.Controls.Add(Me.optSpecifyOD)
        Me.Frame3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame3.Location = New System.Drawing.Point(184, 0)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(213, 113)
        Me.Frame3.TabIndex = 11
        Me.Frame3.TabStop = False
        '
        'optDetSens
        '
        Me.optDetSens.BackColor = System.Drawing.SystemColors.Control
        Me.optDetSens.Cursor = System.Windows.Forms.Cursors.Default
        Me.optDetSens.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optDetSens.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDetSens.Location = New System.Drawing.Point(16, 80)
        Me.optDetSens.Name = "optDetSens"
        Me.optDetSens.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optDetSens.Size = New System.Drawing.Size(194, 17)
        Me.optDetSens.TabIndex = 22
        Me.optDetSens.TabStop = True
        Me.optDetSens.Text = "Specify detector sensitivity"
        Me.optDetSens.UseVisualStyleBackColor = False
        '
        'optAS1603
        '
        Me.optAS1603.BackColor = System.Drawing.SystemColors.Control
        Me.optAS1603.Checked = True
        Me.optAS1603.Cursor = System.Windows.Forms.Cursors.Default
        Me.optAS1603.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optAS1603.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optAS1603.Location = New System.Drawing.Point(16, 48)
        Me.optAS1603.Name = "optAS1603"
        Me.optAS1603.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optAS1603.Size = New System.Drawing.Size(97, 25)
        Me.optAS1603.TabIndex = 13
        Me.optAS1603.TabStop = True
        Me.optAS1603.Text = "Use AS 1603.2"
        Me.optAS1603.UseVisualStyleBackColor = False
        '
        'optSpecifyOD
        '
        Me.optSpecifyOD.BackColor = System.Drawing.SystemColors.Control
        Me.optSpecifyOD.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSpecifyOD.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optSpecifyOD.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optSpecifyOD.Location = New System.Drawing.Point(16, 24)
        Me.optSpecifyOD.Name = "optSpecifyOD"
        Me.optSpecifyOD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSpecifyOD.Size = New System.Drawing.Size(121, 17)
        Me.optSpecifyOD.TabIndex = 12
        Me.optSpecifyOD.TabStop = True
        Me.optSpecifyOD.Text = "Specify Alarm OD"
        Me.optSpecifyOD.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.txtDetSensitivity)
        Me.Frame2.Controls.Add(Me.chkUseODinside)
        Me.Frame2.Controls.Add(Me.txtSDdelay)
        Me.Frame2.Controls.Add(Me.optSDvhigh)
        Me.Frame2.Controls.Add(Me.optSDhigh)
        Me.Frame2.Controls.Add(Me.optSDnormal)
        Me.Frame2.Controls.Add(Me.txtOpticalDensity)
        Me.Frame2.Controls.Add(Me.Label4)
        Me.Frame2.Controls.Add(Me.Label1)
        Me.Frame2.Controls.Add(Me.lblOpticalDensity)
        Me.Frame2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(8, 112)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(386, 177)
        Me.Frame2.TabIndex = 5
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Sensitivity"
        '
        'txtDetSensitivity
        '
        Me.txtDetSensitivity.AcceptsReturn = True
        Me.txtDetSensitivity.BackColor = System.Drawing.SystemColors.Window
        Me.txtDetSensitivity.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDetSensitivity.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDetSensitivity.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDetSensitivity.Location = New System.Drawing.Point(256, 88)
        Me.txtDetSensitivity.MaxLength = 0
        Me.txtDetSensitivity.Name = "txtDetSensitivity"
        Me.txtDetSensitivity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDetSensitivity.Size = New System.Drawing.Size(49, 19)
        Me.txtDetSensitivity.TabIndex = 23
        Me.txtDetSensitivity.Text = "2.5"
        Me.txtDetSensitivity.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtSDdelay
        '
        Me.txtSDdelay.AcceptsReturn = True
        Me.txtSDdelay.BackColor = System.Drawing.SystemColors.Window
        Me.txtSDdelay.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSDdelay.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSDdelay.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSDdelay.Location = New System.Drawing.Point(256, 120)
        Me.txtSDdelay.MaxLength = 0
        Me.txtSDdelay.Name = "txtSDdelay"
        Me.txtSDdelay.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSDdelay.Size = New System.Drawing.Size(49, 20)
        Me.txtSDdelay.TabIndex = 14
        Me.txtSDdelay.Text = "15"
        Me.txtSDdelay.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'optSDvhigh
        '
        Me.optSDvhigh.BackColor = System.Drawing.SystemColors.Control
        Me.optSDvhigh.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSDvhigh.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optSDvhigh.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optSDvhigh.Location = New System.Drawing.Point(224, 16)
        Me.optSDvhigh.Name = "optSDvhigh"
        Me.optSDvhigh.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSDvhigh.Size = New System.Drawing.Size(73, 33)
        Me.optSDvhigh.TabIndex = 10
        Me.optSDvhigh.TabStop = True
        Me.optSDvhigh.Text = "Very High"
        Me.optSDvhigh.UseVisualStyleBackColor = False
        '
        'optSDhigh
        '
        Me.optSDhigh.BackColor = System.Drawing.SystemColors.Control
        Me.optSDhigh.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSDhigh.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optSDhigh.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optSDhigh.Location = New System.Drawing.Point(128, 16)
        Me.optSDhigh.Name = "optSDhigh"
        Me.optSDhigh.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSDhigh.Size = New System.Drawing.Size(57, 33)
        Me.optSDhigh.TabIndex = 9
        Me.optSDhigh.TabStop = True
        Me.optSDhigh.Text = "High"
        Me.optSDhigh.UseVisualStyleBackColor = False
        '
        'optSDnormal
        '
        Me.optSDnormal.BackColor = System.Drawing.SystemColors.Control
        Me.optSDnormal.Checked = True
        Me.optSDnormal.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSDnormal.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optSDnormal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optSDnormal.Location = New System.Drawing.Point(24, 16)
        Me.optSDnormal.Name = "optSDnormal"
        Me.optSDnormal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSDnormal.Size = New System.Drawing.Size(73, 33)
        Me.optSDnormal.TabIndex = 8
        Me.optSDnormal.TabStop = True
        Me.optSDnormal.Text = "Normal"
        Me.optSDnormal.UseVisualStyleBackColor = False
        '
        'txtOpticalDensity
        '
        Me.txtOpticalDensity.AcceptsReturn = True
        Me.txtOpticalDensity.BackColor = System.Drawing.SystemColors.Window
        Me.txtOpticalDensity.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtOpticalDensity.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOpticalDensity.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtOpticalDensity.Location = New System.Drawing.Point(256, 56)
        Me.txtOpticalDensity.MaxLength = 0
        Me.txtOpticalDensity.Name = "txtOpticalDensity"
        Me.txtOpticalDensity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtOpticalDensity.Size = New System.Drawing.Size(49, 20)
        Me.txtOpticalDensity.TabIndex = 6
        Me.txtOpticalDensity.Text = "0.14"
        Me.txtOpticalDensity.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(48, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(193, 17)
        Me.Label4.TabIndex = 24
        Me.Label4.Text = "Detector Sensitivity (% per foot)"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 120)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(233, 17)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "Characteristic Length Number (m)"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblOpticalDensity
        '
        Me.lblOpticalDensity.BackColor = System.Drawing.SystemColors.Control
        Me.lblOpticalDensity.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblOpticalDensity.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOpticalDensity.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOpticalDensity.Location = New System.Drawing.Point(56, 56)
        Me.lblOpticalDensity.Name = "lblOpticalDensity"
        Me.lblOpticalDensity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblOpticalDensity.Size = New System.Drawing.Size(185, 17)
        Me.lblOpticalDensity.TabIndex = 7
        Me.lblOpticalDensity.Text = "Smoke Optical Density at Alarm (1/m)"
        Me.lblOpticalDensity.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Location = New System.Drawing.Point(302, 383)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(89, 25)
        Me.cmdClose.TabIndex = 3
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.chkSmokeDetector)
        Me.Frame1.Controls.Add(Me.lstRoomSD)
        Me.Frame1.Controls.Add(Me.lblRoomSD)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(8, 0)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(169, 113)
        Me.Frame1.TabIndex = 0
        Me.Frame1.TabStop = False
        '
        'chkSmokeDetector
        '
        Me.chkSmokeDetector.BackColor = System.Drawing.SystemColors.Control
        Me.chkSmokeDetector.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkSmokeDetector.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkSmokeDetector.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSmokeDetector.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkSmokeDetector.Location = New System.Drawing.Point(8, 56)
        Me.chkSmokeDetector.Name = "chkSmokeDetector"
        Me.chkSmokeDetector.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkSmokeDetector.Size = New System.Drawing.Size(145, 17)
        Me.chkSmokeDetector.TabIndex = 4
        Me.chkSmokeDetector.Text = "Smoke Detector Present"
        Me.chkSmokeDetector.UseVisualStyleBackColor = False
        '
        'lstRoomSD
        '
        Me.lstRoomSD.BackColor = System.Drawing.SystemColors.Window
        Me.lstRoomSD.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstRoomSD.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstRoomSD.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstRoomSD.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstRoomSD.Location = New System.Drawing.Point(72, 24)
        Me.lstRoomSD.Name = "lstRoomSD"
        Me.lstRoomSD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstRoomSD.Size = New System.Drawing.Size(81, 22)
        Me.lstRoomSD.TabIndex = 1
        '
        'lblRoomSD
        '
        Me.lblRoomSD.BackColor = System.Drawing.SystemColors.Control
        Me.lblRoomSD.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblRoomSD.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRoomSD.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRoomSD.Location = New System.Drawing.Point(8, 24)
        Me.lblRoomSD.Name = "lblRoomSD"
        Me.lblRoomSD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblRoomSD.Size = New System.Drawing.Size(49, 17)
        Me.lblRoomSD.TabIndex = 2
        Me.lblRoomSD.Text = "Room #"
        Me.lblRoomSD.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.txtSDdepth)
        Me.Frame4.Controls.Add(Me.txtSDRadialDist)
        Me.Frame4.Controls.Add(Me.Label3)
        Me.Frame4.Controls.Add(Me.Label2)
        Me.Frame4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(8, 288)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(383, 89)
        Me.Frame4.TabIndex = 16
        Me.Frame4.TabStop = False
        Me.Frame4.Text = "Location"
        Me.Frame4.Visible = False
        '
        'txtSDdepth
        '
        Me.txtSDdepth.AcceptsReturn = True
        Me.txtSDdepth.BackColor = System.Drawing.SystemColors.Window
        Me.txtSDdepth.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSDdepth.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSDdepth.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSDdepth.Location = New System.Drawing.Point(256, 56)
        Me.txtSDdepth.MaxLength = 0
        Me.txtSDdepth.Name = "txtSDdepth"
        Me.txtSDdepth.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSDdepth.Size = New System.Drawing.Size(49, 19)
        Me.txtSDdepth.TabIndex = 19
        Me.txtSDdepth.Text = "0.025"
        Me.txtSDdepth.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtSDRadialDist
        '
        Me.txtSDRadialDist.AcceptsReturn = True
        Me.txtSDRadialDist.BackColor = System.Drawing.SystemColors.Window
        Me.txtSDRadialDist.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSDRadialDist.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSDRadialDist.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSDRadialDist.Location = New System.Drawing.Point(256, 24)
        Me.txtSDRadialDist.MaxLength = 0
        Me.txtSDRadialDist.Name = "txtSDRadialDist"
        Me.txtSDRadialDist.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSDRadialDist.Size = New System.Drawing.Size(49, 19)
        Me.txtSDRadialDist.TabIndex = 17
        Me.txtSDRadialDist.Text = "3"
        Me.txtSDRadialDist.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(16, 56)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(217, 25)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "Distance Below Ceiling (m)"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(16, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(217, 17)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Radial Distance from Fire Plume Centreline (m)"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'frmSmokeDetector
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(409, 415)
        Me.ControlBox = False
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Frame4)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(2, 21)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSmokeDetector"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Optical Smoke Detector Settings"
        Me.Frame3.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.Frame4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class