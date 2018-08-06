<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmGlassBreak
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
	Public WithEvents cmdClose As System.Windows.Forms.Button
	Public WithEvents txtGlassDistance As System.Windows.Forms.TextBox
	Public WithEvents optGlassWithFlame As System.Windows.Forms.RadioButton
	Public WithEvents optGlassNoFlameFlux As System.Windows.Forms.RadioButton
	Public WithEvents Label17 As System.Windows.Forms.Label
	Public WithEvents Label11 As System.Windows.Forms.Label
	Public WithEvents fraHeatFlux As System.Windows.Forms.GroupBox
	Public WithEvents txtGlassFalloutTime As System.Windows.Forms.TextBox
	Public WithEvents txtGLASSexpansion As System.Windows.Forms.TextBox
	Public WithEvents txtGLASSshading As System.Windows.Forms.TextBox
	Public WithEvents txtGLASSbreakingstress As System.Windows.Forms.TextBox
	Public WithEvents txtGLASSYoungsModulus As System.Windows.Forms.TextBox
	Public WithEvents txtGLASSthickness As System.Windows.Forms.TextBox
	Public WithEvents txtGLASSalpha As System.Windows.Forms.TextBox
	Public WithEvents txtGLASSconductivity As System.Windows.Forms.TextBox
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents lblFalloutTime As System.Windows.Forms.Label
	Public WithEvents Label16 As System.Windows.Forms.Label
	Public WithEvents Label15 As System.Windows.Forms.Label
	Public WithEvents Label14 As System.Windows.Forms.Label
	Public WithEvents Label13 As System.Windows.Forms.Label
	Public WithEvents Label12 As System.Windows.Forms.Label
	Public WithEvents Label10 As System.Windows.Forms.Label
	Public WithEvents Label9 As System.Windows.Forms.Label
	Public WithEvents Label8 As System.Windows.Forms.Label
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents fraGlassProps As System.Windows.Forms.GroupBox
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdClose = New System.Windows.Forms.Button
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.fraHeatFlux = New System.Windows.Forms.GroupBox
        Me.txtGlassDistance = New System.Windows.Forms.TextBox
        Me.optGlassWithFlame = New System.Windows.Forms.RadioButton
        Me.optGlassNoFlameFlux = New System.Windows.Forms.RadioButton
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.fraGlassProps = New System.Windows.Forms.GroupBox
        Me.txtGlassFalloutTime = New System.Windows.Forms.TextBox
        Me.txtGLASSexpansion = New System.Windows.Forms.TextBox
        Me.txtGLASSshading = New System.Windows.Forms.TextBox
        Me.txtGLASSbreakingstress = New System.Windows.Forms.TextBox
        Me.txtGLASSYoungsModulus = New System.Windows.Forms.TextBox
        Me.txtGLASSthickness = New System.Windows.Forms.TextBox
        Me.txtGLASSalpha = New System.Windows.Forms.TextBox
        Me.txtGLASSconductivity = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.lblFalloutTime = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.Frame1.SuspendLayout()
        Me.fraHeatFlux.SuspendLayout()
        Me.fraGlassProps.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Location = New System.Drawing.Point(263, 333)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(81, 33)
        Me.cmdClose.TabIndex = 14
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.cmdClose)
        Me.Frame1.Controls.Add(Me.fraHeatFlux)
        Me.Frame1.Controls.Add(Me.fraGlassProps)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(8, 8)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(369, 473)
        Me.Frame1.TabIndex = 0
        Me.Frame1.TabStop = False
        '
        'fraHeatFlux
        '
        Me.fraHeatFlux.BackColor = System.Drawing.SystemColors.Control
        Me.fraHeatFlux.Controls.Add(Me.txtGlassDistance)
        Me.fraHeatFlux.Controls.Add(Me.optGlassWithFlame)
        Me.fraHeatFlux.Controls.Add(Me.optGlassNoFlameFlux)
        Me.fraHeatFlux.Controls.Add(Me.Label17)
        Me.fraHeatFlux.Controls.Add(Me.Label11)
        Me.fraHeatFlux.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraHeatFlux.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraHeatFlux.Location = New System.Drawing.Point(16, 247)
        Me.fraHeatFlux.Name = "fraHeatFlux"
        Me.fraHeatFlux.Padding = New System.Windows.Forms.Padding(0)
        Me.fraHeatFlux.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraHeatFlux.Size = New System.Drawing.Size(241, 121)
        Me.fraHeatFlux.TabIndex = 33
        Me.fraHeatFlux.TabStop = False
        Me.fraHeatFlux.Text = "Heat Flux Options"
        '
        'txtGlassDistance
        '
        Me.txtGlassDistance.AcceptsReturn = True
        Me.txtGlassDistance.BackColor = System.Drawing.SystemColors.Window
        Me.txtGlassDistance.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtGlassDistance.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGlassDistance.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtGlassDistance.Location = New System.Drawing.Point(144, 86)
        Me.txtGlassDistance.MaxLength = 0
        Me.txtGlassDistance.Name = "txtGlassDistance"
        Me.txtGlassDistance.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtGlassDistance.Size = New System.Drawing.Size(57, 20)
        Me.txtGlassDistance.TabIndex = 13
        Me.txtGlassDistance.Text = "0"
        '
        'optGlassWithFlame
        '
        Me.optGlassWithFlame.BackColor = System.Drawing.SystemColors.Control
        Me.optGlassWithFlame.Cursor = System.Windows.Forms.Cursors.Default
        Me.optGlassWithFlame.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optGlassWithFlame.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optGlassWithFlame.Location = New System.Drawing.Point(16, 56)
        Me.optGlassWithFlame.Name = "optGlassWithFlame"
        Me.optGlassWithFlame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optGlassWithFlame.Size = New System.Drawing.Size(209, 17)
        Me.optGlassWithFlame.TabIndex = 12
        Me.optGlassWithFlame.TabStop = True
        Me.optGlassWithFlame.Text = "Glass heated by flame and hot layer"
        Me.optGlassWithFlame.UseVisualStyleBackColor = False
        '
        'optGlassNoFlameFlux
        '
        Me.optGlassNoFlameFlux.BackColor = System.Drawing.SystemColors.Control
        Me.optGlassNoFlameFlux.Checked = True
        Me.optGlassNoFlameFlux.Cursor = System.Windows.Forms.Cursors.Default
        Me.optGlassNoFlameFlux.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optGlassNoFlameFlux.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optGlassNoFlameFlux.Location = New System.Drawing.Point(16, 24)
        Me.optGlassNoFlameFlux.Name = "optGlassNoFlameFlux"
        Me.optGlassNoFlameFlux.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optGlassNoFlameFlux.Size = New System.Drawing.Size(193, 17)
        Me.optGlassNoFlameFlux.TabIndex = 11
        Me.optGlassNoFlameFlux.TabStop = True
        Me.optGlassNoFlameFlux.Text = "Glass heated by hot layer only"
        Me.optGlassNoFlameFlux.UseVisualStyleBackColor = False
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.SystemColors.Control
        Me.Label17.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label17.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label17.Location = New System.Drawing.Point(208, 88)
        Me.Label17.Name = "Label17"
        Me.Label17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label17.Size = New System.Drawing.Size(25, 17)
        Me.Label17.TabIndex = 35
        Me.Label17.Text = "m"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Location = New System.Drawing.Point(16, 88)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(129, 25)
        Me.Label11.TabIndex = 34
        Me.Label11.Text = "Glass to flame distance"
        '
        'fraGlassProps
        '
        Me.fraGlassProps.BackColor = System.Drawing.SystemColors.Control
        Me.fraGlassProps.Controls.Add(Me.txtGlassFalloutTime)
        Me.fraGlassProps.Controls.Add(Me.txtGLASSexpansion)
        Me.fraGlassProps.Controls.Add(Me.txtGLASSshading)
        Me.fraGlassProps.Controls.Add(Me.txtGLASSbreakingstress)
        Me.fraGlassProps.Controls.Add(Me.txtGLASSYoungsModulus)
        Me.fraGlassProps.Controls.Add(Me.txtGLASSthickness)
        Me.fraGlassProps.Controls.Add(Me.txtGLASSalpha)
        Me.fraGlassProps.Controls.Add(Me.txtGLASSconductivity)
        Me.fraGlassProps.Controls.Add(Me.Label3)
        Me.fraGlassProps.Controls.Add(Me.lblFalloutTime)
        Me.fraGlassProps.Controls.Add(Me.Label16)
        Me.fraGlassProps.Controls.Add(Me.Label15)
        Me.fraGlassProps.Controls.Add(Me.Label14)
        Me.fraGlassProps.Controls.Add(Me.Label13)
        Me.fraGlassProps.Controls.Add(Me.Label12)
        Me.fraGlassProps.Controls.Add(Me.Label10)
        Me.fraGlassProps.Controls.Add(Me.Label9)
        Me.fraGlassProps.Controls.Add(Me.Label8)
        Me.fraGlassProps.Controls.Add(Me.Label7)
        Me.fraGlassProps.Controls.Add(Me.Label6)
        Me.fraGlassProps.Controls.Add(Me.Label5)
        Me.fraGlassProps.Controls.Add(Me.Label4)
        Me.fraGlassProps.Controls.Add(Me.Label2)
        Me.fraGlassProps.Controls.Add(Me.Label1)
        Me.fraGlassProps.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraGlassProps.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraGlassProps.Location = New System.Drawing.Point(16, 16)
        Me.fraGlassProps.Name = "fraGlassProps"
        Me.fraGlassProps.Padding = New System.Windows.Forms.Padding(0)
        Me.fraGlassProps.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraGlassProps.Size = New System.Drawing.Size(337, 225)
        Me.fraGlassProps.TabIndex = 16
        Me.fraGlassProps.TabStop = False
        Me.fraGlassProps.Text = "Glass Properties for Current Vent"
        '
        'txtGlassFalloutTime
        '
        Me.txtGlassFalloutTime.AcceptsReturn = True
        Me.txtGlassFalloutTime.BackColor = System.Drawing.SystemColors.Window
        Me.txtGlassFalloutTime.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtGlassFalloutTime.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGlassFalloutTime.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtGlassFalloutTime.Location = New System.Drawing.Point(184, 192)
        Me.txtGlassFalloutTime.MaxLength = 0
        Me.txtGlassFalloutTime.Name = "txtGlassFalloutTime"
        Me.txtGlassFalloutTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtGlassFalloutTime.Size = New System.Drawing.Size(81, 20)
        Me.txtGlassFalloutTime.TabIndex = 10
        Me.txtGlassFalloutTime.Text = "0"
        '
        'txtGLASSexpansion
        '
        Me.txtGLASSexpansion.AcceptsReturn = True
        Me.txtGLASSexpansion.BackColor = System.Drawing.SystemColors.Window
        Me.txtGLASSexpansion.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtGLASSexpansion.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGLASSexpansion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtGLASSexpansion.Location = New System.Drawing.Point(184, 168)
        Me.txtGLASSexpansion.MaxLength = 0
        Me.txtGLASSexpansion.Name = "txtGLASSexpansion"
        Me.txtGLASSexpansion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtGLASSexpansion.Size = New System.Drawing.Size(81, 20)
        Me.txtGLASSexpansion.TabIndex = 9
        Me.txtGLASSexpansion.Text = "0.0000095"
        '
        'txtGLASSshading
        '
        Me.txtGLASSshading.AcceptsReturn = True
        Me.txtGLASSshading.BackColor = System.Drawing.SystemColors.Window
        Me.txtGLASSshading.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtGLASSshading.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGLASSshading.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtGLASSshading.Location = New System.Drawing.Point(184, 144)
        Me.txtGLASSshading.MaxLength = 0
        Me.txtGLASSshading.Name = "txtGLASSshading"
        Me.txtGLASSshading.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtGLASSshading.Size = New System.Drawing.Size(81, 20)
        Me.txtGLASSshading.TabIndex = 8
        Me.txtGLASSshading.Text = "15"
        '
        'txtGLASSbreakingstress
        '
        Me.txtGLASSbreakingstress.AcceptsReturn = True
        Me.txtGLASSbreakingstress.BackColor = System.Drawing.SystemColors.Window
        Me.txtGLASSbreakingstress.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtGLASSbreakingstress.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGLASSbreakingstress.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtGLASSbreakingstress.Location = New System.Drawing.Point(184, 120)
        Me.txtGLASSbreakingstress.MaxLength = 0
        Me.txtGLASSbreakingstress.Name = "txtGLASSbreakingstress"
        Me.txtGLASSbreakingstress.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtGLASSbreakingstress.Size = New System.Drawing.Size(81, 20)
        Me.txtGLASSbreakingstress.TabIndex = 7
        Me.txtGLASSbreakingstress.Text = "47"
        '
        'txtGLASSYoungsModulus
        '
        Me.txtGLASSYoungsModulus.AcceptsReturn = True
        Me.txtGLASSYoungsModulus.BackColor = System.Drawing.SystemColors.Window
        Me.txtGLASSYoungsModulus.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtGLASSYoungsModulus.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGLASSYoungsModulus.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtGLASSYoungsModulus.Location = New System.Drawing.Point(184, 96)
        Me.txtGLASSYoungsModulus.MaxLength = 0
        Me.txtGLASSYoungsModulus.Name = "txtGLASSYoungsModulus"
        Me.txtGLASSYoungsModulus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtGLASSYoungsModulus.Size = New System.Drawing.Size(81, 20)
        Me.txtGLASSYoungsModulus.TabIndex = 6
        Me.txtGLASSYoungsModulus.Text = "72000"
        '
        'txtGLASSthickness
        '
        Me.txtGLASSthickness.AcceptsReturn = True
        Me.txtGLASSthickness.BackColor = System.Drawing.SystemColors.Window
        Me.txtGLASSthickness.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtGLASSthickness.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGLASSthickness.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtGLASSthickness.Location = New System.Drawing.Point(184, 24)
        Me.txtGLASSthickness.MaxLength = 0
        Me.txtGLASSthickness.Name = "txtGLASSthickness"
        Me.txtGLASSthickness.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtGLASSthickness.Size = New System.Drawing.Size(81, 20)
        Me.txtGLASSthickness.TabIndex = 3
        Me.txtGLASSthickness.Text = "4"
        '
        'txtGLASSalpha
        '
        Me.txtGLASSalpha.AcceptsReturn = True
        Me.txtGLASSalpha.BackColor = System.Drawing.SystemColors.Window
        Me.txtGLASSalpha.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtGLASSalpha.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGLASSalpha.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtGLASSalpha.Location = New System.Drawing.Point(184, 72)
        Me.txtGLASSalpha.MaxLength = 0
        Me.txtGLASSalpha.Name = "txtGLASSalpha"
        Me.txtGLASSalpha.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtGLASSalpha.Size = New System.Drawing.Size(81, 20)
        Me.txtGLASSalpha.TabIndex = 5
        Me.txtGLASSalpha.Text = "0.00000036"
        '
        'txtGLASSconductivity
        '
        Me.txtGLASSconductivity.AcceptsReturn = True
        Me.txtGLASSconductivity.BackColor = System.Drawing.SystemColors.Window
        Me.txtGLASSconductivity.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtGLASSconductivity.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGLASSconductivity.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtGLASSconductivity.Location = New System.Drawing.Point(184, 48)
        Me.txtGLASSconductivity.MaxLength = 0
        Me.txtGLASSconductivity.Name = "txtGLASSconductivity"
        Me.txtGLASSconductivity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtGLASSconductivity.Size = New System.Drawing.Size(81, 20)
        Me.txtGLASSconductivity.TabIndex = 4
        Me.txtGLASSconductivity.Text = "0.76"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(16, 192)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(153, 17)
        Me.Label3.TabIndex = 32
        Me.Label3.Text = "Time from fracture to fallout"
        '
        'lblFalloutTime
        '
        Me.lblFalloutTime.BackColor = System.Drawing.SystemColors.Control
        Me.lblFalloutTime.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFalloutTime.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFalloutTime.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFalloutTime.Location = New System.Drawing.Point(272, 192)
        Me.lblFalloutTime.Name = "lblFalloutTime"
        Me.lblFalloutTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFalloutTime.Size = New System.Drawing.Size(33, 17)
        Me.lblFalloutTime.TabIndex = 31
        Me.lblFalloutTime.Text = "sec"
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.SystemColors.Control
        Me.Label16.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label16.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label16.Location = New System.Drawing.Point(272, 168)
        Me.Label16.Name = "Label16"
        Me.Label16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label16.Size = New System.Drawing.Size(41, 17)
        Me.Label16.TabIndex = 30
        Me.Label16.Text = "/°C"
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.SystemColors.Control
        Me.Label15.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label15.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label15.Location = New System.Drawing.Point(272, 144)
        Me.Label15.Name = "Label15"
        Me.Label15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label15.Size = New System.Drawing.Size(41, 17)
        Me.Label15.TabIndex = 29
        Me.Label15.Text = "mm"
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.SystemColors.Control
        Me.Label14.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label14.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label14.Location = New System.Drawing.Point(272, 120)
        Me.Label14.Name = "Label14"
        Me.Label14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label14.Size = New System.Drawing.Size(41, 17)
        Me.Label14.TabIndex = 28
        Me.Label14.Text = "MPa"
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.SystemColors.Control
        Me.Label13.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label13.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label13.Location = New System.Drawing.Point(272, 96)
        Me.Label13.Name = "Label13"
        Me.Label13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label13.Size = New System.Drawing.Size(41, 17)
        Me.Label13.TabIndex = 27
        Me.Label13.Text = "MPa"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.SystemColors.Control
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label12.Location = New System.Drawing.Point(272, 24)
        Me.Label12.Name = "Label12"
        Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label12.Size = New System.Drawing.Size(33, 17)
        Me.Label12.TabIndex = 26
        Me.Label12.Text = "mm"
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(272, 72)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(41, 17)
        Me.Label10.TabIndex = 25
        Me.Label10.Text = "m²/s"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(272, 48)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(41, 17)
        Me.Label9.TabIndex = 24
        Me.Label9.Text = "W/mK"
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(16, 144)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(177, 17)
        Me.Label8.TabIndex = 23
        Me.Label8.Text = "Shading depth"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(16, 24)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(169, 17)
        Me.Label7.TabIndex = 22
        Me.Label7.Text = "Thickness"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(16, 120)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(177, 17)
        Me.Label6.TabIndex = 21
        Me.Label6.Text = "Fracture stress"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(16, 96)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(177, 17)
        Me.Label5.TabIndex = 20
        Me.Label5.Text = "Young's modulus"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(16, 168)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(177, 17)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Thermal expansion coefficient"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(16, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(177, 17)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Thermal diffusivity"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(16, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(169, 17)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "Thermal conductivity"
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'frmGlassBreak
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(377, 394)
        Me.ControlBox = False
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 23)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmGlassBreak"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Vent - Glass Properties"
        Me.TopMost = True
        Me.Frame1.ResumeLayout(False)
        Me.fraHeatFlux.ResumeLayout(False)
        Me.fraHeatFlux.PerformLayout()
        Me.fraGlassProps.ResumeLayout(False)
        Me.fraGlassProps.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
#End Region 
End Class