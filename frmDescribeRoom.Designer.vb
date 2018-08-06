<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmDescribeRoom
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
    Public WithEvents _lstRoomID_1 As System.Windows.Forms.ComboBox
	Public WithEvents _cmdOK_3 As System.Windows.Forms.Button
    Public WithEvents txtWallThickness As System.Windows.Forms.TextBox
	Public WithEvents cmdPickSurface As System.Windows.Forms.Button
	Public WithEvents chkAddWallSubstrate As System.Windows.Forms.CheckBox
	Public WithEvents lblWallThickness As System.Windows.Forms.Label
	Public WithEvents lblWallSurface As System.Windows.Forms.Label
	Public WithEvents fraWallMaterial As System.Windows.Forms.GroupBox
	Public WithEvents txtWallSubThickness As System.Windows.Forms.TextBox
	Public WithEvents cmdPickSubstrate As System.Windows.Forms.Button
	Public WithEvents lblWallSubThickness As System.Windows.Forms.Label
	Public WithEvents lblWallSubstrate As System.Windows.Forms.Label
	Public WithEvents fraWallSubstrate As System.Windows.Forms.GroupBox
	Public WithEvents _Label9_0 As System.Windows.Forms.Label
	Public WithEvents __SSTab1_0_TabPage3 As System.Windows.Forms.TabPage
	Public WithEvents _lstRoomID_2 As System.Windows.Forms.ComboBox
	Public WithEvents _cmdOK_4 As System.Windows.Forms.Button
    Public WithEvents cmdPickCsurface As System.Windows.Forms.Button
	Public WithEvents txtCeilingThickness As System.Windows.Forms.TextBox
	Public WithEvents chkAddSubstrate As System.Windows.Forms.CheckBox
	Public WithEvents lblCeilingSurface As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Frame6 As System.Windows.Forms.GroupBox
	Public WithEvents cmdPickCsubstrate As System.Windows.Forms.Button
	Public WithEvents txtCeilingSubThickness As System.Windows.Forms.TextBox
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents lblCeilingSubstrate As System.Windows.Forms.Label
	Public WithEvents fraCeilingSubstrate As System.Windows.Forms.GroupBox
	Public WithEvents _Label9_1 As System.Windows.Forms.Label
	Public WithEvents __SSTab1_0_TabPage4 As System.Windows.Forms.TabPage
	Public WithEvents txtFloorSubThickness As System.Windows.Forms.TextBox
	Public WithEvents cmdPickFsubstrate As System.Windows.Forms.Button
	Public WithEvents lblFloorSubstrate As System.Windows.Forms.Label
	Public WithEvents Label12 As System.Windows.Forms.Label
	Public WithEvents fraFloorSubstrate As System.Windows.Forms.GroupBox
	Public WithEvents _lstRoomID_3 As System.Windows.Forms.ComboBox
	Public WithEvents _cmdOK_5 As System.Windows.Forms.Button
    Public WithEvents chkAddFloorSubstrate As System.Windows.Forms.CheckBox
	Public WithEvents txtFloorThickness As System.Windows.Forms.TextBox
	Public WithEvents cmdPickFloor As System.Windows.Forms.Button
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents lblFloorSurface As System.Windows.Forms.Label
	Public WithEvents Frame7 As System.Windows.Forms.GroupBox
	Public WithEvents _Label9_2 As System.Windows.Forms.Label
	Public WithEvents __SSTab1_0_TabPage5 As System.Windows.Forms.TabPage
	Public WithEvents _SSTab1_0 As System.Windows.Forms.TabControl
	Public WithEvents Label9 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents SSTab1 As Microsoft.VisualBasic.Compatibility.VB6.TabControlArray
    Public WithEvents cmdOK As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	Public WithEvents lstRoomID As Microsoft.VisualBasic.Compatibility.VB6.ComboBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdPickSurface = New System.Windows.Forms.Button
        Me.cmdPickSubstrate = New System.Windows.Forms.Button
        Me.cmdPickCsurface = New System.Windows.Forms.Button
        Me.cmdPickCsubstrate = New System.Windows.Forms.Button
        Me.cmdPickFsubstrate = New System.Windows.Forms.Button
        Me.cmdPickFloor = New System.Windows.Forms.Button
        Me._SSTab1_0 = New System.Windows.Forms.TabControl
        Me.__SSTab1_0_TabPage3 = New System.Windows.Forms.TabPage
        Me._lstRoomID_1 = New System.Windows.Forms.ComboBox
        Me._cmdOK_3 = New System.Windows.Forms.Button
        Me.fraWallMaterial = New System.Windows.Forms.GroupBox
        Me.txtWallThickness = New System.Windows.Forms.TextBox
        Me.chkAddWallSubstrate = New System.Windows.Forms.CheckBox
        Me.lblWallThickness = New System.Windows.Forms.Label
        Me.lblWallSurface = New System.Windows.Forms.Label
        Me.fraWallSubstrate = New System.Windows.Forms.GroupBox
        Me.txtWallSubThickness = New System.Windows.Forms.TextBox
        Me.lblWallSubThickness = New System.Windows.Forms.Label
        Me.lblWallSubstrate = New System.Windows.Forms.Label
        Me._Label9_0 = New System.Windows.Forms.Label
        Me.__SSTab1_0_TabPage4 = New System.Windows.Forms.TabPage
        Me._lstRoomID_2 = New System.Windows.Forms.ComboBox
        Me._cmdOK_4 = New System.Windows.Forms.Button
        Me.Frame6 = New System.Windows.Forms.GroupBox
        Me.txtCeilingThickness = New System.Windows.Forms.TextBox
        Me.chkAddSubstrate = New System.Windows.Forms.CheckBox
        Me.lblCeilingSurface = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.fraCeilingSubstrate = New System.Windows.Forms.GroupBox
        Me.txtCeilingSubThickness = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.lblCeilingSubstrate = New System.Windows.Forms.Label
        Me._Label9_1 = New System.Windows.Forms.Label
        Me.__SSTab1_0_TabPage5 = New System.Windows.Forms.TabPage
        Me.fraFloorSubstrate = New System.Windows.Forms.GroupBox
        Me.txtFloorSubThickness = New System.Windows.Forms.TextBox
        Me.lblFloorSubstrate = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me._lstRoomID_3 = New System.Windows.Forms.ComboBox
        Me._cmdOK_5 = New System.Windows.Forms.Button
        Me.Frame7 = New System.Windows.Forms.GroupBox
        Me.chkAddFloorSubstrate = New System.Windows.Forms.CheckBox
        Me.txtFloorThickness = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblFloorSurface = New System.Windows.Forms.Label
        Me._Label9_2 = New System.Windows.Forms.Label
        Me.Label9 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.SSTab1 = New Microsoft.VisualBasic.Compatibility.VB6.TabControlArray(Me.components)
        Me.cmdOK = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(Me.components)
        Me.lstRoomID = New Microsoft.VisualBasic.Compatibility.VB6.ComboBoxArray(Me.components)
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me._SSTab1_0.SuspendLayout()
        Me.__SSTab1_0_TabPage3.SuspendLayout()
        Me.fraWallMaterial.SuspendLayout()
        Me.fraWallSubstrate.SuspendLayout()
        Me.__SSTab1_0_TabPage4.SuspendLayout()
        Me.Frame6.SuspendLayout()
        Me.fraCeilingSubstrate.SuspendLayout()
        Me.__SSTab1_0_TabPage5.SuspendLayout()
        Me.fraFloorSubstrate.SuspendLayout()
        Me.Frame7.SuspendLayout()
        CType(Me.Label9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SSTab1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdOK, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lstRoomID, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdPickSurface
        '
        Me.cmdPickSurface.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPickSurface.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPickSurface.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPickSurface.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPickSurface.Location = New System.Drawing.Point(16, 64)
        Me.cmdPickSurface.Name = "cmdPickSurface"
        Me.cmdPickSurface.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPickSurface.Size = New System.Drawing.Size(73, 25)
        Me.cmdPickSurface.TabIndex = 19
        Me.cmdPickSurface.Text = "Pick"
        Me.ToolTip1.SetToolTip(Me.cmdPickSurface, "select material from thermal database")
        Me.cmdPickSurface.UseVisualStyleBackColor = False
        '
        'cmdPickSubstrate
        '
        Me.cmdPickSubstrate.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPickSubstrate.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPickSubstrate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPickSubstrate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPickSubstrate.Location = New System.Drawing.Point(16, 64)
        Me.cmdPickSubstrate.Name = "cmdPickSubstrate"
        Me.cmdPickSubstrate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPickSubstrate.Size = New System.Drawing.Size(73, 25)
        Me.cmdPickSubstrate.TabIndex = 23
        Me.cmdPickSubstrate.Text = "Pick"
        Me.ToolTip1.SetToolTip(Me.cmdPickSubstrate, "select material from thermal database")
        Me.cmdPickSubstrate.UseVisualStyleBackColor = False
        '
        'cmdPickCsurface
        '
        Me.cmdPickCsurface.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPickCsurface.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPickCsurface.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPickCsurface.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPickCsurface.Location = New System.Drawing.Point(16, 64)
        Me.cmdPickCsurface.Name = "cmdPickCsurface"
        Me.cmdPickCsurface.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPickCsurface.Size = New System.Drawing.Size(73, 25)
        Me.cmdPickCsurface.TabIndex = 9
        Me.cmdPickCsurface.Text = "Pick"
        Me.ToolTip1.SetToolTip(Me.cmdPickCsurface, "select material from thermal database")
        Me.cmdPickCsurface.UseVisualStyleBackColor = False
        '
        'cmdPickCsubstrate
        '
        Me.cmdPickCsubstrate.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPickCsubstrate.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPickCsubstrate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPickCsubstrate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPickCsubstrate.Location = New System.Drawing.Point(16, 64)
        Me.cmdPickCsubstrate.Name = "cmdPickCsubstrate"
        Me.cmdPickCsubstrate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPickCsubstrate.Size = New System.Drawing.Size(73, 25)
        Me.cmdPickCsubstrate.TabIndex = 13
        Me.cmdPickCsubstrate.Text = "Pick"
        Me.ToolTip1.SetToolTip(Me.cmdPickCsubstrate, "select material from thermal database")
        Me.cmdPickCsubstrate.UseVisualStyleBackColor = False
        '
        'cmdPickFsubstrate
        '
        Me.cmdPickFsubstrate.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPickFsubstrate.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPickFsubstrate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPickFsubstrate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPickFsubstrate.Location = New System.Drawing.Point(16, 64)
        Me.cmdPickFsubstrate.Name = "cmdPickFsubstrate"
        Me.cmdPickFsubstrate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPickFsubstrate.Size = New System.Drawing.Size(73, 25)
        Me.cmdPickFsubstrate.TabIndex = 106
        Me.cmdPickFsubstrate.Text = "Pick"
        Me.ToolTip1.SetToolTip(Me.cmdPickFsubstrate, "select material from thermal database")
        Me.cmdPickFsubstrate.UseVisualStyleBackColor = False
        '
        'cmdPickFloor
        '
        Me.cmdPickFloor.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPickFloor.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPickFloor.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPickFloor.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPickFloor.Location = New System.Drawing.Point(16, 64)
        Me.cmdPickFloor.Name = "cmdPickFloor"
        Me.cmdPickFloor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPickFloor.Size = New System.Drawing.Size(73, 25)
        Me.cmdPickFloor.TabIndex = 3
        Me.cmdPickFloor.Text = "Pick"
        Me.ToolTip1.SetToolTip(Me.cmdPickFloor, "select material from thermal database")
        Me.cmdPickFloor.UseVisualStyleBackColor = False
        '
        '_SSTab1_0
        '
        Me._SSTab1_0.Controls.Add(Me.__SSTab1_0_TabPage3)
        Me._SSTab1_0.Controls.Add(Me.__SSTab1_0_TabPage4)
        Me._SSTab1_0.Controls.Add(Me.__SSTab1_0_TabPage5)
        Me._SSTab1_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SSTab1.SetIndex(Me._SSTab1_0, CType(0, Short))
        Me._SSTab1_0.ItemSize = New System.Drawing.Size(42, 15)
        Me._SSTab1_0.Location = New System.Drawing.Point(0, 8)
        Me._SSTab1_0.Name = "_SSTab1_0"
        Me._SSTab1_0.SelectedIndex = 0
        Me._SSTab1_0.Size = New System.Drawing.Size(517, 316)
        Me._SSTab1_0.TabIndex = 0
        '
        '__SSTab1_0_TabPage3
        '
        Me.__SSTab1_0_TabPage3.BackColor = System.Drawing.SystemColors.Control
        Me.__SSTab1_0_TabPage3.Controls.Add(Me._lstRoomID_1)
        Me.__SSTab1_0_TabPage3.Controls.Add(Me._cmdOK_3)
        Me.__SSTab1_0_TabPage3.Controls.Add(Me.fraWallMaterial)
        Me.__SSTab1_0_TabPage3.Controls.Add(Me.fraWallSubstrate)
        Me.__SSTab1_0_TabPage3.Controls.Add(Me._Label9_0)
        Me.__SSTab1_0_TabPage3.Location = New System.Drawing.Point(4, 19)
        Me.__SSTab1_0_TabPage3.Name = "__SSTab1_0_TabPage3"
        Me.__SSTab1_0_TabPage3.Size = New System.Drawing.Size(509, 293)
        Me.__SSTab1_0_TabPage3.TabIndex = 3
        Me.__SSTab1_0_TabPage3.Text = "Wall Material"
        '
        '_lstRoomID_1
        '
        Me._lstRoomID_1.BackColor = System.Drawing.SystemColors.Window
        Me._lstRoomID_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lstRoomID_1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me._lstRoomID_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lstRoomID_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstRoomID.SetIndex(Me._lstRoomID_1, CType(1, Short))
        Me._lstRoomID_1.Location = New System.Drawing.Point(128, 256)
        Me._lstRoomID_1.Name = "_lstRoomID_1"
        Me._lstRoomID_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lstRoomID_1.Size = New System.Drawing.Size(89, 22)
        Me._lstRoomID_1.TabIndex = 24
        '
        '_cmdOK_3
        '
        Me._cmdOK_3.BackColor = System.Drawing.SystemColors.Control
        Me._cmdOK_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._cmdOK_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cmdOK_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.SetIndex(Me._cmdOK_3, CType(3, Short))
        Me._cmdOK_3.Location = New System.Drawing.Point(415, 256)
        Me._cmdOK_3.Name = "_cmdOK_3"
        Me._cmdOK_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cmdOK_3.Size = New System.Drawing.Size(73, 25)
        Me._cmdOK_3.TabIndex = 25
        Me._cmdOK_3.Text = "OK"
        Me._cmdOK_3.UseVisualStyleBackColor = False
        '
        'fraWallMaterial
        '
        Me.fraWallMaterial.BackColor = System.Drawing.SystemColors.Control
        Me.fraWallMaterial.Controls.Add(Me.txtWallThickness)
        Me.fraWallMaterial.Controls.Add(Me.cmdPickSurface)
        Me.fraWallMaterial.Controls.Add(Me.chkAddWallSubstrate)
        Me.fraWallMaterial.Controls.Add(Me.lblWallThickness)
        Me.fraWallMaterial.Controls.Add(Me.lblWallSurface)
        Me.fraWallMaterial.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraWallMaterial.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraWallMaterial.Location = New System.Drawing.Point(48, 19)
        Me.fraWallMaterial.Name = "fraWallMaterial"
        Me.fraWallMaterial.Padding = New System.Windows.Forms.Padding(0)
        Me.fraWallMaterial.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraWallMaterial.Size = New System.Drawing.Size(440, 105)
        Me.fraWallMaterial.TabIndex = 73
        Me.fraWallMaterial.TabStop = False
        Me.fraWallMaterial.Text = "Surface Material"
        '
        'txtWallThickness
        '
        Me.txtWallThickness.AcceptsReturn = True
        Me.txtWallThickness.BackColor = System.Drawing.Color.White
        Me.txtWallThickness.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWallThickness.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWallThickness.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWallThickness.Location = New System.Drawing.Point(369, 24)
        Me.txtWallThickness.MaxLength = 0
        Me.txtWallThickness.Name = "txtWallThickness"
        Me.txtWallThickness.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWallThickness.Size = New System.Drawing.Size(49, 20)
        Me.txtWallThickness.TabIndex = 18
        Me.txtWallThickness.Text = "13.0"
        '
        'chkAddWallSubstrate
        '
        Me.chkAddWallSubstrate.BackColor = System.Drawing.SystemColors.Control
        Me.chkAddWallSubstrate.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkAddWallSubstrate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAddWallSubstrate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkAddWallSubstrate.Location = New System.Drawing.Point(289, 72)
        Me.chkAddWallSubstrate.Name = "chkAddWallSubstrate"
        Me.chkAddWallSubstrate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkAddWallSubstrate.Size = New System.Drawing.Size(113, 17)
        Me.chkAddWallSubstrate.TabIndex = 20
        Me.chkAddWallSubstrate.Text = "Include Substrate"
        Me.chkAddWallSubstrate.UseVisualStyleBackColor = False
        '
        'lblWallThickness
        '
        Me.lblWallThickness.BackColor = System.Drawing.Color.Transparent
        Me.lblWallThickness.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblWallThickness.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWallThickness.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblWallThickness.Location = New System.Drawing.Point(249, 24)
        Me.lblWallThickness.Name = "lblWallThickness"
        Me.lblWallThickness.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblWallThickness.Size = New System.Drawing.Size(113, 25)
        Me.lblWallThickness.TabIndex = 74
        Me.lblWallThickness.Text = "Thickness (mm)"
        Me.lblWallThickness.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblWallSurface
        '
        Me.lblWallSurface.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblWallSurface.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblWallSurface.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblWallSurface.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWallSurface.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWallSurface.Location = New System.Drawing.Point(16, 24)
        Me.lblWallSurface.Name = "lblWallSurface"
        Me.lblWallSurface.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblWallSurface.Size = New System.Drawing.Size(238, 20)
        Me.lblWallSurface.TabIndex = 17
        Me.lblWallSurface.Text = "none"
        '
        'fraWallSubstrate
        '
        Me.fraWallSubstrate.BackColor = System.Drawing.SystemColors.Control
        Me.fraWallSubstrate.Controls.Add(Me.txtWallSubThickness)
        Me.fraWallSubstrate.Controls.Add(Me.cmdPickSubstrate)
        Me.fraWallSubstrate.Controls.Add(Me.lblWallSubThickness)
        Me.fraWallSubstrate.Controls.Add(Me.lblWallSubstrate)
        Me.fraWallSubstrate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraWallSubstrate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraWallSubstrate.Location = New System.Drawing.Point(48, 130)
        Me.fraWallSubstrate.Name = "fraWallSubstrate"
        Me.fraWallSubstrate.Padding = New System.Windows.Forms.Padding(0)
        Me.fraWallSubstrate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraWallSubstrate.Size = New System.Drawing.Size(440, 105)
        Me.fraWallSubstrate.TabIndex = 71
        Me.fraWallSubstrate.TabStop = False
        Me.fraWallSubstrate.Text = "Substrate Material"
        Me.fraWallSubstrate.Visible = False
        '
        'txtWallSubThickness
        '
        Me.txtWallSubThickness.AcceptsReturn = True
        Me.txtWallSubThickness.BackColor = System.Drawing.Color.White
        Me.txtWallSubThickness.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWallSubThickness.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWallSubThickness.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWallSubThickness.Location = New System.Drawing.Point(369, 24)
        Me.txtWallSubThickness.MaxLength = 0
        Me.txtWallSubThickness.Name = "txtWallSubThickness"
        Me.txtWallSubThickness.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWallSubThickness.Size = New System.Drawing.Size(49, 20)
        Me.txtWallSubThickness.TabIndex = 22
        Me.txtWallSubThickness.Text = "150.0"
        '
        'lblWallSubThickness
        '
        Me.lblWallSubThickness.BackColor = System.Drawing.Color.Transparent
        Me.lblWallSubThickness.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblWallSubThickness.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWallSubThickness.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblWallSubThickness.Location = New System.Drawing.Point(249, 24)
        Me.lblWallSubThickness.Name = "lblWallSubThickness"
        Me.lblWallSubThickness.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblWallSubThickness.Size = New System.Drawing.Size(113, 25)
        Me.lblWallSubThickness.TabIndex = 72
        Me.lblWallSubThickness.Text = "Thickness (mm)"
        Me.lblWallSubThickness.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblWallSubstrate
        '
        Me.lblWallSubstrate.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblWallSubstrate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblWallSubstrate.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblWallSubstrate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWallSubstrate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWallSubstrate.Location = New System.Drawing.Point(16, 24)
        Me.lblWallSubstrate.Name = "lblWallSubstrate"
        Me.lblWallSubstrate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblWallSubstrate.Size = New System.Drawing.Size(238, 20)
        Me.lblWallSubstrate.TabIndex = 21
        Me.lblWallSubstrate.Text = "none"
        '
        '_Label9_0
        '
        Me._Label9_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label9_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label9_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label9_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.SetIndex(Me._Label9_0, CType(0, Short))
        Me._Label9_0.Location = New System.Drawing.Point(24, 256)
        Me._Label9_0.Name = "_Label9_0"
        Me._Label9_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label9_0.Size = New System.Drawing.Size(89, 17)
        Me._Label9_0.TabIndex = 81
        Me._Label9_0.Text = "Current Room"
        Me._Label9_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '__SSTab1_0_TabPage4
        '
        Me.__SSTab1_0_TabPage4.BackColor = System.Drawing.SystemColors.Control
        Me.__SSTab1_0_TabPage4.Controls.Add(Me._lstRoomID_2)
        Me.__SSTab1_0_TabPage4.Controls.Add(Me._cmdOK_4)
        Me.__SSTab1_0_TabPage4.Controls.Add(Me.Frame6)
        Me.__SSTab1_0_TabPage4.Controls.Add(Me.fraCeilingSubstrate)
        Me.__SSTab1_0_TabPage4.Controls.Add(Me._Label9_1)
        Me.__SSTab1_0_TabPage4.Location = New System.Drawing.Point(4, 19)
        Me.__SSTab1_0_TabPage4.Name = "__SSTab1_0_TabPage4"
        Me.__SSTab1_0_TabPage4.Size = New System.Drawing.Size(509, 293)
        Me.__SSTab1_0_TabPage4.TabIndex = 4
        Me.__SSTab1_0_TabPage4.Text = "Ceiling Material"
        '
        '_lstRoomID_2
        '
        Me._lstRoomID_2.BackColor = System.Drawing.SystemColors.Window
        Me._lstRoomID_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lstRoomID_2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me._lstRoomID_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lstRoomID_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstRoomID.SetIndex(Me._lstRoomID_2, CType(2, Short))
        Me._lstRoomID_2.Location = New System.Drawing.Point(128, 256)
        Me._lstRoomID_2.Name = "_lstRoomID_2"
        Me._lstRoomID_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lstRoomID_2.Size = New System.Drawing.Size(89, 22)
        Me._lstRoomID_2.TabIndex = 14
        '
        '_cmdOK_4
        '
        Me._cmdOK_4.BackColor = System.Drawing.SystemColors.Control
        Me._cmdOK_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._cmdOK_4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cmdOK_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.SetIndex(Me._cmdOK_4, CType(4, Short))
        Me._cmdOK_4.Location = New System.Drawing.Point(415, 256)
        Me._cmdOK_4.Name = "_cmdOK_4"
        Me._cmdOK_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cmdOK_4.Size = New System.Drawing.Size(73, 25)
        Me._cmdOK_4.TabIndex = 15
        Me._cmdOK_4.Text = "OK"
        Me._cmdOK_4.UseVisualStyleBackColor = False
        '
        'Frame6
        '
        Me.Frame6.BackColor = System.Drawing.SystemColors.Control
        Me.Frame6.Controls.Add(Me.cmdPickCsurface)
        Me.Frame6.Controls.Add(Me.txtCeilingThickness)
        Me.Frame6.Controls.Add(Me.chkAddSubstrate)
        Me.Frame6.Controls.Add(Me.lblCeilingSurface)
        Me.Frame6.Controls.Add(Me.Label4)
        Me.Frame6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame6.Location = New System.Drawing.Point(48, 19)
        Me.Frame6.Name = "Frame6"
        Me.Frame6.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame6.Size = New System.Drawing.Size(440, 105)
        Me.Frame6.TabIndex = 77
        Me.Frame6.TabStop = False
        Me.Frame6.Text = "Surface Material"
        '
        'txtCeilingThickness
        '
        Me.txtCeilingThickness.AcceptsReturn = True
        Me.txtCeilingThickness.BackColor = System.Drawing.Color.White
        Me.txtCeilingThickness.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCeilingThickness.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCeilingThickness.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCeilingThickness.Location = New System.Drawing.Point(369, 24)
        Me.txtCeilingThickness.MaxLength = 0
        Me.txtCeilingThickness.Name = "txtCeilingThickness"
        Me.txtCeilingThickness.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCeilingThickness.Size = New System.Drawing.Size(49, 20)
        Me.txtCeilingThickness.TabIndex = 8
        Me.txtCeilingThickness.Text = "13.0"
        '
        'chkAddSubstrate
        '
        Me.chkAddSubstrate.BackColor = System.Drawing.SystemColors.Control
        Me.chkAddSubstrate.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkAddSubstrate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAddSubstrate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkAddSubstrate.Location = New System.Drawing.Point(289, 72)
        Me.chkAddSubstrate.Name = "chkAddSubstrate"
        Me.chkAddSubstrate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkAddSubstrate.Size = New System.Drawing.Size(113, 17)
        Me.chkAddSubstrate.TabIndex = 10
        Me.chkAddSubstrate.Text = "Include Substrate"
        Me.chkAddSubstrate.UseVisualStyleBackColor = False
        '
        'lblCeilingSurface
        '
        Me.lblCeilingSurface.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblCeilingSurface.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCeilingSurface.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCeilingSurface.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCeilingSurface.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCeilingSurface.Location = New System.Drawing.Point(16, 24)
        Me.lblCeilingSurface.Name = "lblCeilingSurface"
        Me.lblCeilingSurface.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCeilingSurface.Size = New System.Drawing.Size(238, 20)
        Me.lblCeilingSurface.TabIndex = 7
        Me.lblCeilingSurface.Text = "none"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label4.Location = New System.Drawing.Point(249, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(113, 25)
        Me.Label4.TabIndex = 78
        Me.Label4.Text = "Thickness (mm)"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'fraCeilingSubstrate
        '
        Me.fraCeilingSubstrate.BackColor = System.Drawing.SystemColors.Control
        Me.fraCeilingSubstrate.Controls.Add(Me.cmdPickCsubstrate)
        Me.fraCeilingSubstrate.Controls.Add(Me.txtCeilingSubThickness)
        Me.fraCeilingSubstrate.Controls.Add(Me.Label2)
        Me.fraCeilingSubstrate.Controls.Add(Me.lblCeilingSubstrate)
        Me.fraCeilingSubstrate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraCeilingSubstrate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraCeilingSubstrate.Location = New System.Drawing.Point(48, 130)
        Me.fraCeilingSubstrate.Name = "fraCeilingSubstrate"
        Me.fraCeilingSubstrate.Padding = New System.Windows.Forms.Padding(0)
        Me.fraCeilingSubstrate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraCeilingSubstrate.Size = New System.Drawing.Size(440, 105)
        Me.fraCeilingSubstrate.TabIndex = 75
        Me.fraCeilingSubstrate.TabStop = False
        Me.fraCeilingSubstrate.Text = "Substrate Material"
        Me.fraCeilingSubstrate.Visible = False
        '
        'txtCeilingSubThickness
        '
        Me.txtCeilingSubThickness.AcceptsReturn = True
        Me.txtCeilingSubThickness.BackColor = System.Drawing.Color.White
        Me.txtCeilingSubThickness.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCeilingSubThickness.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCeilingSubThickness.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCeilingSubThickness.Location = New System.Drawing.Point(369, 24)
        Me.txtCeilingSubThickness.MaxLength = 0
        Me.txtCeilingSubThickness.Name = "txtCeilingSubThickness"
        Me.txtCeilingSubThickness.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCeilingSubThickness.Size = New System.Drawing.Size(49, 20)
        Me.txtCeilingSubThickness.TabIndex = 12
        Me.txtCeilingSubThickness.Text = "150.0"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label2.Location = New System.Drawing.Point(249, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(113, 25)
        Me.Label2.TabIndex = 76
        Me.Label2.Text = "Thickness (mm)"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblCeilingSubstrate
        '
        Me.lblCeilingSubstrate.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblCeilingSubstrate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCeilingSubstrate.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCeilingSubstrate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCeilingSubstrate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCeilingSubstrate.Location = New System.Drawing.Point(16, 24)
        Me.lblCeilingSubstrate.Name = "lblCeilingSubstrate"
        Me.lblCeilingSubstrate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCeilingSubstrate.Size = New System.Drawing.Size(238, 20)
        Me.lblCeilingSubstrate.TabIndex = 11
        Me.lblCeilingSubstrate.Text = "none"
        '
        '_Label9_1
        '
        Me._Label9_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label9_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label9_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label9_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.SetIndex(Me._Label9_1, CType(1, Short))
        Me._Label9_1.Location = New System.Drawing.Point(24, 256)
        Me._Label9_1.Name = "_Label9_1"
        Me._Label9_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label9_1.Size = New System.Drawing.Size(89, 17)
        Me._Label9_1.TabIndex = 82
        Me._Label9_1.Text = "Current Room"
        Me._Label9_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '__SSTab1_0_TabPage5
        '
        Me.__SSTab1_0_TabPage5.BackColor = System.Drawing.SystemColors.Control
        Me.__SSTab1_0_TabPage5.Controls.Add(Me.fraFloorSubstrate)
        Me.__SSTab1_0_TabPage5.Controls.Add(Me._lstRoomID_3)
        Me.__SSTab1_0_TabPage5.Controls.Add(Me._cmdOK_5)
        Me.__SSTab1_0_TabPage5.Controls.Add(Me.Frame7)
        Me.__SSTab1_0_TabPage5.Controls.Add(Me._Label9_2)
        Me.__SSTab1_0_TabPage5.Location = New System.Drawing.Point(4, 19)
        Me.__SSTab1_0_TabPage5.Name = "__SSTab1_0_TabPage5"
        Me.__SSTab1_0_TabPage5.Size = New System.Drawing.Size(509, 293)
        Me.__SSTab1_0_TabPage5.TabIndex = 5
        Me.__SSTab1_0_TabPage5.Text = "Floor Material"
        '
        'fraFloorSubstrate
        '
        Me.fraFloorSubstrate.BackColor = System.Drawing.SystemColors.Control
        Me.fraFloorSubstrate.Controls.Add(Me.txtFloorSubThickness)
        Me.fraFloorSubstrate.Controls.Add(Me.cmdPickFsubstrate)
        Me.fraFloorSubstrate.Controls.Add(Me.lblFloorSubstrate)
        Me.fraFloorSubstrate.Controls.Add(Me.Label12)
        Me.fraFloorSubstrate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraFloorSubstrate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraFloorSubstrate.Location = New System.Drawing.Point(48, 130)
        Me.fraFloorSubstrate.Name = "fraFloorSubstrate"
        Me.fraFloorSubstrate.Padding = New System.Windows.Forms.Padding(0)
        Me.fraFloorSubstrate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraFloorSubstrate.Size = New System.Drawing.Size(440, 105)
        Me.fraFloorSubstrate.TabIndex = 105
        Me.fraFloorSubstrate.TabStop = False
        Me.fraFloorSubstrate.Text = "Substrate Material"
        Me.fraFloorSubstrate.Visible = False
        '
        'txtFloorSubThickness
        '
        Me.txtFloorSubThickness.AcceptsReturn = True
        Me.txtFloorSubThickness.BackColor = System.Drawing.Color.White
        Me.txtFloorSubThickness.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFloorSubThickness.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFloorSubThickness.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFloorSubThickness.Location = New System.Drawing.Point(369, 24)
        Me.txtFloorSubThickness.MaxLength = 0
        Me.txtFloorSubThickness.Name = "txtFloorSubThickness"
        Me.txtFloorSubThickness.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFloorSubThickness.Size = New System.Drawing.Size(49, 20)
        Me.txtFloorSubThickness.TabIndex = 107
        Me.txtFloorSubThickness.Text = "150.0"
        '
        'lblFloorSubstrate
        '
        Me.lblFloorSubstrate.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblFloorSubstrate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFloorSubstrate.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFloorSubstrate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFloorSubstrate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFloorSubstrate.Location = New System.Drawing.Point(16, 24)
        Me.lblFloorSubstrate.Name = "lblFloorSubstrate"
        Me.lblFloorSubstrate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFloorSubstrate.Size = New System.Drawing.Size(238, 20)
        Me.lblFloorSubstrate.TabIndex = 109
        Me.lblFloorSubstrate.Text = "none"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.Transparent
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label12.Location = New System.Drawing.Point(249, 24)
        Me.Label12.Name = "Label12"
        Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label12.Size = New System.Drawing.Size(113, 25)
        Me.Label12.TabIndex = 108
        Me.Label12.Text = "Thickness (mm)"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lstRoomID_3
        '
        Me._lstRoomID_3.BackColor = System.Drawing.SystemColors.Window
        Me._lstRoomID_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lstRoomID_3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me._lstRoomID_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lstRoomID_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstRoomID.SetIndex(Me._lstRoomID_3, CType(3, Short))
        Me._lstRoomID_3.Location = New System.Drawing.Point(128, 256)
        Me._lstRoomID_3.Name = "_lstRoomID_3"
        Me._lstRoomID_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lstRoomID_3.Size = New System.Drawing.Size(89, 22)
        Me._lstRoomID_3.TabIndex = 4
        '
        '_cmdOK_5
        '
        Me._cmdOK_5.BackColor = System.Drawing.SystemColors.Control
        Me._cmdOK_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._cmdOK_5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cmdOK_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.SetIndex(Me._cmdOK_5, CType(5, Short))
        Me._cmdOK_5.Location = New System.Drawing.Point(415, 256)
        Me._cmdOK_5.Name = "_cmdOK_5"
        Me._cmdOK_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cmdOK_5.Size = New System.Drawing.Size(73, 25)
        Me._cmdOK_5.TabIndex = 5
        Me._cmdOK_5.Text = "OK"
        Me._cmdOK_5.UseVisualStyleBackColor = False
        '
        'Frame7
        '
        Me.Frame7.BackColor = System.Drawing.SystemColors.Control
        Me.Frame7.Controls.Add(Me.chkAddFloorSubstrate)
        Me.Frame7.Controls.Add(Me.txtFloorThickness)
        Me.Frame7.Controls.Add(Me.cmdPickFloor)
        Me.Frame7.Controls.Add(Me.Label1)
        Me.Frame7.Controls.Add(Me.lblFloorSurface)
        Me.Frame7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame7.Location = New System.Drawing.Point(48, 19)
        Me.Frame7.Name = "Frame7"
        Me.Frame7.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame7.Size = New System.Drawing.Size(440, 105)
        Me.Frame7.TabIndex = 79
        Me.Frame7.TabStop = False
        Me.Frame7.Text = "Surface Material"
        '
        'chkAddFloorSubstrate
        '
        Me.chkAddFloorSubstrate.BackColor = System.Drawing.SystemColors.Control
        Me.chkAddFloorSubstrate.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkAddFloorSubstrate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAddFloorSubstrate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkAddFloorSubstrate.Location = New System.Drawing.Point(289, 72)
        Me.chkAddFloorSubstrate.Name = "chkAddFloorSubstrate"
        Me.chkAddFloorSubstrate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkAddFloorSubstrate.Size = New System.Drawing.Size(113, 17)
        Me.chkAddFloorSubstrate.TabIndex = 110
        Me.chkAddFloorSubstrate.Text = "Include Substrate"
        Me.chkAddFloorSubstrate.UseVisualStyleBackColor = False
        '
        'txtFloorThickness
        '
        Me.txtFloorThickness.AcceptsReturn = True
        Me.txtFloorThickness.BackColor = System.Drawing.Color.White
        Me.txtFloorThickness.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFloorThickness.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFloorThickness.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFloorThickness.Location = New System.Drawing.Point(369, 24)
        Me.txtFloorThickness.MaxLength = 0
        Me.txtFloorThickness.Name = "txtFloorThickness"
        Me.txtFloorThickness.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFloorThickness.Size = New System.Drawing.Size(49, 20)
        Me.txtFloorThickness.TabIndex = 2
        Me.txtFloorThickness.Text = "13.0"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.Location = New System.Drawing.Point(249, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(113, 25)
        Me.Label1.TabIndex = 80
        Me.Label1.Text = "Thickness (mm)"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblFloorSurface
        '
        Me.lblFloorSurface.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblFloorSurface.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFloorSurface.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFloorSurface.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFloorSurface.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFloorSurface.Location = New System.Drawing.Point(16, 24)
        Me.lblFloorSurface.Name = "lblFloorSurface"
        Me.lblFloorSurface.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFloorSurface.Size = New System.Drawing.Size(238, 20)
        Me.lblFloorSurface.TabIndex = 1
        Me.lblFloorSurface.Text = "none"
        '
        '_Label9_2
        '
        Me._Label9_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label9_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label9_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label9_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.SetIndex(Me._Label9_2, CType(2, Short))
        Me._Label9_2.Location = New System.Drawing.Point(24, 256)
        Me._Label9_2.Name = "_Label9_2"
        Me._Label9_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label9_2.Size = New System.Drawing.Size(89, 17)
        Me._Label9_2.TabIndex = 83
        Me._Label9_2.Text = "Current Room"
        Me._Label9_2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'cmdOK
        '
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'frmDescribeRoom
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(521, 325)
        Me.Controls.Add(Me._SSTab1_0)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(419, 195)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmDescribeRoom"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Building / Room Specification"
        Me.TopMost = True
        Me._SSTab1_0.ResumeLayout(False)
        Me.__SSTab1_0_TabPage3.ResumeLayout(False)
        Me.fraWallMaterial.ResumeLayout(False)
        Me.fraWallMaterial.PerformLayout()
        Me.fraWallSubstrate.ResumeLayout(False)
        Me.fraWallSubstrate.PerformLayout()
        Me.__SSTab1_0_TabPage4.ResumeLayout(False)
        Me.Frame6.ResumeLayout(False)
        Me.Frame6.PerformLayout()
        Me.fraCeilingSubstrate.ResumeLayout(False)
        Me.fraCeilingSubstrate.PerformLayout()
        Me.__SSTab1_0_TabPage5.ResumeLayout(False)
        Me.fraFloorSubstrate.ResumeLayout(False)
        Me.fraFloorSubstrate.PerformLayout()
        Me.Frame7.ResumeLayout(False)
        Me.Frame7.PerformLayout()
        CType(Me.Label9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SSTab1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdOK, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lstRoomID, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
#End Region
End Class