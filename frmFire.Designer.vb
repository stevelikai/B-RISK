<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmFire
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
	Public WithEvents lstObjectID As System.Windows.Forms.ComboBox
    Public WithEvents cmdPlot As System.Windows.Forms.Button
	Public WithEvents cmdViewData As System.Windows.Forms.Button
	Public WithEvents lstFireRoom2 As System.Windows.Forms.ComboBox
    Public WithEvents txtFireHeight As System.Windows.Forms.TextBox
	Public WithEvents cmdDeleteObject As System.Windows.Forms.Button
	Public WithEvents txtCO2Yield As System.Windows.Forms.TextBox
	Public WithEvents txtEnergyYield As System.Windows.Forms.TextBox
	Public WithEvents cmdAddObject As System.Windows.Forms.Button
	Public WithEvents txtSootYield As System.Windows.Forms.TextBox
	Public WithEvents cmdClose As System.Windows.Forms.Button
	Public WithEvents txtHCNYield As System.Windows.Forms.TextBox
	Public WithEvents lstObjectLocation As System.Windows.Forms.ComboBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents lblObjectDescription As System.Windows.Forms.Label
	Public WithEvents lblSourceID As System.Windows.Forms.Label
	Public WithEvents lblEnergyYield As System.Windows.Forms.Label
	Public WithEvents lblCO2Yield As System.Windows.Forms.Label
	Public WithEvents lblSootYield As System.Windows.Forms.Label
	Public WithEvents lblWaterVaporYield As System.Windows.Forms.Label
	Public WithEvents lblFireHeight As System.Windows.Forms.Label
	Public WithEvents lblObjectLocation As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdPlot = New System.Windows.Forms.Button
        Me.cmdViewData = New System.Windows.Forms.Button
        Me.cmdDeleteObject = New System.Windows.Forms.Button
        Me.cmdAddObject = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.lstObjectID = New System.Windows.Forms.ComboBox
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.lblObjectLocation = New System.Windows.Forms.Label
        Me.lstFireRoom2 = New System.Windows.Forms.ComboBox
        Me.lstObjectLocation = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblSourceID = New System.Windows.Forms.Label
        Me.txtFireHeight = New System.Windows.Forms.TextBox
        Me.txtCO2Yield = New System.Windows.Forms.TextBox
        Me.txtEnergyYield = New System.Windows.Forms.TextBox
        Me.txtSootYield = New System.Windows.Forms.TextBox
        Me.txtHCNYield = New System.Windows.Forms.TextBox
        Me.lblObjectDescription = New System.Windows.Forms.Label
        Me.lblEnergyYield = New System.Windows.Forms.Label
        Me.lblCO2Yield = New System.Windows.Forms.Label
        Me.lblSootYield = New System.Windows.Forms.Label
        Me.lblWaterVaporYield = New System.Windows.Forms.Label
        Me.lblFireHeight = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Frame1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdPlot
        '
        Me.cmdPlot.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPlot.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPlot.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPlot.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPlot.Location = New System.Drawing.Point(149, 260)
        Me.cmdPlot.Name = "cmdPlot"
        Me.cmdPlot.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPlot.Size = New System.Drawing.Size(89, 25)
        Me.cmdPlot.TabIndex = 9
        Me.cmdPlot.Tag = "View the graph of heat release versus time"
        Me.cmdPlot.Text = "Plot"
        Me.ToolTip1.SetToolTip(Me.cmdPlot, "plot the time - heat release data")
        Me.cmdPlot.UseVisualStyleBackColor = False
        '
        'cmdViewData
        '
        Me.cmdViewData.BackColor = System.Drawing.SystemColors.Control
        Me.cmdViewData.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdViewData.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdViewData.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdViewData.Location = New System.Drawing.Point(149, 229)
        Me.cmdViewData.Name = "cmdViewData"
        Me.cmdViewData.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdViewData.Size = New System.Drawing.Size(89, 25)
        Me.cmdViewData.TabIndex = 11
        Me.cmdViewData.Text = "View Data"
        Me.ToolTip1.SetToolTip(Me.cmdViewData, "show time - heat release data")
        Me.cmdViewData.UseVisualStyleBackColor = False
        Me.cmdViewData.Visible = False
        '
        'cmdDeleteObject
        '
        Me.cmdDeleteObject.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDeleteObject.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDeleteObject.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDeleteObject.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDeleteObject.Location = New System.Drawing.Point(125, 184)
        Me.cmdDeleteObject.Name = "cmdDeleteObject"
        Me.cmdDeleteObject.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDeleteObject.Size = New System.Drawing.Size(113, 25)
        Me.cmdDeleteObject.TabIndex = 10
        Me.cmdDeleteObject.Tag = "Delete the currently selected burning object "
        Me.cmdDeleteObject.Text = "Delete Current Fire"
        Me.ToolTip1.SetToolTip(Me.cmdDeleteObject, "remove the current fire object from the simulation")
        Me.cmdDeleteObject.UseVisualStyleBackColor = False
        '
        'cmdAddObject
        '
        Me.cmdAddObject.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAddObject.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAddObject.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddObject.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAddObject.Location = New System.Drawing.Point(125, 152)
        Me.cmdAddObject.Name = "cmdAddObject"
        Me.cmdAddObject.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAddObject.Size = New System.Drawing.Size(113, 25)
        Me.cmdAddObject.TabIndex = 13
        Me.cmdAddObject.Tag = "Add an additional burning object"
        Me.cmdAddObject.Text = "Add New Fire"
        Me.ToolTip1.SetToolTip(Me.cmdAddObject, "pick an additional fire from the database to add to this simulation")
        Me.cmdAddObject.UseVisualStyleBackColor = False
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Location = New System.Drawing.Point(125, 344)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(113, 25)
        Me.cmdClose.TabIndex = 14
        Me.cmdClose.Text = "Close"
        Me.ToolTip1.SetToolTip(Me.cmdClose, "close the form")
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'lstObjectID
        '
        Me.lstObjectID.BackColor = System.Drawing.SystemColors.Window
        Me.lstObjectID.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstObjectID.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstObjectID.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstObjectID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstObjectID.Location = New System.Drawing.Point(133, 26)
        Me.lstObjectID.Name = "lstObjectID"
        Me.lstObjectID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstObjectID.Size = New System.Drawing.Size(105, 22)
        Me.lstObjectID.TabIndex = 6
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.cmdPlot)
        Me.Frame1.Controls.Add(Me.lstObjectID)
        Me.Frame1.Controls.Add(Me.lblObjectLocation)
        Me.Frame1.Controls.Add(Me.cmdViewData)
        Me.Frame1.Controls.Add(Me.lstFireRoom2)
        Me.Frame1.Controls.Add(Me.cmdDeleteObject)
        Me.Frame1.Controls.Add(Me.cmdAddObject)
        Me.Frame1.Controls.Add(Me.cmdClose)
        Me.Frame1.Controls.Add(Me.lstObjectLocation)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.Controls.Add(Me.lblSourceID)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(243, 12)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(251, 380)
        Me.Frame1.TabIndex = 0
        Me.Frame1.TabStop = False
        '
        'lblObjectLocation
        '
        Me.lblObjectLocation.BackColor = System.Drawing.Color.Transparent
        Me.lblObjectLocation.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblObjectLocation.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblObjectLocation.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblObjectLocation.Location = New System.Drawing.Point(-3, 63)
        Me.lblObjectLocation.Name = "lblObjectLocation"
        Me.lblObjectLocation.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblObjectLocation.Size = New System.Drawing.Size(121, 17)
        Me.lblObjectLocation.TabIndex = 15
        Me.lblObjectLocation.Text = "Object Location"
        Me.lblObjectLocation.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lstFireRoom2
        '
        Me.lstFireRoom2.BackColor = System.Drawing.SystemColors.Window
        Me.lstFireRoom2.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstFireRoom2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstFireRoom2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstFireRoom2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstFireRoom2.Location = New System.Drawing.Point(133, 95)
        Me.lstFireRoom2.MaxDropDownItems = 10
        Me.lstFireRoom2.Name = "lstFireRoom2"
        Me.lstFireRoom2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstFireRoom2.Size = New System.Drawing.Size(105, 22)
        Me.lstFireRoom2.TabIndex = 8
        '
        'lstObjectLocation
        '
        Me.lstObjectLocation.BackColor = System.Drawing.SystemColors.Window
        Me.lstObjectLocation.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstObjectLocation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstObjectLocation.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstObjectLocation.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstObjectLocation.Location = New System.Drawing.Point(133, 63)
        Me.lstObjectLocation.Name = "lstObjectLocation"
        Me.lstObjectLocation.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstObjectLocation.Size = New System.Drawing.Size(105, 22)
        Me.lstObjectLocation.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(5, 95)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(113, 17)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "Fire Room Location"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblSourceID
        '
        Me.lblSourceID.BackColor = System.Drawing.Color.Transparent
        Me.lblSourceID.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSourceID.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSourceID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblSourceID.Location = New System.Drawing.Point(-3, 31)
        Me.lblSourceID.Name = "lblSourceID"
        Me.lblSourceID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSourceID.Size = New System.Drawing.Size(121, 17)
        Me.lblSourceID.TabIndex = 21
        Me.lblSourceID.Text = "Fire Object ID"
        Me.lblSourceID.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtFireHeight
        '
        Me.txtFireHeight.AcceptsReturn = True
        Me.txtFireHeight.BackColor = System.Drawing.Color.White
        Me.txtFireHeight.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFireHeight.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFireHeight.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFireHeight.Location = New System.Drawing.Point(164, 193)
        Me.txtFireHeight.MaxLength = 0
        Me.txtFireHeight.Name = "txtFireHeight"
        Me.txtFireHeight.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFireHeight.Size = New System.Drawing.Size(49, 20)
        Me.txtFireHeight.TabIndex = 5
        Me.txtFireHeight.Text = "0"
        '
        'txtCO2Yield
        '
        Me.txtCO2Yield.AcceptsReturn = True
        Me.txtCO2Yield.BackColor = System.Drawing.Color.White
        Me.txtCO2Yield.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCO2Yield.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCO2Yield.ForeColor = System.Drawing.Color.Black
        Me.txtCO2Yield.Location = New System.Drawing.Point(164, 97)
        Me.txtCO2Yield.MaxLength = 0
        Me.txtCO2Yield.Name = "txtCO2Yield"
        Me.txtCO2Yield.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCO2Yield.Size = New System.Drawing.Size(49, 20)
        Me.txtCO2Yield.TabIndex = 2
        Me.txtCO2Yield.Tag = "Carbon dioxide yield "
        Me.txtCO2Yield.Text = "1.2"
        '
        'txtEnergyYield
        '
        Me.txtEnergyYield.AcceptsReturn = True
        Me.txtEnergyYield.BackColor = System.Drawing.Color.White
        Me.txtEnergyYield.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEnergyYield.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEnergyYield.ForeColor = System.Drawing.Color.Black
        Me.txtEnergyYield.Location = New System.Drawing.Point(164, 65)
        Me.txtEnergyYield.MaxLength = 0
        Me.txtEnergyYield.Name = "txtEnergyYield"
        Me.txtEnergyYield.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEnergyYield.Size = New System.Drawing.Size(49, 20)
        Me.txtEnergyYield.TabIndex = 1
        Me.txtEnergyYield.Tag = "Energy yield "
        Me.txtEnergyYield.Text = "16"
        '
        'txtSootYield
        '
        Me.txtSootYield.AcceptsReturn = True
        Me.txtSootYield.BackColor = System.Drawing.Color.White
        Me.txtSootYield.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSootYield.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSootYield.ForeColor = System.Drawing.Color.Black
        Me.txtSootYield.Location = New System.Drawing.Point(164, 129)
        Me.txtSootYield.MaxLength = 0
        Me.txtSootYield.Name = "txtSootYield"
        Me.txtSootYield.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSootYield.Size = New System.Drawing.Size(49, 20)
        Me.txtSootYield.TabIndex = 3
        Me.txtSootYield.Tag = "Soot yield "
        Me.txtSootYield.Text = "0.03"
        '
        'txtHCNYield
        '
        Me.txtHCNYield.AcceptsReturn = True
        Me.txtHCNYield.BackColor = System.Drawing.Color.White
        Me.txtHCNYield.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtHCNYield.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHCNYield.ForeColor = System.Drawing.Color.Black
        Me.txtHCNYield.Location = New System.Drawing.Point(164, 161)
        Me.txtHCNYield.MaxLength = 0
        Me.txtHCNYield.Name = "txtHCNYield"
        Me.txtHCNYield.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtHCNYield.Size = New System.Drawing.Size(49, 20)
        Me.txtHCNYield.TabIndex = 4
        Me.txtHCNYield.Tag = "HCNYield"
        Me.txtHCNYield.Text = "0"
        '
        'lblObjectDescription
        '
        Me.lblObjectDescription.BackColor = System.Drawing.SystemColors.Control
        Me.lblObjectDescription.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblObjectDescription.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblObjectDescription.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblObjectDescription.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblObjectDescription.Location = New System.Drawing.Point(7, 27)
        Me.lblObjectDescription.Name = "lblObjectDescription"
        Me.lblObjectDescription.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblObjectDescription.Size = New System.Drawing.Size(206, 21)
        Me.lblObjectDescription.TabIndex = 23
        Me.lblObjectDescription.Text = "Object Description"
        Me.lblObjectDescription.UseMnemonic = False
        '
        'lblEnergyYield
        '
        Me.lblEnergyYield.BackColor = System.Drawing.Color.Transparent
        Me.lblEnergyYield.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEnergyYield.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEnergyYield.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblEnergyYield.Location = New System.Drawing.Point(12, 65)
        Me.lblEnergyYield.Name = "lblEnergyYield"
        Me.lblEnergyYield.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEnergyYield.Size = New System.Drawing.Size(137, 25)
        Me.lblEnergyYield.TabIndex = 20
        Me.lblEnergyYield.Text = "Energy Yield (kJ/g)"
        Me.lblEnergyYield.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblCO2Yield
        '
        Me.lblCO2Yield.BackColor = System.Drawing.Color.Transparent
        Me.lblCO2Yield.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCO2Yield.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCO2Yield.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblCO2Yield.Location = New System.Drawing.Point(4, 97)
        Me.lblCO2Yield.Name = "lblCO2Yield"
        Me.lblCO2Yield.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCO2Yield.Size = New System.Drawing.Size(145, 25)
        Me.lblCO2Yield.TabIndex = 19
        Me.lblCO2Yield.Text = "CO2 Yield (kg/kg fuel)"
        Me.lblCO2Yield.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblSootYield
        '
        Me.lblSootYield.BackColor = System.Drawing.Color.Transparent
        Me.lblSootYield.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSootYield.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSootYield.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblSootYield.Location = New System.Drawing.Point(4, 129)
        Me.lblSootYield.Name = "lblSootYield"
        Me.lblSootYield.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSootYield.Size = New System.Drawing.Size(145, 33)
        Me.lblSootYield.TabIndex = 18
        Me.lblSootYield.Text = "Soot Yield (kg/kg fuel)"
        Me.lblSootYield.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblWaterVaporYield
        '
        Me.lblWaterVaporYield.BackColor = System.Drawing.Color.Transparent
        Me.lblWaterVaporYield.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblWaterVaporYield.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWaterVaporYield.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblWaterVaporYield.Location = New System.Drawing.Point(4, 161)
        Me.lblWaterVaporYield.Name = "lblWaterVaporYield"
        Me.lblWaterVaporYield.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblWaterVaporYield.Size = New System.Drawing.Size(145, 25)
        Me.lblWaterVaporYield.TabIndex = 17
        Me.lblWaterVaporYield.Text = "HCN Yield (kg/kg fuel)"
        Me.lblWaterVaporYield.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblFireHeight
        '
        Me.lblFireHeight.BackColor = System.Drawing.Color.Transparent
        Me.lblFireHeight.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFireHeight.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFireHeight.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblFireHeight.Location = New System.Drawing.Point(4, 193)
        Me.lblFireHeight.Name = "lblFireHeight"
        Me.lblFireHeight.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFireHeight.Size = New System.Drawing.Size(145, 33)
        Me.lblFireHeight.TabIndex = 16
        Me.lblFireHeight.Text = "Fire Height above Floor (m)"
        Me.lblFireHeight.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblEnergyYield)
        Me.GroupBox1.Controls.Add(Me.lblFireHeight)
        Me.GroupBox1.Controls.Add(Me.txtFireHeight)
        Me.GroupBox1.Controls.Add(Me.lblWaterVaporYield)
        Me.GroupBox1.Controls.Add(Me.lblSootYield)
        Me.GroupBox1.Controls.Add(Me.txtCO2Yield)
        Me.GroupBox1.Controls.Add(Me.lblCO2Yield)
        Me.GroupBox1.Controls.Add(Me.txtEnergyYield)
        Me.GroupBox1.Controls.Add(Me.txtHCNYield)
        Me.GroupBox1.Controls.Add(Me.lblObjectDescription)
        Me.GroupBox1.Controls.Add(Me.txtSootYield)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(225, 380)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Current Object"
        '
        'frmFire
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(507, 404)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(3, 18)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmFire"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Selected Fire Objects to use in this Project Only"
        Me.Frame1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
#End Region
End Class