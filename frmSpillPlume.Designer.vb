<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmSpillPlume
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
	Public WithEvents optDoubleSided As System.Windows.Forms.RadioButton
	Public WithEvents optSingleSided As System.Windows.Forms.RadioButton
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents cmdClose As System.Windows.Forms.Button
	Public WithEvents OptHarrison As System.Windows.Forms.RadioButton
	Public WithEvents OptBRE368 As System.Windows.Forms.RadioButton
	Public WithEvents Frame26 As System.Windows.Forms.GroupBox
	Public WithEvents optSpillPlumeBalc As System.Windows.Forms.RadioButton
	Public WithEvents optSpillPlumeEdge As System.Windows.Forms.RadioButton
	Public WithEvents txtDownstand As System.Windows.Forms.TextBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents fraBSOptions As System.Windows.Forms.GroupBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.optDoubleSided = New System.Windows.Forms.RadioButton
        Me.optSingleSided = New System.Windows.Forms.RadioButton
        Me.cmdClose = New System.Windows.Forms.Button
        Me.Frame26 = New System.Windows.Forms.GroupBox
        Me.OptHarrison = New System.Windows.Forms.RadioButton
        Me.OptBRE368 = New System.Windows.Forms.RadioButton
        Me.fraBSOptions = New System.Windows.Forms.GroupBox
        Me.optSpillPlumeBalc = New System.Windows.Forms.RadioButton
        Me.optSpillPlumeEdge = New System.Windows.Forms.RadioButton
        Me.txtDownstand = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Frame1.SuspendLayout()
        Me.Frame26.SuspendLayout()
        Me.fraBSOptions.SuspendLayout()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.optDoubleSided)
        Me.Frame1.Controls.Add(Me.optSingleSided)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(8, 224)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(337, 65)
        Me.Frame1.TabIndex = 9
        Me.Frame1.TabStop = False
        '
        'optDoubleSided
        '
        Me.optDoubleSided.BackColor = System.Drawing.SystemColors.Control
        Me.optDoubleSided.Checked = True
        Me.optDoubleSided.Cursor = System.Windows.Forms.Cursors.Default
        Me.optDoubleSided.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optDoubleSided.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDoubleSided.Location = New System.Drawing.Point(168, 32)
        Me.optDoubleSided.Name = "optDoubleSided"
        Me.optDoubleSided.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optDoubleSided.Size = New System.Drawing.Size(145, 17)
        Me.optDoubleSided.TabIndex = 11
        Me.optDoubleSided.TabStop = True
        Me.optDoubleSided.Text = "Double-sided Spill Plume"
        Me.optDoubleSided.UseVisualStyleBackColor = False
        '
        'optSingleSided
        '
        Me.optSingleSided.BackColor = System.Drawing.SystemColors.Control
        Me.optSingleSided.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSingleSided.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optSingleSided.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optSingleSided.Location = New System.Drawing.Point(16, 32)
        Me.optSingleSided.Name = "optSingleSided"
        Me.optSingleSided.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSingleSided.Size = New System.Drawing.Size(153, 17)
        Me.optSingleSided.TabIndex = 10
        Me.optSingleSided.TabStop = True
        Me.optSingleSided.Text = "Single-sided Spill Plume"
        Me.optSingleSided.UseVisualStyleBackColor = False
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Location = New System.Drawing.Point(264, 296)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(81, 33)
        Me.cmdClose.TabIndex = 8
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'Frame26
        '
        Me.Frame26.BackColor = System.Drawing.SystemColors.Control
        Me.Frame26.Controls.Add(Me.OptHarrison)
        Me.Frame26.Controls.Add(Me.OptBRE368)
        Me.Frame26.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame26.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame26.Location = New System.Drawing.Point(8, 8)
        Me.Frame26.Name = "Frame26"
        Me.Frame26.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame26.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame26.Size = New System.Drawing.Size(337, 57)
        Me.Frame26.TabIndex = 5
        Me.Frame26.TabStop = False
        Me.Frame26.Text = "Balcony Spill Plume Model"
        '
        'OptHarrison
        '
        Me.OptHarrison.BackColor = System.Drawing.SystemColors.Control
        Me.OptHarrison.Checked = True
        Me.OptHarrison.Cursor = System.Windows.Forms.Cursors.Default
        Me.OptHarrison.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OptHarrison.ForeColor = System.Drawing.SystemColors.ControlText
        Me.OptHarrison.Location = New System.Drawing.Point(120, 24)
        Me.OptHarrison.Name = "OptHarrison"
        Me.OptHarrison.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.OptHarrison.Size = New System.Drawing.Size(161, 17)
        Me.OptHarrison.TabIndex = 7
        Me.OptHarrison.TabStop = True
        Me.OptHarrison.Text = "Harrison's Correlation (2004)"
        Me.OptHarrison.UseVisualStyleBackColor = False
        '
        'OptBRE368
        '
        Me.OptBRE368.BackColor = System.Drawing.SystemColors.Control
        Me.OptBRE368.Cursor = System.Windows.Forms.Cursors.Default
        Me.OptBRE368.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OptBRE368.ForeColor = System.Drawing.SystemColors.ControlText
        Me.OptBRE368.Location = New System.Drawing.Point(16, 24)
        Me.OptBRE368.Name = "OptBRE368"
        Me.OptBRE368.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.OptBRE368.Size = New System.Drawing.Size(105, 17)
        Me.OptBRE368.TabIndex = 6
        Me.OptBRE368.TabStop = True
        Me.OptBRE368.Text = "BRE 368"
        Me.OptBRE368.UseVisualStyleBackColor = False
        '
        'fraBSOptions
        '
        Me.fraBSOptions.BackColor = System.Drawing.SystemColors.Control
        Me.fraBSOptions.Controls.Add(Me.optSpillPlumeBalc)
        Me.fraBSOptions.Controls.Add(Me.optSpillPlumeEdge)
        Me.fraBSOptions.Controls.Add(Me.txtDownstand)
        Me.fraBSOptions.Controls.Add(Me.Label1)
        Me.fraBSOptions.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraBSOptions.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraBSOptions.Location = New System.Drawing.Point(8, 72)
        Me.fraBSOptions.Name = "fraBSOptions"
        Me.fraBSOptions.Padding = New System.Windows.Forms.Padding(0)
        Me.fraBSOptions.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraBSOptions.Size = New System.Drawing.Size(337, 153)
        Me.fraBSOptions.TabIndex = 0
        Me.fraBSOptions.TabStop = False
        Me.fraBSOptions.Text = "Balcony Spill Plume Options"
        '
        'optSpillPlumeBalc
        '
        Me.optSpillPlumeBalc.BackColor = System.Drawing.SystemColors.Control
        Me.optSpillPlumeBalc.Checked = True
        Me.optSpillPlumeBalc.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSpillPlumeBalc.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optSpillPlumeBalc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optSpillPlumeBalc.Location = New System.Drawing.Point(32, 64)
        Me.optSpillPlumeBalc.Name = "optSpillPlumeBalc"
        Me.optSpillPlumeBalc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSpillPlumeBalc.Size = New System.Drawing.Size(273, 33)
        Me.optSpillPlumeBalc.TabIndex = 4
        Me.optSpillPlumeBalc.TabStop = True
        Me.optSpillPlumeBalc.Text = "Balcony extends beyond the compartment opening"
        Me.optSpillPlumeBalc.UseVisualStyleBackColor = False
        '
        'optSpillPlumeEdge
        '
        Me.optSpillPlumeEdge.BackColor = System.Drawing.SystemColors.Control
        Me.optSpillPlumeEdge.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSpillPlumeEdge.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optSpillPlumeEdge.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optSpillPlumeEdge.Location = New System.Drawing.Point(32, 32)
        Me.optSpillPlumeEdge.Name = "optSpillPlumeEdge"
        Me.optSpillPlumeEdge.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSpillPlumeEdge.Size = New System.Drawing.Size(249, 33)
        Me.optSpillPlumeEdge.TabIndex = 3
        Me.optSpillPlumeEdge.TabStop = True
        Me.optSpillPlumeEdge.Text = "Compartment opening is at the spill edge"
        Me.optSpillPlumeEdge.UseVisualStyleBackColor = False
        '
        'txtDownstand
        '
        Me.txtDownstand.AcceptsReturn = True
        Me.txtDownstand.BackColor = System.Drawing.SystemColors.Window
        Me.txtDownstand.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDownstand.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDownstand.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDownstand.Location = New System.Drawing.Point(232, 112)
        Me.txtDownstand.MaxLength = 0
        Me.txtDownstand.Name = "txtDownstand"
        Me.txtDownstand.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDownstand.Size = New System.Drawing.Size(49, 19)
        Me.txtDownstand.TabIndex = 1
        Me.txtDownstand.Text = "0"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(16, 112)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(201, 25)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Depth of downstand at the spill edge (m)"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'frmSpillPlume
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(362, 335)
        Me.ControlBox = False
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.Frame26)
        Me.Controls.Add(Me.fraBSOptions)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 28)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSpillPlume"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.Text = "Spill Plumes"
        Me.Frame1.ResumeLayout(False)
        Me.Frame26.ResumeLayout(False)
        Me.fraBSOptions.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class