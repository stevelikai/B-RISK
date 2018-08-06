<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmQuintiere
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
	Public WithEvents optNoneFloor As System.Windows.Forms.RadioButton
	Public WithEvents optWindAided As System.Windows.Forms.RadioButton
	Public WithEvents optOpposedFlow As System.Windows.Forms.RadioButton
	Public WithEvents Frame3 As System.Windows.Forms.GroupBox
	Public WithEvents optFTP As System.Windows.Forms.RadioButton
	Public WithEvents optJanssens As System.Windows.Forms.RadioButton
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents cmdClose As System.Windows.Forms.Button
	Public WithEvents chkDisableLateralSpread As System.Windows.Forms.CheckBox
	Public WithEvents chkSpreadAdjacentRoom As System.Windows.Forms.CheckBox
	Public WithEvents optUseOneCurve As System.Windows.Forms.RadioButton
	Public WithEvents optUseAllCurves As System.Windows.Forms.RadioButton
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtVFSlimit = New System.Windows.Forms.TextBox()
        Me.txtHFSlimit = New System.Windows.Forms.TextBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.optNoneFloor = New System.Windows.Forms.RadioButton()
        Me.optWindAided = New System.Windows.Forms.RadioButton()
        Me.optOpposedFlow = New System.Windows.Forms.RadioButton()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.optFTP = New System.Windows.Forms.RadioButton()
        Me.optJanssens = New System.Windows.Forms.RadioButton()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.chkDisableLateralSpread = New System.Windows.Forms.CheckBox()
        Me.chkSpreadAdjacentRoom = New System.Windows.Forms.CheckBox()
        Me.optUseOneCurve = New System.Windows.Forms.RadioButton()
        Me.optUseAllCurves = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.chkPessimiseWall = New System.Windows.Forms.CheckBox()
        Me.LblVFS = New System.Windows.Forms.Label()
        Me.LblHFS = New System.Windows.Forms.Label()
        Me.lblCeilingpercent = New System.Windows.Forms.Label()
        Me.NumericCeilingPercent = New System.Windows.Forms.NumericUpDown()
        Me.lblwallpercent = New System.Windows.Forms.Label()
        Me.NumericWallPercent = New System.Windows.Forms.NumericUpDown()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.Frame3.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.NumericCeilingPercent, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericWallPercent, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtVFSlimit
        '
        Me.txtVFSlimit.Location = New System.Drawing.Point(215, 128)
        Me.txtVFSlimit.Name = "txtVFSlimit"
        Me.txtVFSlimit.Size = New System.Drawing.Size(50, 20)
        Me.txtVFSlimit.TabIndex = 5
        Me.txtVFSlimit.Text = "0"
        Me.ToolTip1.SetToolTip(Me.txtVFSlimit, "Maximum vertical distance from the floor that bounds the extent of flame spread p" &
        "ermitted.")
        '
        'txtHFSlimit
        '
        Me.txtHFSlimit.Location = New System.Drawing.Point(215, 102)
        Me.txtHFSlimit.Name = "txtHFSlimit"
        Me.txtHFSlimit.Size = New System.Drawing.Size(50, 20)
        Me.txtHFSlimit.TabIndex = 4
        Me.txtHFSlimit.Text = "0"
        Me.ToolTip1.SetToolTip(Me.txtHFSlimit, "Maximum horizontal distance from the burner that bounds the extent of flame sprea" &
        "d permitted. ")
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.optNoneFloor)
        Me.Frame3.Controls.Add(Me.optWindAided)
        Me.Frame3.Controls.Add(Me.optOpposedFlow)
        Me.Frame3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame3.Location = New System.Drawing.Point(8, 425)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(321, 48)
        Me.Frame3.TabIndex = 9
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Floor Covering Flame Spread Model"
        Me.Frame3.Visible = False
        '
        'optNoneFloor
        '
        Me.optNoneFloor.BackColor = System.Drawing.SystemColors.Control
        Me.optNoneFloor.Checked = True
        Me.optNoneFloor.Cursor = System.Windows.Forms.Cursors.Default
        Me.optNoneFloor.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optNoneFloor.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optNoneFloor.Location = New System.Drawing.Point(16, 24)
        Me.optNoneFloor.Name = "optNoneFloor"
        Me.optNoneFloor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optNoneFloor.Size = New System.Drawing.Size(89, 17)
        Me.optNoneFloor.TabIndex = 12
        Me.optNoneFloor.TabStop = True
        Me.optNoneFloor.Text = "None"
        Me.optNoneFloor.UseVisualStyleBackColor = False
        '
        'optWindAided
        '
        Me.optWindAided.BackColor = System.Drawing.SystemColors.Control
        Me.optWindAided.Cursor = System.Windows.Forms.Cursors.Default
        Me.optWindAided.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optWindAided.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optWindAided.Location = New System.Drawing.Point(216, 24)
        Me.optWindAided.Name = "optWindAided"
        Me.optWindAided.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optWindAided.Size = New System.Drawing.Size(105, 17)
        Me.optWindAided.TabIndex = 11
        Me.optWindAided.TabStop = True
        Me.optWindAided.Text = "Wind-aided Flow"
        Me.optWindAided.UseVisualStyleBackColor = False
        '
        'optOpposedFlow
        '
        Me.optOpposedFlow.BackColor = System.Drawing.SystemColors.Control
        Me.optOpposedFlow.Cursor = System.Windows.Forms.Cursors.Default
        Me.optOpposedFlow.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optOpposedFlow.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optOpposedFlow.Location = New System.Drawing.Point(104, 24)
        Me.optOpposedFlow.Name = "optOpposedFlow"
        Me.optOpposedFlow.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optOpposedFlow.Size = New System.Drawing.Size(105, 17)
        Me.optOpposedFlow.TabIndex = 10
        Me.optOpposedFlow.TabStop = True
        Me.optOpposedFlow.Text = "Opposed Flow "
        Me.optOpposedFlow.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.optFTP)
        Me.Frame2.Controls.Add(Me.optJanssens)
        Me.Frame2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(7, 326)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(322, 71)
        Me.Frame2.TabIndex = 5
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Ignition Correlations"
        '
        'optFTP
        '
        Me.optFTP.BackColor = System.Drawing.SystemColors.Control
        Me.optFTP.Checked = True
        Me.optFTP.Cursor = System.Windows.Forms.Cursors.Default
        Me.optFTP.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optFTP.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optFTP.Location = New System.Drawing.Point(16, 47)
        Me.optFTP.Name = "optFTP"
        Me.optFTP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optFTP.Size = New System.Drawing.Size(265, 17)
        Me.optFTP.TabIndex = 7
        Me.optFTP.TabStop = True
        Me.optFTP.Text = "Flux Time Product method (default)"
        Me.optFTP.UseVisualStyleBackColor = False
        '
        'optJanssens
        '
        Me.optJanssens.BackColor = System.Drawing.SystemColors.Control
        Me.optJanssens.Cursor = System.Windows.Forms.Cursors.Default
        Me.optJanssens.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optJanssens.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optJanssens.Location = New System.Drawing.Point(16, 24)
        Me.optJanssens.Name = "optJanssens"
        Me.optJanssens.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optJanssens.Size = New System.Drawing.Size(209, 17)
        Me.optJanssens.TabIndex = 6
        Me.optJanssens.TabStop = True
        Me.optJanssens.Text = "Grenier and Janssens"
        Me.optJanssens.UseVisualStyleBackColor = False
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Location = New System.Drawing.Point(256, 403)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(73, 25)
        Me.cmdClose.TabIndex = 4
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.chkDisableLateralSpread)
        Me.Frame1.Controls.Add(Me.chkSpreadAdjacentRoom)
        Me.Frame1.Controls.Add(Me.optUseOneCurve)
        Me.Frame1.Controls.Add(Me.optUseAllCurves)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(8, 8)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(321, 152)
        Me.Frame1.TabIndex = 0
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Options for Flame Spread Model"
        '
        'chkDisableLateralSpread
        '
        Me.chkDisableLateralSpread.BackColor = System.Drawing.SystemColors.Control
        Me.chkDisableLateralSpread.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDisableLateralSpread.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkDisableLateralSpread.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDisableLateralSpread.Location = New System.Drawing.Point(15, 97)
        Me.chkDisableLateralSpread.Name = "chkDisableLateralSpread"
        Me.chkDisableLateralSpread.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDisableLateralSpread.Size = New System.Drawing.Size(241, 17)
        Me.chkDisableLateralSpread.TabIndex = 8
        Me.chkDisableLateralSpread.Text = "Disable Lateral Flame Spread on Walls"
        Me.chkDisableLateralSpread.UseVisualStyleBackColor = False
        '
        'chkSpreadAdjacentRoom
        '
        Me.chkSpreadAdjacentRoom.BackColor = System.Drawing.SystemColors.Control
        Me.chkSpreadAdjacentRoom.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkSpreadAdjacentRoom.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSpreadAdjacentRoom.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkSpreadAdjacentRoom.Location = New System.Drawing.Point(15, 120)
        Me.chkSpreadAdjacentRoom.Name = "chkSpreadAdjacentRoom"
        Me.chkSpreadAdjacentRoom.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkSpreadAdjacentRoom.Size = New System.Drawing.Size(273, 17)
        Me.chkSpreadAdjacentRoom.TabIndex = 3
        Me.chkSpreadAdjacentRoom.Text = "Model Flame Spread in Adjacent Rooms"
        Me.chkSpreadAdjacentRoom.UseVisualStyleBackColor = False
        Me.chkSpreadAdjacentRoom.Visible = False
        '
        'optUseOneCurve
        '
        Me.optUseOneCurve.BackColor = System.Drawing.SystemColors.Control
        Me.optUseOneCurve.Cursor = System.Windows.Forms.Cursors.Default
        Me.optUseOneCurve.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optUseOneCurve.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optUseOneCurve.Location = New System.Drawing.Point(16, 16)
        Me.optUseOneCurve.Name = "optUseOneCurve"
        Me.optUseOneCurve.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optUseOneCurve.Size = New System.Drawing.Size(293, 37)
        Me.optUseOneCurve.TabIndex = 2
        Me.optUseOneCurve.TabStop = True
        Me.optUseOneCurve.Text = "use one cone calorimeter HRR curve and scale data"
        Me.optUseOneCurve.UseVisualStyleBackColor = False
        '
        'optUseAllCurves
        '
        Me.optUseAllCurves.BackColor = System.Drawing.SystemColors.Control
        Me.optUseAllCurves.Checked = True
        Me.optUseAllCurves.Cursor = System.Windows.Forms.Cursors.Default
        Me.optUseAllCurves.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optUseAllCurves.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optUseAllCurves.Location = New System.Drawing.Point(16, 56)
        Me.optUseAllCurves.Name = "optUseAllCurves"
        Me.optUseAllCurves.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optUseAllCurves.Size = New System.Drawing.Size(293, 25)
        Me.optUseAllCurves.TabIndex = 1
        Me.optUseAllCurves.TabStop = True
        Me.optUseAllCurves.Text = "use all cone calorimeter data provided and interpolate"
        Me.optUseAllCurves.UseVisualStyleBackColor = False
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.Info
        Me.GroupBox1.Controls.Add(Me.chkPessimiseWall)
        Me.GroupBox1.Controls.Add(Me.LblVFS)
        Me.GroupBox1.Controls.Add(Me.LblHFS)
        Me.GroupBox1.Controls.Add(Me.txtVFSlimit)
        Me.GroupBox1.Controls.Add(Me.txtHFSlimit)
        Me.GroupBox1.Controls.Add(Me.lblCeilingpercent)
        Me.GroupBox1.Controls.Add(Me.NumericCeilingPercent)
        Me.GroupBox1.Controls.Add(Me.lblwallpercent)
        Me.GroupBox1.Controls.Add(Me.NumericWallPercent)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 166)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(317, 154)
        Me.GroupBox1.TabIndex = 10
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Partial Coverage "
        '
        'chkPessimiseWall
        '
        Me.chkPessimiseWall.AutoSize = True
        Me.chkPessimiseWall.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkPessimiseWall.Location = New System.Drawing.Point(38, 26)
        Me.chkPessimiseWall.Name = "chkPessimiseWall"
        Me.chkPessimiseWall.Size = New System.Drawing.Size(227, 18)
        Me.chkPessimiseWall.TabIndex = 8
        Me.chkPessimiseWall.Text = "Pessimise combustible wall lining location "
        Me.chkPessimiseWall.UseVisualStyleBackColor = True
        '
        'LblVFS
        '
        Me.LblVFS.AutoSize = True
        Me.LblVFS.Location = New System.Drawing.Point(37, 128)
        Me.LblVFS.Name = "LblVFS"
        Me.LblVFS.Size = New System.Drawing.Size(172, 14)
        Me.LblVFS.TabIndex = 7
        Me.LblVFS.Text = "Vertical combustible limit - wall (m)"
        '
        'LblHFS
        '
        Me.LblHFS.AutoSize = True
        Me.LblHFS.Location = New System.Drawing.Point(25, 102)
        Me.LblHFS.Name = "LblHFS"
        Me.LblHFS.Size = New System.Drawing.Size(184, 14)
        Me.LblHFS.TabIndex = 6
        Me.LblHFS.Text = "Horizontal combustible limit - wall (m)"
        '
        'lblCeilingpercent
        '
        Me.lblCeilingpercent.AutoSize = True
        Me.lblCeilingpercent.Location = New System.Drawing.Point(61, 76)
        Me.lblCeilingpercent.Name = "lblCeilingpercent"
        Me.lblCeilingpercent.Size = New System.Drawing.Size(126, 14)
        Me.lblCeilingpercent.TabIndex = 3
        Me.lblCeilingpercent.Text = "Ceiling - combustible (%)"
        '
        'NumericCeilingPercent
        '
        Me.NumericCeilingPercent.Location = New System.Drawing.Point(215, 76)
        Me.NumericCeilingPercent.Name = "NumericCeilingPercent"
        Me.NumericCeilingPercent.Size = New System.Drawing.Size(50, 20)
        Me.NumericCeilingPercent.TabIndex = 2
        Me.NumericCeilingPercent.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'lblwallpercent
        '
        Me.lblwallpercent.AutoSize = True
        Me.lblwallpercent.Location = New System.Drawing.Point(72, 50)
        Me.lblwallpercent.Name = "lblwallpercent"
        Me.lblwallpercent.Size = New System.Drawing.Size(115, 14)
        Me.lblwallpercent.TabIndex = 1
        Me.lblwallpercent.Text = "Wall - combustible (%)"
        '
        'NumericWallPercent
        '
        Me.NumericWallPercent.Location = New System.Drawing.Point(215, 50)
        Me.NumericWallPercent.Name = "NumericWallPercent"
        Me.NumericWallPercent.Size = New System.Drawing.Size(50, 20)
        Me.NumericWallPercent.TabIndex = 0
        Me.NumericWallPercent.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'frmQuintiere
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(341, 435)
        Me.ControlBox = False
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmQuintiere"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.Text = "Flame Spread Options"
        Me.TopMost = True
        Me.Frame3.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.NumericCeilingPercent, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericWallPercent, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents NumericWallPercent As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblwallpercent As System.Windows.Forms.Label
    Friend WithEvents lblCeilingpercent As System.Windows.Forms.Label
    Friend WithEvents NumericCeilingPercent As System.Windows.Forms.NumericUpDown
    Friend WithEvents LblHFS As System.Windows.Forms.Label
    Friend WithEvents txtVFSlimit As System.Windows.Forms.TextBox
    Friend WithEvents txtHFSlimit As System.Windows.Forms.TextBox
    Friend WithEvents LblVFS As System.Windows.Forms.Label
    Friend WithEvents chkPessimiseWall As System.Windows.Forms.CheckBox
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
#End Region 
End Class