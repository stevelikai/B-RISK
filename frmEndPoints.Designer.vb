<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmEndPoints
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
    Public WithEvents lblFED As System.Windows.Forms.Label
    Public WithEvents lblVisibility As System.Windows.Forms.Label
    Public WithEvents lblTemp As System.Windows.Forms.Label
    Public WithEvents lblSprinkler As System.Windows.Forms.Label
    Public WithEvents lblConvect As System.Windows.Forms.Label
    Public WithEvents lblTarget As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdClose = New System.Windows.Forms.Button
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.lblFED = New System.Windows.Forms.Label
        Me.lblVisibility = New System.Windows.Forms.Label
        Me.lblTemp = New System.Windows.Forms.Label
        Me.lblSprinkler = New System.Windows.Forms.Label
        Me.lblConvect = New System.Windows.Forms.Label
        Me.lblTarget = New System.Windows.Forms.Label
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Location = New System.Drawing.Point(328, 264)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(57, 25)
        Me.cmdClose.TabIndex = 1
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.lblFED)
        Me.Frame1.Controls.Add(Me.lblVisibility)
        Me.Frame1.Controls.Add(Me.lblTemp)
        Me.Frame1.Controls.Add(Me.lblSprinkler)
        Me.Frame1.Controls.Add(Me.lblConvect)
        Me.Frame1.Controls.Add(Me.lblTarget)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(8, 0)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(377, 257)
        Me.Frame1.TabIndex = 0
        Me.Frame1.TabStop = False
        '
        'lblFED
        '
        Me.lblFED.BackColor = System.Drawing.Color.Transparent
        Me.lblFED.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFED.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFED.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblFED.Location = New System.Drawing.Point(16, 144)
        Me.lblFED.Name = "lblFED"
        Me.lblFED.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFED.Size = New System.Drawing.Size(345, 25)
        Me.lblFED.TabIndex = 7
        Me.lblFED.Text = "fed"
        '
        'lblVisibility
        '
        Me.lblVisibility.BackColor = System.Drawing.Color.Transparent
        Me.lblVisibility.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVisibility.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVisibility.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblVisibility.Location = New System.Drawing.Point(16, 104)
        Me.lblVisibility.Name = "lblVisibility"
        Me.lblVisibility.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVisibility.Size = New System.Drawing.Size(345, 25)
        Me.lblVisibility.TabIndex = 6
        Me.lblVisibility.Text = "visibility"
        '
        'lblTemp
        '
        Me.lblTemp.BackColor = System.Drawing.Color.Transparent
        Me.lblTemp.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTemp.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTemp.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblTemp.Location = New System.Drawing.Point(16, 64)
        Me.lblTemp.Name = "lblTemp"
        Me.lblTemp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTemp.Size = New System.Drawing.Size(345, 25)
        Me.lblTemp.TabIndex = 5
        Me.lblTemp.Text = "temp"
        '
        'lblSprinkler
        '
        Me.lblSprinkler.BackColor = System.Drawing.Color.Transparent
        Me.lblSprinkler.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSprinkler.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSprinkler.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblSprinkler.Location = New System.Drawing.Point(16, 224)
        Me.lblSprinkler.Name = "lblSprinkler"
        Me.lblSprinkler.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSprinkler.Size = New System.Drawing.Size(345, 25)
        Me.lblSprinkler.TabIndex = 4
        Me.lblSprinkler.Text = "sprinkler"
        '
        'lblConvect
        '
        Me.lblConvect.BackColor = System.Drawing.Color.Transparent
        Me.lblConvect.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblConvect.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblConvect.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblConvect.Location = New System.Drawing.Point(16, 184)
        Me.lblConvect.Name = "lblConvect"
        Me.lblConvect.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblConvect.Size = New System.Drawing.Size(345, 25)
        Me.lblConvect.TabIndex = 3
        Me.lblConvect.Text = "convective heat"
        '
        'lblTarget
        '
        Me.lblTarget.BackColor = System.Drawing.Color.Transparent
        Me.lblTarget.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTarget.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTarget.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblTarget.Location = New System.Drawing.Point(16, 24)
        Me.lblTarget.Name = "lblTarget"
        Me.lblTarget.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTarget.Size = New System.Drawing.Size(345, 25)
        Me.lblTarget.TabIndex = 2
        Me.lblTarget.Text = "target"
        '
        'frmEndPoints
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(393, 296)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 23)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmEndPoints"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.Text = "Summary of End-Points for Room of Fire Origin"
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class