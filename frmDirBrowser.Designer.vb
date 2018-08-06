<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmDirBrowser
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
	Public WithEvents cmdQuit As System.Windows.Forms.Button
	Public WithEvents cmdBrowse As System.Windows.Forms.Button
	Public WithEvents _Option1_3 As System.Windows.Forms.RadioButton
	Public WithEvents _Option1_2 As System.Windows.Forms.RadioButton
	Public WithEvents _Option1_1 As System.Windows.Forms.RadioButton
	Public WithEvents _Option1_0 As System.Windows.Forms.RadioButton
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents Option1 As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmDirBrowser))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdQuit = New System.Windows.Forms.Button
		Me.cmdBrowse = New System.Windows.Forms.Button
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me._Option1_3 = New System.Windows.Forms.RadioButton
		Me._Option1_2 = New System.Windows.Forms.RadioButton
		Me._Option1_1 = New System.Windows.Forms.RadioButton
		Me._Option1_0 = New System.Windows.Forms.RadioButton
		Me.Option1 = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Option1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "Directory Browser"
		Me.ClientSize = New System.Drawing.Size(216, 139)
		Me.Location = New System.Drawing.Point(404, 122)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmDirBrowser"
		Me.cmdQuit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdQuit.Text = "Quit"
		Me.cmdQuit.Size = New System.Drawing.Size(57, 25)
		Me.cmdQuit.Location = New System.Drawing.Point(144, 104)
		Me.cmdQuit.TabIndex = 6
		Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdQuit.CausesValidation = True
		Me.cmdQuit.Enabled = True
		Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdQuit.TabStop = True
		Me.cmdQuit.Name = "cmdQuit"
		Me.cmdBrowse.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdBrowse.Text = "Browse"
		Me.cmdBrowse.Size = New System.Drawing.Size(57, 25)
		Me.cmdBrowse.Location = New System.Drawing.Point(144, 72)
		Me.cmdBrowse.TabIndex = 5
		Me.cmdBrowse.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdBrowse.BackColor = System.Drawing.SystemColors.Control
		Me.cmdBrowse.CausesValidation = True
		Me.cmdBrowse.Enabled = True
		Me.cmdBrowse.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdBrowse.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdBrowse.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdBrowse.TabStop = True
		Me.cmdBrowse.Name = "cmdBrowse"
		Me.Frame1.Text = "Browse For:"
		Me.Frame1.Size = New System.Drawing.Size(129, 121)
		Me.Frame1.Location = New System.Drawing.Point(8, 8)
		Me.Frame1.TabIndex = 0
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me._Option1_3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Option1_3.Text = "Printers"
		Me._Option1_3.Size = New System.Drawing.Size(105, 21)
		Me._Option1_3.Location = New System.Drawing.Point(16, 96)
		Me._Option1_3.TabIndex = 4
		Me._Option1_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Option1_3.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Option1_3.BackColor = System.Drawing.SystemColors.Control
		Me._Option1_3.CausesValidation = True
		Me._Option1_3.Enabled = True
		Me._Option1_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Option1_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._Option1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Option1_3.Appearance = System.Windows.Forms.Appearance.Normal
		Me._Option1_3.TabStop = True
		Me._Option1_3.Checked = False
		Me._Option1_3.Visible = True
		Me._Option1_3.Name = "_Option1_3"
		Me._Option1_2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Option1_2.Text = "Computers"
		Me._Option1_2.Size = New System.Drawing.Size(105, 21)
		Me._Option1_2.Location = New System.Drawing.Point(16, 72)
		Me._Option1_2.TabIndex = 3
		Me._Option1_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Option1_2.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Option1_2.BackColor = System.Drawing.SystemColors.Control
		Me._Option1_2.CausesValidation = True
		Me._Option1_2.Enabled = True
		Me._Option1_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Option1_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Option1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Option1_2.Appearance = System.Windows.Forms.Appearance.Normal
		Me._Option1_2.TabStop = True
		Me._Option1_2.Checked = False
		Me._Option1_2.Visible = True
		Me._Option1_2.Name = "_Option1_2"
		Me._Option1_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Option1_1.Text = "Folders and Files"
		Me._Option1_1.Size = New System.Drawing.Size(105, 21)
		Me._Option1_1.Location = New System.Drawing.Point(16, 48)
		Me._Option1_1.TabIndex = 2
		Me._Option1_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Option1_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Option1_1.BackColor = System.Drawing.SystemColors.Control
		Me._Option1_1.CausesValidation = True
		Me._Option1_1.Enabled = True
		Me._Option1_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Option1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Option1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Option1_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._Option1_1.TabStop = True
		Me._Option1_1.Checked = False
		Me._Option1_1.Visible = True
		Me._Option1_1.Name = "_Option1_1"
		Me._Option1_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Option1_0.Text = "Folders"
		Me._Option1_0.Size = New System.Drawing.Size(105, 21)
		Me._Option1_0.Location = New System.Drawing.Point(16, 24)
		Me._Option1_0.TabIndex = 1
		Me._Option1_0.Checked = True
		Me._Option1_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Option1_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Option1_0.BackColor = System.Drawing.SystemColors.Control
		Me._Option1_0.CausesValidation = True
		Me._Option1_0.Enabled = True
		Me._Option1_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Option1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Option1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Option1_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._Option1_0.TabStop = True
		Me._Option1_0.Visible = True
		Me._Option1_0.Name = "_Option1_0"
		Me.Controls.Add(cmdQuit)
		Me.Controls.Add(cmdBrowse)
		Me.Controls.Add(Frame1)
		Me.Frame1.Controls.Add(_Option1_3)
		Me.Frame1.Controls.Add(_Option1_2)
		Me.Frame1.Controls.Add(_Option1_1)
		Me.Frame1.Controls.Add(_Option1_0)
		Me.Option1.SetIndex(_Option1_3, CType(3, Short))
		Me.Option1.SetIndex(_Option1_2, CType(2, Short))
		Me.Option1.SetIndex(_Option1_1, CType(1, Short))
		Me.Option1.SetIndex(_Option1_0, CType(0, Short))
		CType(Me.Option1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class