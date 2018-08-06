<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPopulate
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.cmdPopulate = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.ItemsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.txtNumber_Items = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cmdClose = New System.Windows.Forms.Button
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.txtVentClearance = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.cmdStart = New System.Windows.Forms.Button
        Me.cmdStop = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Label2 = New System.Windows.Forms.Label
        Me.updown_itcounter = New System.Windows.Forms.NumericUpDown
        Me.cmdLayout = New System.Windows.Forms.Button
        Me.txtGridSize = New System.Windows.Forms.TextBox
        Me.lblGridSize = New System.Windows.Forms.Label
        Me.MenuStrip2 = New System.Windows.Forms.MenuStrip
        Me.CopyLayoutToClipboardToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.chkShowSprink = New System.Windows.Forms.CheckBox
        Me.txtFLED = New System.Windows.Forms.TextBox
        Me.lblFLED = New System.Windows.Forms.Label
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.optuserlabel = New System.Windows.Forms.RadioButton
        Me.optidnum = New System.Windows.Forms.RadioButton
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.chkShowSD = New System.Windows.Forms.CheckBox
        Me.Panel1.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.updown_itcounter, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip2.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdPopulate
        '
        Me.cmdPopulate.Location = New System.Drawing.Point(294, 29)
        Me.cmdPopulate.Name = "cmdPopulate"
        Me.cmdPopulate.Size = New System.Drawing.Size(114, 23)
        Me.cmdPopulate.TabIndex = 0
        Me.cmdPopulate.Text = "Populate Room"
        Me.cmdPopulate.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.MenuStrip1)
        Me.Panel1.Location = New System.Drawing.Point(218, 24)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(500, 500)
        Me.Panel1.TabIndex = 1
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.Color.LightBlue
        Me.MenuStrip1.Dock = System.Windows.Forms.DockStyle.None
        Me.MenuStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Visible
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ItemsToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(-170, 208)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(91, 24)
        Me.MenuStrip1.TabIndex = 23
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ItemsToolStripMenuItem
        '
        Me.ItemsToolStripMenuItem.Name = "ItemsToolStripMenuItem"
        Me.ItemsToolStripMenuItem.Size = New System.Drawing.Size(79, 20)
        Me.ItemsToolStripMenuItem.Text = "View  Items"
        '
        'txtNumber_Items
        '
        Me.txtNumber_Items.Enabled = False
        Me.txtNumber_Items.Location = New System.Drawing.Point(126, 96)
        Me.txtNumber_Items.Name = "txtNumber_Items"
        Me.txtNumber_Items.ReadOnly = True
        Me.txtNumber_Items.Size = New System.Drawing.Size(39, 20)
        Me.txtNumber_Items.TabIndex = 2
        Me.txtNumber_Items.Text = "3"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(51, 96)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "No. of Items"
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(414, 29)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(83, 23)
        Me.cmdClose.TabIndex = 6
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.CheckBox1.Location = New System.Drawing.Point(74, 185)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(92, 17)
        Me.CheckBox1.TabIndex = 9
        Me.CheckBox1.Text = "Show Vectors"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'txtVentClearance
        '
        Me.txtVentClearance.Location = New System.Drawing.Point(126, 148)
        Me.txtVentClearance.Name = "txtVentClearance"
        Me.txtVentClearance.Size = New System.Drawing.Size(39, 20)
        Me.txtVentClearance.TabIndex = 10
        Me.txtVentClearance.Text = "1"
        Me.ToolTip1.SetToolTip(Me.txtVentClearance, "The distance in front of the wall vent/openings" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "where room contents will be loca" & _
                "ted no closer.")
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(18, 148)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(97, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Vent Clearance (m)"
        '
        'Timer1
        '
        Me.Timer1.Interval = 500
        '
        'cmdStart
        '
        Me.cmdStart.Location = New System.Drawing.Point(9, 27)
        Me.cmdStart.Name = "cmdStart"
        Me.cmdStart.Size = New System.Drawing.Size(75, 23)
        Me.cmdStart.TabIndex = 12
        Me.cmdStart.Text = "start demo"
        Me.cmdStart.UseVisualStyleBackColor = True
        '
        'cmdStop
        '
        Me.cmdStop.Location = New System.Drawing.Point(90, 27)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(75, 23)
        Me.cmdStop.TabIndex = 13
        Me.cmdStop.Text = "stop demo"
        Me.cmdStop.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.updown_itcounter)
        Me.Panel2.Controls.Add(Me.cmdClose)
        Me.Panel2.Controls.Add(Me.cmdLayout)
        Me.Panel2.Controls.Add(Me.cmdPopulate)
        Me.Panel2.Location = New System.Drawing.Point(218, 535)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(500, 66)
        Me.Panel2.TabIndex = 20
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(309, 5)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(122, 13)
        Me.Label2.TabIndex = 29
        Me.Label2.Text = "Recall layout by iteration"
        '
        'updown_itcounter
        '
        Me.updown_itcounter.Location = New System.Drawing.Point(437, 3)
        Me.updown_itcounter.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.updown_itcounter.Name = "updown_itcounter"
        Me.updown_itcounter.Size = New System.Drawing.Size(63, 20)
        Me.updown_itcounter.TabIndex = 30
        Me.updown_itcounter.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'cmdLayout
        '
        Me.cmdLayout.Location = New System.Drawing.Point(190, 29)
        Me.cmdLayout.Name = "cmdLayout"
        Me.cmdLayout.Size = New System.Drawing.Size(98, 23)
        Me.cmdLayout.TabIndex = 29
        Me.cmdLayout.Text = "Show Layout"
        Me.cmdLayout.UseVisualStyleBackColor = True
        Me.cmdLayout.Visible = False
        '
        'txtGridSize
        '
        Me.txtGridSize.Location = New System.Drawing.Point(126, 122)
        Me.txtGridSize.Name = "txtGridSize"
        Me.txtGridSize.Size = New System.Drawing.Size(40, 20)
        Me.txtGridSize.TabIndex = 21
        Me.txtGridSize.Text = "0.1"
        Me.ToolTip1.SetToolTip(Me.txtGridSize, "The room is overlaid with a grid. Item centre points " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "can only be located at a g" & _
                "rid intersection point.")
        '
        'lblGridSize
        '
        Me.lblGridSize.AutoSize = True
        Me.lblGridSize.Location = New System.Drawing.Point(49, 122)
        Me.lblGridSize.Name = "lblGridSize"
        Me.lblGridSize.Size = New System.Drawing.Size(66, 13)
        Me.lblGridSize.TabIndex = 22
        Me.lblGridSize.Text = "Grid Size (m)"
        '
        'MenuStrip2
        '
        Me.MenuStrip2.BackColor = System.Drawing.SystemColors.MenuBar
        Me.MenuStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CopyLayoutToClipboardToolStripMenuItem})
        Me.MenuStrip2.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip2.Name = "MenuStrip2"
        Me.MenuStrip2.Size = New System.Drawing.Size(727, 24)
        Me.MenuStrip2.TabIndex = 23
        Me.MenuStrip2.Text = "MenuStrip2"
        '
        'CopyLayoutToClipboardToolStripMenuItem
        '
        Me.CopyLayoutToClipboardToolStripMenuItem.Name = "CopyLayoutToClipboardToolStripMenuItem"
        Me.CopyLayoutToClipboardToolStripMenuItem.Size = New System.Drawing.Size(155, 20)
        Me.CopyLayoutToClipboardToolStripMenuItem.Text = "Copy Layout to Clipboard"
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.CheckBox2.Location = New System.Drawing.Point(91, 217)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(75, 17)
        Me.CheckBox2.TabIndex = 25
        Me.CheckBox2.Text = "Show Grid"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'chkShowSprink
        '
        Me.chkShowSprink.AutoSize = True
        Me.chkShowSprink.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkShowSprink.Location = New System.Drawing.Point(64, 250)
        Me.chkShowSprink.Name = "chkShowSprink"
        Me.chkShowSprink.Size = New System.Drawing.Size(102, 17)
        Me.chkShowSprink.TabIndex = 26
        Me.chkShowSprink.Text = "Show Sprinklers"
        Me.chkShowSprink.UseVisualStyleBackColor = True
        '
        'txtFLED
        '
        Me.txtFLED.BackColor = System.Drawing.SystemColors.Control
        Me.txtFLED.Enabled = False
        Me.txtFLED.Location = New System.Drawing.Point(126, 70)
        Me.txtFLED.Name = "txtFLED"
        Me.txtFLED.Size = New System.Drawing.Size(40, 20)
        Me.txtFLED.TabIndex = 27
        '
        'lblFLED
        '
        Me.lblFLED.AutoSize = True
        Me.lblFLED.Location = New System.Drawing.Point(39, 70)
        Me.lblFLED.Name = "lblFLED"
        Me.lblFLED.Size = New System.Drawing.Size(76, 13)
        Me.lblFLED.TabIndex = 28
        Me.lblFLED.Text = "FLED (MJ/m2)"
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(103, 26)
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(102, 22)
        Me.ToolStripMenuItem1.Text = "Copy"
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.RadioButton1.Location = New System.Drawing.Point(46, 29)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(120, 17)
        Me.RadioButton1.TabIndex = 75
        Me.RadioButton1.Text = "Auto Populate Items"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.RadioButton2.Checked = True
        Me.RadioButton2.Location = New System.Drawing.Point(12, 52)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(154, 17)
        Me.RadioButton2.TabIndex = 76
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "Manual Positioning of Items"
        Me.RadioButton2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.optuserlabel)
        Me.GroupBox1.Controls.Add(Me.optidnum)
        Me.GroupBox1.Location = New System.Drawing.Point(21, 335)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(178, 92)
        Me.GroupBox1.TabIndex = 78
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Show Item Labels"
        '
        'optuserlabel
        '
        Me.optuserlabel.AutoSize = True
        Me.optuserlabel.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.optuserlabel.Location = New System.Drawing.Point(60, 54)
        Me.optuserlabel.Name = "optuserlabel"
        Me.optuserlabel.Size = New System.Drawing.Size(106, 17)
        Me.optuserlabel.TabIndex = 77
        Me.optuserlabel.Text = "Show User Label"
        Me.optuserlabel.UseVisualStyleBackColor = True
        '
        'optidnum
        '
        Me.optidnum.AutoSize = True
        Me.optidnum.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.optidnum.Checked = True
        Me.optidnum.Location = New System.Drawing.Point(60, 31)
        Me.optidnum.Name = "optidnum"
        Me.optidnum.Size = New System.Drawing.Size(106, 17)
        Me.optidnum.TabIndex = 76
        Me.optidnum.TabStop = True
        Me.optidnum.Text = "Show ID Number"
        Me.optidnum.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.RadioButton1)
        Me.GroupBox2.Controls.Add(Me.RadioButton2)
        Me.GroupBox2.Location = New System.Drawing.Point(21, 442)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(178, 82)
        Me.GroupBox2.TabIndex = 80
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Population Method"
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'chkShowSD
        '
        Me.chkShowSD.AutoSize = True
        Me.chkShowSD.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkShowSD.Location = New System.Drawing.Point(28, 282)
        Me.chkShowSD.Name = "chkShowSD"
        Me.chkShowSD.Size = New System.Drawing.Size(138, 17)
        Me.chkShowSD.TabIndex = 81
        Me.chkShowSD.Text = "Show Smoke Detectors"
        Me.chkShowSD.UseVisualStyleBackColor = True
        '
        'frmPopulate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.BackColor = System.Drawing.SystemColors.ControlLight
        Me.ClientSize = New System.Drawing.Size(727, 613)
        Me.ContextMenuStrip = Me.ContextMenuStrip1
        Me.Controls.Add(Me.chkShowSD)
        Me.Controls.Add(Me.lblFLED)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.txtFLED)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.chkShowSprink)
        Me.Controls.Add(Me.lblGridSize)
        Me.Controls.Add(Me.txtGridSize)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.cmdStart)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtVentClearance)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtNumber_Items)
        Me.Controls.Add(Me.MenuStrip2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel2)
        Me.Name = "frmPopulate"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Populate Room Items"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.updown_itcounter, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip2.ResumeLayout(False)
        Me.MenuStrip2.PerformLayout()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdPopulate As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents txtNumber_Items As System.Windows.Forms.TextBox
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents txtVentClearance As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents cmdStart As System.Windows.Forms.Button
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents txtGridSize As System.Windows.Forms.TextBox
    Friend WithEvents lblGridSize As System.Windows.Forms.Label
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ItemsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MenuStrip2 As System.Windows.Forms.MenuStrip
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowSprink As System.Windows.Forms.CheckBox
    Friend WithEvents txtFLED As System.Windows.Forms.TextBox
    Friend WithEvents lblFLED As System.Windows.Forms.Label
    Friend WithEvents updown_itcounter As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents CopyLayoutToClipboardToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents optuserlabel As System.Windows.Forms.RadioButton
    Friend WithEvents optidnum As System.Windows.Forms.RadioButton
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents cmdLayout As System.Windows.Forms.Button
    Friend WithEvents chkShowSD As System.Windows.Forms.CheckBox
End Class
