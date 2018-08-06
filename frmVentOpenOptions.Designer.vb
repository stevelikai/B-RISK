<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmVentOpenOptions
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cboSDroom = New System.Windows.Forms.ComboBox
        Me.chkHD = New System.Windows.Forms.CheckBox
        Me.chkSD = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtTriggerOpenDelay = New System.Windows.Forms.TextBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtTriggerTimeVentOpen = New System.Windows.Forms.TextBox
        Me.cmdClose = New System.Windows.Forms.Button
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.txtFireSize = New System.Windows.Forms.TextBox
        Me.chkHRR = New System.Windows.Forms.CheckBox
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtHRROpenDelay = New System.Windows.Forms.TextBox
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtHRRTimeVentOpen = New System.Windows.Forms.TextBox
        Me.chkAutoVentOpen = New System.Windows.Forms.CheckBox
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.GroupBox7 = New System.Windows.Forms.GroupBox
        Me.chkFO = New System.Windows.Forms.CheckBox
        Me.GroupBox8 = New System.Windows.Forms.GroupBox
        Me.chkVL = New System.Windows.Forms.CheckBox
        Me.GroupBox9 = New System.Windows.Forms.GroupBox
        Me.txtHODlabel = New System.Windows.Forms.TextBox
        Me.chkHoldOpen = New System.Windows.Forms.CheckBox
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox7.SuspendLayout()
        Me.GroupBox8.SuspendLayout()
        Me.GroupBox9.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cboSDroom)
        Me.GroupBox1.Controls.Add(Me.chkHD)
        Me.GroupBox1.Controls.Add(Me.chkSD)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 45)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(243, 118)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Auto Opening Trigger"
        '
        'cboSDroom
        '
        Me.cboSDroom.FormattingEnabled = True
        Me.cboSDroom.Location = New System.Drawing.Point(168, 41)
        Me.cboSDroom.Name = "cboSDroom"
        Me.cboSDroom.Size = New System.Drawing.Size(57, 21)
        Me.cboSDroom.TabIndex = 9
        '
        'chkHD
        '
        Me.chkHD.AutoSize = True
        Me.chkHD.Location = New System.Drawing.Point(17, 83)
        Me.chkHD.Name = "chkHD"
        Me.chkHD.Size = New System.Drawing.Size(164, 17)
        Me.chkHD.TabIndex = 8
        Me.chkHD.Text = "Thermal Detector or Sprinkler"
        Me.chkHD.UseVisualStyleBackColor = True
        '
        'chkSD
        '
        Me.chkSD.AutoSize = True
        Me.chkSD.Location = New System.Drawing.Point(17, 43)
        Me.chkSD.Name = "chkSD"
        Me.chkSD.Size = New System.Drawing.Size(145, 17)
        Me.chkSD.TabIndex = 7
        Me.chkSD.Text = "Smoke Detector in Room"
        Me.chkSD.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.txtTriggerOpenDelay)
        Me.GroupBox2.Location = New System.Drawing.Point(261, 49)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(167, 114)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Time Vent Opens After Trigger"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(103, 46)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(24, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "sec"
        '
        'txtTriggerOpenDelay
        '
        Me.txtTriggerOpenDelay.Location = New System.Drawing.Point(25, 39)
        Me.txtTriggerOpenDelay.Name = "txtTriggerOpenDelay"
        Me.txtTriggerOpenDelay.Size = New System.Drawing.Size(72, 20)
        Me.txtTriggerOpenDelay.TabIndex = 2
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.txtTriggerTimeVentOpen)
        Me.GroupBox3.Location = New System.Drawing.Point(434, 48)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(147, 115)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Duration Vent is Open"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(98, 47)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(24, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "sec"
        '
        'txtTriggerTimeVentOpen
        '
        Me.txtTriggerTimeVentOpen.Location = New System.Drawing.Point(20, 40)
        Me.txtTriggerTimeVentOpen.Name = "txtTriggerTimeVentOpen"
        Me.txtTriggerTimeVentOpen.Size = New System.Drawing.Size(72, 20)
        Me.txtTriggerTimeVentOpen.TabIndex = 1
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(477, 357)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(104, 31)
        Me.cmdClose.TabIndex = 3
        Me.cmdClose.Text = "Save Settings"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.txtFireSize)
        Me.GroupBox4.Controls.Add(Me.chkHRR)
        Me.GroupBox4.Location = New System.Drawing.Point(10, 169)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(245, 60)
        Me.GroupBox4.TabIndex = 4
        Me.GroupBox4.TabStop = False
        '
        'txtFireSize
        '
        Me.txtFireSize.Location = New System.Drawing.Point(128, 19)
        Me.txtFireSize.Name = "txtFireSize"
        Me.txtFireSize.Size = New System.Drawing.Size(66, 20)
        Me.txtFireSize.TabIndex = 1
        '
        'chkHRR
        '
        Me.chkHRR.AutoSize = True
        Me.chkHRR.Location = New System.Drawing.Point(19, 19)
        Me.chkHRR.Name = "chkHRR"
        Me.chkHRR.Size = New System.Drawing.Size(92, 17)
        Me.chkHRR.TabIndex = 0
        Me.chkHRR.Text = "Fire Size (kW)"
        Me.chkHRR.UseVisualStyleBackColor = True
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.Label1)
        Me.GroupBox5.Controls.Add(Me.txtHRROpenDelay)
        Me.GroupBox5.Location = New System.Drawing.Point(261, 169)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(167, 60)
        Me.GroupBox5.TabIndex = 5
        Me.GroupBox5.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(103, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(24, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "sec"
        '
        'txtHRROpenDelay
        '
        Me.txtHRROpenDelay.Location = New System.Drawing.Point(25, 19)
        Me.txtHRROpenDelay.Name = "txtHRROpenDelay"
        Me.txtHRROpenDelay.Size = New System.Drawing.Size(72, 20)
        Me.txtHRROpenDelay.TabIndex = 3
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.Label4)
        Me.GroupBox6.Controls.Add(Me.txtHRRTimeVentOpen)
        Me.GroupBox6.Location = New System.Drawing.Point(434, 169)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(147, 60)
        Me.GroupBox6.TabIndex = 6
        Me.GroupBox6.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(98, 26)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(24, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "sec"
        '
        'txtHRRTimeVentOpen
        '
        Me.txtHRRTimeVentOpen.Location = New System.Drawing.Point(20, 19)
        Me.txtHRRTimeVentOpen.Name = "txtHRRTimeVentOpen"
        Me.txtHRRTimeVentOpen.Size = New System.Drawing.Size(72, 20)
        Me.txtHRRTimeVentOpen.TabIndex = 4
        '
        'chkAutoVentOpen
        '
        Me.chkAutoVentOpen.AutoSize = True
        Me.chkAutoVentOpen.Location = New System.Drawing.Point(29, 12)
        Me.chkAutoVentOpen.Name = "chkAutoVentOpen"
        Me.chkAutoVentOpen.Size = New System.Drawing.Size(250, 17)
        Me.chkAutoVentOpen.TabIndex = 7
        Me.chkAutoVentOpen.Text = "Enable Automatic Opening Options for this Vent"
        Me.chkAutoVentOpen.UseVisualStyleBackColor = True
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.chkFO)
        Me.GroupBox7.Location = New System.Drawing.Point(12, 235)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(243, 49)
        Me.GroupBox7.TabIndex = 8
        Me.GroupBox7.TabStop = False
        '
        'chkFO
        '
        Me.chkFO.AutoSize = True
        Me.chkFO.Location = New System.Drawing.Point(17, 19)
        Me.chkFO.Name = "chkFO"
        Me.chkFO.Size = New System.Drawing.Size(113, 17)
        Me.chkFO.TabIndex = 9
        Me.chkFO.Text = "Open at Flashover"
        Me.chkFO.UseVisualStyleBackColor = True
        '
        'GroupBox8
        '
        Me.GroupBox8.Controls.Add(Me.chkVL)
        Me.GroupBox8.Location = New System.Drawing.Point(12, 290)
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.Size = New System.Drawing.Size(243, 49)
        Me.GroupBox8.TabIndex = 9
        Me.GroupBox8.TabStop = False
        '
        'chkVL
        '
        Me.chkVL.AutoSize = True
        Me.chkVL.Location = New System.Drawing.Point(17, 19)
        Me.chkVL.Name = "chkVL"
        Me.chkVL.Size = New System.Drawing.Size(194, 17)
        Me.chkVL.TabIndex = 9
        Me.chkVL.Text = "Open when ventilation limit reached"
        Me.chkVL.UseVisualStyleBackColor = True
        '
        'GroupBox9
        '
        Me.GroupBox9.Controls.Add(Me.txtHODlabel)
        Me.GroupBox9.Controls.Add(Me.chkHoldOpen)
        Me.GroupBox9.Location = New System.Drawing.Point(261, 235)
        Me.GroupBox9.Name = "GroupBox9"
        Me.GroupBox9.Size = New System.Drawing.Size(317, 104)
        Me.GroupBox9.TabIndex = 10
        Me.GroupBox9.TabStop = False
        '
        'txtHODlabel
        '
        Me.txtHODlabel.BackColor = System.Drawing.SystemColors.Control
        Me.txtHODlabel.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtHODlabel.Location = New System.Drawing.Point(25, 43)
        Me.txtHODlabel.Multiline = True
        Me.txtHODlabel.Name = "txtHODlabel"
        Me.txtHODlabel.ReadOnly = True
        Me.txtHODlabel.Size = New System.Drawing.Size(253, 35)
        Me.txtHODlabel.TabIndex = 11
        Me.txtHODlabel.Text = "Device closes the vent when any smoke detector responds in room"
        '
        'chkHoldOpen
        '
        Me.chkHoldOpen.AutoSize = True
        Me.chkHoldOpen.Location = New System.Drawing.Point(17, 19)
        Me.chkHoldOpen.Name = "chkHoldOpen"
        Me.chkHoldOpen.Size = New System.Drawing.Size(191, 17)
        Me.chkHoldOpen.TabIndex = 9
        Me.chkHoldOpen.Text = "Hold-open device fitted to this vent"
        Me.chkHoldOpen.UseVisualStyleBackColor = True
        '
        'frmVentOpenOptions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(590, 400)
        Me.Controls.Add(Me.GroupBox9)
        Me.Controls.Add(Me.GroupBox8)
        Me.Controls.Add(Me.GroupBox7)
        Me.Controls.Add(Me.chkAutoVentOpen)
        Me.Controls.Add(Me.GroupBox6)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "frmVentOpenOptions"
        Me.Text = "Vent Opening Options"
        Me.TopMost = True
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox7.PerformLayout()
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox8.PerformLayout()
        Me.GroupBox9.ResumeLayout(False)
        Me.GroupBox9.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents cboSDroom As System.Windows.Forms.ComboBox
    Friend WithEvents chkHD As System.Windows.Forms.CheckBox
    Friend WithEvents chkSD As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents txtFireSize As System.Windows.Forms.TextBox
    Friend WithEvents chkHRR As System.Windows.Forms.CheckBox
    Friend WithEvents chkAutoVentOpen As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtTriggerOpenDelay As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtTriggerTimeVentOpen As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtHRROpenDelay As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtHRRTimeVentOpen As System.Windows.Forms.TextBox
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents chkFO As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox8 As System.Windows.Forms.GroupBox
    Friend WithEvents chkVL As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox9 As System.Windows.Forms.GroupBox
    Friend WithEvents chkHoldOpen As System.Windows.Forms.CheckBox
    Friend WithEvents txtHODlabel As System.Windows.Forms.TextBox
End Class
