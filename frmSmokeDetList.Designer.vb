<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSmokeDetList
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
        Me.lstSmokeDet = New System.Windows.Forms.ListBox
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.cmdDist_sprreliability = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtSmokeDetReliability = New System.Windows.Forms.TextBox
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.cmdExit = New System.Windows.Forms.Button
        Me.cmdRemoveSmokeDet = New System.Windows.Forms.Button
        Me.cmdAddSmokeDet = New System.Windows.Forms.Button
        Me.cmdCopySmokeDet = New System.Windows.Forms.Button
        Me.cmdEditSmokeDet = New System.Windows.Forms.Button
        Me.chkCalcSRadialDist = New System.Windows.Forms.CheckBox
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.GroupBox2.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lstSmokeDet
        '
        Me.lstSmokeDet.FormattingEnabled = True
        Me.lstSmokeDet.Location = New System.Drawing.Point(17, 44)
        Me.lstSmokeDet.Name = "lstSmokeDet"
        Me.lstSmokeDet.Size = New System.Drawing.Size(353, 160)
        Me.lstSmokeDet.TabIndex = 125
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cmdDist_sprreliability)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.txtSmokeDetReliability)
        Me.GroupBox2.Location = New System.Drawing.Point(15, 235)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(431, 79)
        Me.GroupBox2.TabIndex = 132
        Me.GroupBox2.TabStop = False
        '
        'cmdDist_sprreliability
        '
        Me.cmdDist_sprreliability.Location = New System.Drawing.Point(263, 38)
        Me.cmdDist_sprreliability.Name = "cmdDist_sprreliability"
        Me.cmdDist_sprreliability.Size = New System.Drawing.Size(68, 19)
        Me.cmdDist_sprreliability.TabIndex = 152
        Me.cmdDist_sprreliability.Text = "distribution"
        Me.cmdDist_sprreliability.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 41)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(168, 13)
        Me.Label1.TabIndex = 120
        Me.Label1.Text = "Smoke Detector System Reliability"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtSmokeDetReliability
        '
        Me.txtSmokeDetReliability.BackColor = System.Drawing.SystemColors.Window
        Me.txtSmokeDetReliability.Location = New System.Drawing.Point(180, 38)
        Me.txtSmokeDetReliability.Name = "txtSmokeDetReliability"
        Me.txtSmokeDetReliability.Size = New System.Drawing.Size(64, 20)
        Me.txtSmokeDetReliability.TabIndex = 119
        Me.txtSmokeDetReliability.Tag = "Enter probability that the smoke detection system will activate"
        Me.txtSmokeDetReliability.Text = "1.00"
        Me.txtSmokeDetReliability.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'ListBox1
        '
        Me.ListBox1.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(17, 21)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(352, 17)
        Me.ListBox1.TabIndex = 129
        '
        'cmdExit
        '
        Me.cmdExit.Location = New System.Drawing.Point(469, 291)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(75, 23)
        Me.cmdExit.TabIndex = 128
        Me.cmdExit.Text = "Close"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'cmdRemoveSmokeDet
        '
        Me.cmdRemoveSmokeDet.Location = New System.Drawing.Point(375, 172)
        Me.cmdRemoveSmokeDet.Name = "cmdRemoveSmokeDet"
        Me.cmdRemoveSmokeDet.Size = New System.Drawing.Size(75, 23)
        Me.cmdRemoveSmokeDet.TabIndex = 127
        Me.cmdRemoveSmokeDet.Text = "Remove"
        Me.cmdRemoveSmokeDet.UseVisualStyleBackColor = True
        '
        'cmdAddSmokeDet
        '
        Me.cmdAddSmokeDet.Location = New System.Drawing.Point(374, 44)
        Me.cmdAddSmokeDet.Name = "cmdAddSmokeDet"
        Me.cmdAddSmokeDet.Size = New System.Drawing.Size(170, 23)
        Me.cmdAddSmokeDet.TabIndex = 126
        Me.cmdAddSmokeDet.Text = "Add Smoke Detector"
        Me.cmdAddSmokeDet.UseVisualStyleBackColor = True
        '
        'cmdCopySmokeDet
        '
        Me.cmdCopySmokeDet.Location = New System.Drawing.Point(376, 143)
        Me.cmdCopySmokeDet.Name = "cmdCopySmokeDet"
        Me.cmdCopySmokeDet.Size = New System.Drawing.Size(75, 23)
        Me.cmdCopySmokeDet.TabIndex = 131
        Me.cmdCopySmokeDet.Text = "Copy"
        Me.cmdCopySmokeDet.UseVisualStyleBackColor = True
        '
        'cmdEditSmokeDet
        '
        Me.cmdEditSmokeDet.Location = New System.Drawing.Point(376, 113)
        Me.cmdEditSmokeDet.Name = "cmdEditSmokeDet"
        Me.cmdEditSmokeDet.Size = New System.Drawing.Size(74, 24)
        Me.cmdEditSmokeDet.TabIndex = 130
        Me.cmdEditSmokeDet.Text = "Edit"
        Me.cmdEditSmokeDet.UseVisualStyleBackColor = True
        '
        'chkCalcSRadialDist
        '
        Me.chkCalcSRadialDist.AutoSize = True
        Me.chkCalcSRadialDist.Location = New System.Drawing.Point(18, 212)
        Me.chkCalcSRadialDist.Name = "chkCalcSRadialDist"
        Me.chkCalcSRadialDist.Size = New System.Drawing.Size(332, 17)
        Me.chkCalcSRadialDist.TabIndex = 136
        Me.chkCalcSRadialDist.Text = "Calculate fire to sensor radial distance (overrides detector setting)"
        Me.chkCalcSRadialDist.UseVisualStyleBackColor = True
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'frmSmokeDetList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(554, 325)
        Me.Controls.Add(Me.lstSmokeDet)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdRemoveSmokeDet)
        Me.Controls.Add(Me.cmdAddSmokeDet)
        Me.Controls.Add(Me.cmdCopySmokeDet)
        Me.Controls.Add(Me.cmdEditSmokeDet)
        Me.Controls.Add(Me.chkCalcSRadialDist)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MinimizeBox = False
        Me.Name = "frmSmokeDetList"
        Me.Text = "Smoke Detectors"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lstSmokeDet As System.Windows.Forms.ListBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdDist_sprreliability As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtSmokeDetReliability As System.Windows.Forms.TextBox
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents cmdRemoveSmokeDet As System.Windows.Forms.Button
    Friend WithEvents cmdAddSmokeDet As System.Windows.Forms.Button
    Friend WithEvents cmdCopySmokeDet As System.Windows.Forms.Button
    Friend WithEvents cmdEditSmokeDet As System.Windows.Forms.Button
    Friend WithEvents chkCalcSRadialDist As System.Windows.Forms.CheckBox
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
End Class
