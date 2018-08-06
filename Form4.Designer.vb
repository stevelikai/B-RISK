<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPowerlaw
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtAlphaT = New System.Windows.Forms.TextBox
        Me.txtStoreHeight = New System.Windows.Forms.TextBox
        Me.txtPeakHRR = New System.Windows.Forms.TextBox
        Me.optAlpha2 = New System.Windows.Forms.RadioButton
        Me.optAlpha3 = New System.Windows.Forms.RadioButton
        Me.cmdDist_alphaT = New System.Windows.Forms.Button
        Me.cmdPeakHRR = New System.Windows.Forms.Button
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.cmdClose = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(39, 134)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "alpha (kW/s3/m)"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(32, 213)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(95, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Storage Height (m)"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(42, 175)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(85, 13)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Peak HRR (kW)"
        '
        'TxtAlphaT
        '
        Me.TxtAlphaT.Location = New System.Drawing.Point(133, 131)
        Me.TxtAlphaT.Name = "TxtAlphaT"
        Me.TxtAlphaT.Size = New System.Drawing.Size(59, 20)
        Me.TxtAlphaT.TabIndex = 6
        Me.TxtAlphaT.Text = "0.0469"
        '
        'txtStoreHeight
        '
        Me.txtStoreHeight.Location = New System.Drawing.Point(133, 210)
        Me.txtStoreHeight.Name = "txtStoreHeight"
        Me.txtStoreHeight.Size = New System.Drawing.Size(59, 20)
        Me.txtStoreHeight.TabIndex = 8
        Me.txtStoreHeight.Text = "1"
        '
        'txtPeakHRR
        '
        Me.txtPeakHRR.Location = New System.Drawing.Point(133, 172)
        Me.txtPeakHRR.Name = "txtPeakHRR"
        Me.txtPeakHRR.Size = New System.Drawing.Size(59, 20)
        Me.txtPeakHRR.TabIndex = 9
        Me.txtPeakHRR.Text = "20000"
        '
        'optAlpha2
        '
        Me.optAlpha2.AutoSize = True
        Me.optAlpha2.BackColor = System.Drawing.Color.White
        Me.optAlpha2.Checked = True
        Me.optAlpha2.Image = Global.BRISK.My.Resources.Resources.alpha2
        Me.optAlpha2.Location = New System.Drawing.Point(37, 31)
        Me.optAlpha2.Name = "optAlpha2"
        Me.optAlpha2.Size = New System.Drawing.Size(131, 49)
        Me.optAlpha2.TabIndex = 11
        Me.optAlpha2.TabStop = True
        Me.optAlpha2.UseVisualStyleBackColor = False
        '
        'optAlpha3
        '
        Me.optAlpha3.AutoSize = True
        Me.optAlpha3.BackColor = System.Drawing.Color.White
        Me.optAlpha3.Image = Global.BRISK.My.Resources.Resources.alpha3
        Me.optAlpha3.Location = New System.Drawing.Point(201, 29)
        Me.optAlpha3.Name = "optAlpha3"
        Me.optAlpha3.Size = New System.Drawing.Size(153, 51)
        Me.optAlpha3.TabIndex = 10
        Me.optAlpha3.UseVisualStyleBackColor = False
        '
        'cmdDist_alphaT
        '
        Me.cmdDist_alphaT.Location = New System.Drawing.Point(211, 131)
        Me.cmdDist_alphaT.Name = "cmdDist_alphaT"
        Me.cmdDist_alphaT.Size = New System.Drawing.Size(68, 19)
        Me.cmdDist_alphaT.TabIndex = 152
        Me.cmdDist_alphaT.Text = "distribution"
        Me.cmdDist_alphaT.UseVisualStyleBackColor = True
        '
        'cmdPeakHRR
        '
        Me.cmdPeakHRR.Location = New System.Drawing.Point(211, 172)
        Me.cmdPeakHRR.Name = "cmdPeakHRR"
        Me.cmdPeakHRR.Size = New System.Drawing.Size(68, 19)
        Me.cmdPeakHRR.TabIndex = 153
        Me.cmdPeakHRR.Text = "distribution"
        Me.cmdPeakHRR.UseVisualStyleBackColor = True
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(337, 204)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 154
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox1.Controls.Add(Me.optAlpha2)
        Me.GroupBox1.Controls.Add(Me.optAlpha3)
        Me.GroupBox1.Location = New System.Drawing.Point(24, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(390, 97)
        Me.GroupBox1.TabIndex = 156
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Select form of power law design fire"
        '
        'frmPowerlaw
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(425, 243)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.cmdPeakHRR)
        Me.Controls.Add(Me.cmdDist_alphaT)
        Me.Controls.Add(Me.txtPeakHRR)
        Me.Controls.Add(Me.txtStoreHeight)
        Me.Controls.Add(Me.TxtAlphaT)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmPowerlaw"
        Me.Text = "Power Law Design Fire"
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtAlphaT As System.Windows.Forms.TextBox
    Friend WithEvents txtStoreHeight As System.Windows.Forms.TextBox
    Friend WithEvents txtPeakHRR As System.Windows.Forms.TextBox
    Friend WithEvents optAlpha3 As System.Windows.Forms.RadioButton
    Friend WithEvents optAlpha2 As System.Windows.Forms.RadioButton
    Friend WithEvents cmdDist_alphaT As System.Windows.Forms.Button
    Friend WithEvents cmdPeakHRR As System.Windows.Forms.Button
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
End Class
