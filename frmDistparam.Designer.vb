<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDistParam
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtValue = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.cmd_SaveDist = New System.Windows.Forms.Button
        Me.txtMean = New System.Windows.Forms.TextBox
        Me.txtMode = New System.Windows.Forms.TextBox
        Me.txtVariance = New System.Windows.Forms.TextBox
        Me.txtUpperBound = New System.Windows.Forms.TextBox
        Me.txtLowerBound = New System.Windows.Forms.TextBox
        Me.txtAlpha = New System.Windows.Forms.TextBox
        Me.txtBeta = New System.Windows.Forms.TextBox
        Me.lblDistName = New System.Windows.Forms.Label
        Me.lstDistribution = New System.Windows.Forms.ComboBox
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.LblInstruction = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(121, 107)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(34, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Value"
        '
        'txtValue
        '
        Me.txtValue.Location = New System.Drawing.Point(161, 107)
        Me.txtValue.Name = "txtValue"
        Me.txtValue.Size = New System.Drawing.Size(75, 20)
        Me.txtValue.TabIndex = 1
        Me.txtValue.Text = "0"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(121, 143)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(34, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Mean"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(121, 172)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(34, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Mode"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(104, 198)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(49, 13)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Variance"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(83, 224)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(70, 13)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "Upper Bound"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(83, 251)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(70, 13)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "Lower Bound"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(122, 277)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(34, 13)
        Me.Label8.TabIndex = 8
        Me.Label8.Text = "Alpha"
        Me.Label8.Visible = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(127, 305)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(29, 13)
        Me.Label9.TabIndex = 9
        Me.Label9.Text = "Beta"
        Me.Label9.Visible = False
        '
        'cmd_SaveDist
        '
        Me.cmd_SaveDist.Location = New System.Drawing.Point(161, 342)
        Me.cmd_SaveDist.Name = "cmd_SaveDist"
        Me.cmd_SaveDist.Size = New System.Drawing.Size(75, 23)
        Me.cmd_SaveDist.TabIndex = 10
        Me.cmd_SaveDist.Text = "Save"
        Me.cmd_SaveDist.UseVisualStyleBackColor = True
        '
        'txtMean
        '
        Me.txtMean.Location = New System.Drawing.Point(161, 136)
        Me.txtMean.Name = "txtMean"
        Me.txtMean.Size = New System.Drawing.Size(75, 20)
        Me.txtMean.TabIndex = 11
        Me.txtMean.Text = "0"
        '
        'txtMode
        '
        Me.txtMode.Location = New System.Drawing.Point(161, 165)
        Me.txtMode.Name = "txtMode"
        Me.txtMode.Size = New System.Drawing.Size(75, 20)
        Me.txtMode.TabIndex = 12
        Me.txtMode.Text = "0"
        '
        'txtVariance
        '
        Me.txtVariance.Location = New System.Drawing.Point(161, 195)
        Me.txtVariance.Name = "txtVariance"
        Me.txtVariance.Size = New System.Drawing.Size(75, 20)
        Me.txtVariance.TabIndex = 13
        Me.txtVariance.Text = "0"
        '
        'txtUpperBound
        '
        Me.txtUpperBound.Location = New System.Drawing.Point(161, 224)
        Me.txtUpperBound.Name = "txtUpperBound"
        Me.txtUpperBound.Size = New System.Drawing.Size(75, 20)
        Me.txtUpperBound.TabIndex = 14
        Me.txtUpperBound.Text = "0"
        '
        'txtLowerBound
        '
        Me.txtLowerBound.Location = New System.Drawing.Point(161, 251)
        Me.txtLowerBound.Name = "txtLowerBound"
        Me.txtLowerBound.Size = New System.Drawing.Size(75, 20)
        Me.txtLowerBound.TabIndex = 15
        Me.txtLowerBound.Text = "0"
        '
        'txtAlpha
        '
        Me.txtAlpha.Location = New System.Drawing.Point(161, 277)
        Me.txtAlpha.Name = "txtAlpha"
        Me.txtAlpha.Size = New System.Drawing.Size(75, 20)
        Me.txtAlpha.TabIndex = 16
        Me.txtAlpha.Text = "0"
        Me.txtAlpha.Visible = False
        '
        'txtBeta
        '
        Me.txtBeta.Location = New System.Drawing.Point(161, 303)
        Me.txtBeta.Name = "txtBeta"
        Me.txtBeta.Size = New System.Drawing.Size(75, 20)
        Me.txtBeta.TabIndex = 17
        Me.txtBeta.Text = "0"
        Me.txtBeta.Visible = False
        '
        'lblDistName
        '
        Me.lblDistName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDistName.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblDistName.Location = New System.Drawing.Point(12, 9)
        Me.lblDistName.Name = "lblDistName"
        Me.lblDistName.Size = New System.Drawing.Size(224, 23)
        Me.lblDistName.TabIndex = 18
        Me.lblDistName.Text = "Label4"
        Me.lblDistName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lstDistribution
        '
        Me.lstDistribution.FormattingEnabled = True
        Me.lstDistribution.Items.AddRange(New Object() {"None", "Normal", "Uniform", "Triangular", "Log Normal"})
        Me.lstDistribution.Location = New System.Drawing.Point(107, 71)
        Me.lstDistribution.Name = "lstDistribution"
        Me.lstDistribution.Size = New System.Drawing.Size(129, 21)
        Me.lstDistribution.TabIndex = 19
        Me.lstDistribution.Text = "None"
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(161, 371)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 20
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'LblInstruction
        '
        Me.LblInstruction.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblInstruction.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.LblInstruction.Location = New System.Drawing.Point(12, 32)
        Me.LblInstruction.Name = "LblInstruction"
        Me.LblInstruction.Size = New System.Drawing.Size(224, 36)
        Me.LblInstruction.TabIndex = 21
        Me.LblInstruction.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'frmDistParam
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(260, 405)
        Me.Controls.Add(Me.LblInstruction)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.lstDistribution)
        Me.Controls.Add(Me.lblDistName)
        Me.Controls.Add(Me.txtBeta)
        Me.Controls.Add(Me.txtAlpha)
        Me.Controls.Add(Me.txtLowerBound)
        Me.Controls.Add(Me.txtUpperBound)
        Me.Controls.Add(Me.txtVariance)
        Me.Controls.Add(Me.txtMode)
        Me.Controls.Add(Me.txtMean)
        Me.Controls.Add(Me.cmd_SaveDist)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtValue)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmDistParam"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Distribution Parameters"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtValue As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cmd_SaveDist As System.Windows.Forms.Button
    Friend WithEvents txtMean As System.Windows.Forms.TextBox
    Friend WithEvents txtMode As System.Windows.Forms.TextBox
    Friend WithEvents txtVariance As System.Windows.Forms.TextBox
    Friend WithEvents txtUpperBound As System.Windows.Forms.TextBox
    Friend WithEvents txtLowerBound As System.Windows.Forms.TextBox
    Friend WithEvents txtAlpha As System.Windows.Forms.TextBox
    Friend WithEvents txtBeta As System.Windows.Forms.TextBox
    Friend WithEvents lblDistName As System.Windows.Forms.Label
    Friend WithEvents lstDistribution As System.Windows.Forms.ComboBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents LblInstruction As System.Windows.Forms.Label
End Class
