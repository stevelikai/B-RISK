<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSensitivity
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
        Me.LB_input = New System.Windows.Forms.ListBox()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.LB_input_ID = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.LB_output = New System.Windows.Forms.ListBox()
        Me.cmdPlotSENS = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'LB_input
        '
        Me.LB_input.FormattingEnabled = True
        Me.LB_input.Items.AddRange(New Object() {"Wall vent width", "Wall vent height", "Room length", "Room width", "Heat of combustion (by item)", "Soot yield (by item)", "Radiant loss fraction (by item)", "HRRPUA (by item)", "Latent heat of gasification (by item)", "Ceiling vent area"})
        Me.LB_input.Location = New System.Drawing.Point(12, 50)
        Me.LB_input.Name = "LB_input"
        Me.LB_input.Size = New System.Drawing.Size(200, 95)
        Me.LB_input.TabIndex = 0
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(240, 287)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 1
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'LB_input_ID
        '
        Me.LB_input_ID.FormattingEnabled = True
        Me.LB_input_ID.Location = New System.Drawing.Point(218, 50)
        Me.LB_input_ID.Name = "LB_input_ID"
        Me.LB_input_ID.Size = New System.Drawing.Size(97, 95)
        Me.LB_input_ID.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 34)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(81, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Input parameter"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(215, 34)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(43, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Vent ID"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 161)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(89, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Output parameter"
        '
        'LB_output
        '
        Me.LB_output.FormattingEnabled = True
        Me.LB_output.Items.AddRange(New Object() {"Maximum upper layer temperature", "Time for layer hight to drop below 2 m", "Time for visibility to drop below 10 m", "Time for visibility to drop below 5 m", "Time for FEDCO to exceed 0.3", "Time for FED thermal to exceed 0.3", "Time for upper layer temperature to exceed 500 C"})
        Me.LB_output.Location = New System.Drawing.Point(12, 177)
        Me.LB_output.Name = "LB_output"
        Me.LB_output.Size = New System.Drawing.Size(303, 95)
        Me.LB_output.TabIndex = 6
        '
        'cmdPlotSENS
        '
        Me.cmdPlotSENS.Location = New System.Drawing.Point(159, 287)
        Me.cmdPlotSENS.Name = "cmdPlotSENS"
        Me.cmdPlotSENS.Size = New System.Drawing.Size(75, 23)
        Me.cmdPlotSENS.TabIndex = 7
        Me.cmdPlotSENS.Text = "Plot"
        Me.cmdPlotSENS.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 9)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(147, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Select your parameters to plot"
        '
        'frmSensitivity
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(326, 322)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cmdPlotSENS)
        Me.Controls.Add(Me.LB_output)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.LB_input_ID)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.LB_input)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frmSensitivity"
        Me.Text = "Sensitivity Analysis Parameters"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents LB_input As ListBox
    Friend WithEvents cmdClose As Button
    Friend WithEvents LB_input_ID As ListBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents LB_output As ListBox
    Public WithEvents cmdPlotSENS As Button
    Friend WithEvents Label4 As Label
End Class
