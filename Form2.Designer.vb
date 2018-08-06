<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
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
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOK = New System.Windows.Forms.Button
        Me.cmdRotate1 = New System.Windows.Forms.Button
        Me.cmdRotate2 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(290, 231)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 0
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(371, 231)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(75, 23)
        Me.cmdOK.TabIndex = 1
        Me.cmdOK.Text = "OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'cmdRotate1
        '
        Me.cmdRotate1.AutoSize = True
        Me.cmdRotate1.Location = New System.Drawing.Point(12, 12)
        Me.cmdRotate1.Name = "cmdRotate1"
        Me.cmdRotate1.Size = New System.Drawing.Size(82, 23)
        Me.cmdRotate1.TabIndex = 2
        Me.cmdRotate1.Text = "Rotate Wall 1"
        Me.cmdRotate1.UseVisualStyleBackColor = True
        '
        'cmdRotate2
        '
        Me.cmdRotate2.AutoSize = True
        Me.cmdRotate2.Location = New System.Drawing.Point(371, 12)
        Me.cmdRotate2.Name = "cmdRotate2"
        Me.cmdRotate2.Size = New System.Drawing.Size(82, 23)
        Me.cmdRotate2.TabIndex = 3
        Me.cmdRotate2.Text = "Rotate Wall 2"
        Me.cmdRotate2.UseVisualStyleBackColor = True
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(458, 266)
        Me.Controls.Add(Me.cmdRotate2)
        Me.Controls.Add(Me.cmdRotate1)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdCancel)
        Me.Name = "Form2"
        Me.Text = "Form2"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdRotate1 As System.Windows.Forms.Button
    Friend WithEvents cmdRotate2 As System.Windows.Forms.Button
End Class
