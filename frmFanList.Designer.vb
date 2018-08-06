<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFanList
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
        Me.lstFan = New System.Windows.Forms.ListBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtFanReliability = New System.Windows.Forms.TextBox
        Me.cmdDist_fanreliability = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.cmdExit = New System.Windows.Forms.Button
        Me.cmdRemoveFan = New System.Windows.Forms.Button
        Me.cmdAddFan = New System.Windows.Forms.Button
        Me.cmdCopyFan = New System.Windows.Forms.Button
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdEditFan = New System.Windows.Forms.Button
        Me.GroupBox2.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lstFan
        '
        Me.lstFan.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstFan.FormattingEnabled = True
        Me.lstFan.Location = New System.Drawing.Point(12, 35)
        Me.lstFan.Name = "lstFan"
        Me.lstFan.Size = New System.Drawing.Size(353, 147)
        Me.lstFan.TabIndex = 137
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 41)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(143, 13)
        Me.Label1.TabIndex = 120
        Me.Label1.Text = "Mech Vent System Reliability"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtFanReliability
        '
        Me.txtFanReliability.BackColor = System.Drawing.SystemColors.Window
        Me.txtFanReliability.Location = New System.Drawing.Point(171, 37)
        Me.txtFanReliability.Name = "txtFanReliability"
        Me.txtFanReliability.Size = New System.Drawing.Size(64, 20)
        Me.txtFanReliability.TabIndex = 119
        Me.txtFanReliability.Tag = "Enter probability that the mechanical ventilation system will activate"
        Me.txtFanReliability.Text = "1.00"
        Me.txtFanReliability.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cmdDist_fanreliability
        '
        Me.cmdDist_fanreliability.Location = New System.Drawing.Point(241, 38)
        Me.cmdDist_fanreliability.Name = "cmdDist_fanreliability"
        Me.cmdDist_fanreliability.Size = New System.Drawing.Size(68, 19)
        Me.cmdDist_fanreliability.TabIndex = 152
        Me.cmdDist_fanreliability.Text = "distribution"
        Me.cmdDist_fanreliability.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cmdDist_fanreliability)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.txtFanReliability)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 210)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(353, 79)
        Me.GroupBox2.TabIndex = 144
        Me.GroupBox2.TabStop = False
        '
        'ListBox1
        '
        Me.ListBox1.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(12, 12)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(352, 17)
        Me.ListBox1.TabIndex = 141
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'cmdExit
        '
        Me.cmdExit.Location = New System.Drawing.Point(371, 266)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(75, 23)
        Me.cmdExit.TabIndex = 140
        Me.cmdExit.Text = "Close"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'cmdRemoveFan
        '
        Me.cmdRemoveFan.Location = New System.Drawing.Point(370, 163)
        Me.cmdRemoveFan.Name = "cmdRemoveFan"
        Me.cmdRemoveFan.Size = New System.Drawing.Size(75, 23)
        Me.cmdRemoveFan.TabIndex = 139
        Me.cmdRemoveFan.Text = "Remove"
        Me.cmdRemoveFan.UseVisualStyleBackColor = True
        '
        'cmdAddFan
        '
        Me.cmdAddFan.Location = New System.Drawing.Point(369, 35)
        Me.cmdAddFan.Name = "cmdAddFan"
        Me.cmdAddFan.Size = New System.Drawing.Size(77, 23)
        Me.cmdAddFan.TabIndex = 138
        Me.cmdAddFan.Text = "Add Fan"
        Me.cmdAddFan.UseVisualStyleBackColor = True
        '
        'cmdCopyFan
        '
        Me.cmdCopyFan.Location = New System.Drawing.Point(371, 134)
        Me.cmdCopyFan.Name = "cmdCopyFan"
        Me.cmdCopyFan.Size = New System.Drawing.Size(75, 23)
        Me.cmdCopyFan.TabIndex = 143
        Me.cmdCopyFan.Text = "Copy"
        Me.cmdCopyFan.UseVisualStyleBackColor = True
        '
        'cmdEditFan
        '
        Me.cmdEditFan.Location = New System.Drawing.Point(371, 104)
        Me.cmdEditFan.Name = "cmdEditFan"
        Me.cmdEditFan.Size = New System.Drawing.Size(74, 24)
        Me.cmdEditFan.TabIndex = 142
        Me.cmdEditFan.Text = "Edit"
        Me.cmdEditFan.UseVisualStyleBackColor = True
        '
        'frmFanList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(460, 309)
        Me.Controls.Add(Me.lstFan)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdRemoveFan)
        Me.Controls.Add(Me.cmdAddFan)
        Me.Controls.Add(Me.cmdCopyFan)
        Me.Controls.Add(Me.cmdEditFan)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MinimizeBox = False
        Me.Name = "frmFanList"
        Me.Text = "Mechanical Ventilation - Fan List"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lstFan As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtFanReliability As System.Windows.Forms.TextBox
    Friend WithEvents cmdDist_fanreliability As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents cmdRemoveFan As System.Windows.Forms.Button
    Friend WithEvents cmdAddFan As System.Windows.Forms.Button
    Friend WithEvents cmdCopyFan As System.Windows.Forms.Button
    Friend WithEvents cmdEditFan As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
End Class
