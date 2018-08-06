<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSprinklerList
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
        Me.lstSprinklers = New System.Windows.Forms.ListBox
        Me.cmdAddSprinkler = New System.Windows.Forms.Button
        Me.cmdRemoveSprinkler = New System.Windows.Forms.Button
        Me.cmdExit = New System.Windows.Forms.Button
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.cmdEditSprinkler = New System.Windows.Forms.Button
        Me.cmdCopySprinkler = New System.Windows.Forms.Button
        Me.txtSprReliability = New System.Windows.Forms.TextBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.cmdDist_sprcooling = New System.Windows.Forms.Button
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.txtSprCoolingCoeff = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.cmdNoSprink = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.cmdDist_sprsuppression = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtSprSuppressProb = New System.Windows.Forms.TextBox
        Me.cmdDist_sprreliability = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.chkCalcSprinkRadialDist = New System.Windows.Forms.CheckBox
        Me.cmdRemoveAll = New System.Windows.Forms.Button
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.GroupBox2.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lstSprinklers
        '
        Me.lstSprinklers.FormattingEnabled = True
        Me.lstSprinklers.Location = New System.Drawing.Point(12, 25)
        Me.lstSprinklers.Name = "lstSprinklers"
        Me.lstSprinklers.Size = New System.Drawing.Size(353, 160)
        Me.lstSprinklers.TabIndex = 0
        '
        'cmdAddSprinkler
        '
        Me.cmdAddSprinkler.Location = New System.Drawing.Point(370, 16)
        Me.cmdAddSprinkler.Name = "cmdAddSprinkler"
        Me.cmdAddSprinkler.Size = New System.Drawing.Size(170, 23)
        Me.cmdAddSprinkler.TabIndex = 1
        Me.cmdAddSprinkler.Text = "Add Standard Resp Sprinkler"
        Me.cmdAddSprinkler.UseVisualStyleBackColor = True
        '
        'cmdRemoveSprinkler
        '
        Me.cmdRemoveSprinkler.Location = New System.Drawing.Point(370, 162)
        Me.cmdRemoveSprinkler.Name = "cmdRemoveSprinkler"
        Me.cmdRemoveSprinkler.Size = New System.Drawing.Size(75, 23)
        Me.cmdRemoveSprinkler.TabIndex = 7
        Me.cmdRemoveSprinkler.Text = "Remove"
        Me.cmdRemoveSprinkler.UseVisualStyleBackColor = True
        '
        'cmdExit
        '
        Me.cmdExit.Location = New System.Drawing.Point(464, 372)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(75, 23)
        Me.cmdExit.TabIndex = 20
        Me.cmdExit.Text = "Close"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'ListBox1
        '
        Me.ListBox1.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(13, 7)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(352, 17)
        Me.ListBox1.TabIndex = 5
        Me.ListBox1.TabStop = False
        '
        'cmdEditSprinkler
        '
        Me.cmdEditSprinkler.Location = New System.Drawing.Point(371, 132)
        Me.cmdEditSprinkler.Name = "cmdEditSprinkler"
        Me.cmdEditSprinkler.Size = New System.Drawing.Size(74, 24)
        Me.cmdEditSprinkler.TabIndex = 5
        Me.cmdEditSprinkler.Text = "Edit"
        Me.cmdEditSprinkler.UseVisualStyleBackColor = True
        '
        'cmdCopySprinkler
        '
        Me.cmdCopySprinkler.Location = New System.Drawing.Point(465, 134)
        Me.cmdCopySprinkler.Name = "cmdCopySprinkler"
        Me.cmdCopySprinkler.Size = New System.Drawing.Size(75, 23)
        Me.cmdCopySprinkler.TabIndex = 6
        Me.cmdCopySprinkler.Text = "Copy"
        Me.cmdCopySprinkler.UseVisualStyleBackColor = True
        '
        'txtSprReliability
        '
        Me.txtSprReliability.BackColor = System.Drawing.SystemColors.Info
        Me.txtSprReliability.Location = New System.Drawing.Point(177, 38)
        Me.txtSprReliability.Name = "txtSprReliability"
        Me.txtSprReliability.Size = New System.Drawing.Size(64, 20)
        Me.txtSprReliability.TabIndex = 10
        Me.txtSprReliability.Tag = "Enter probability that the sprinkler system will activate"
        Me.txtSprReliability.Text = "1.00"
        Me.txtSprReliability.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cmdDist_sprcooling)
        Me.GroupBox2.Controls.Add(Me.LinkLabel1)
        Me.GroupBox2.Controls.Add(Me.txtSprCoolingCoeff)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.cmdNoSprink)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.cmdDist_sprsuppression)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.txtSprSuppressProb)
        Me.GroupBox2.Controls.Add(Me.cmdDist_sprreliability)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.txtSprReliability)
        Me.GroupBox2.Location = New System.Drawing.Point(10, 216)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(431, 179)
        Me.GroupBox2.TabIndex = 120
        Me.GroupBox2.TabStop = False
        '
        'cmdDist_sprcooling
        '
        Me.cmdDist_sprcooling.Location = New System.Drawing.Point(261, 104)
        Me.cmdDist_sprcooling.Name = "cmdDist_sprcooling"
        Me.cmdDist_sprcooling.Size = New System.Drawing.Size(68, 19)
        Me.cmdDist_sprcooling.TabIndex = 15
        Me.cmdDist_sprcooling.Text = "distribution"
        Me.cmdDist_sprcooling.UseVisualStyleBackColor = True
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(355, 106)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(57, 13)
        Me.LinkLabel1.TabIndex = 161
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Reference"
        '
        'txtSprCoolingCoeff
        '
        Me.txtSprCoolingCoeff.BackColor = System.Drawing.SystemColors.Info
        Me.txtSprCoolingCoeff.Location = New System.Drawing.Point(177, 103)
        Me.txtSprCoolingCoeff.Name = "txtSprCoolingCoeff"
        Me.txtSprCoolingCoeff.Size = New System.Drawing.Size(64, 20)
        Me.txtSprCoolingCoeff.TabIndex = 12
        Me.txtSprCoolingCoeff.Text = "1.00"
        Me.txtSprCoolingCoeff.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtSprCoolingCoeff, "See Crocker et al.  http://users.wpi.edu/~rangwala/lab/Ali/Lab_Students_files/Cro" & _
                "cker%20Fire%20Technology.pdf")
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(14, 103)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(139, 13)
        Me.Label4.TabIndex = 159
        Me.Label4.Text = "Sprinkler Cooling Coefficient"
        '
        'cmdNoSprink
        '
        Me.cmdNoSprink.Location = New System.Drawing.Point(261, 138)
        Me.cmdNoSprink.Name = "cmdNoSprink"
        Me.cmdNoSprink.Size = New System.Drawing.Size(68, 19)
        Me.cmdNoSprink.TabIndex = 16
        Me.cmdNoSprink.Text = "distribution"
        Me.cmdNoSprink.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 144)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(244, 13)
        Me.Label3.TabIndex = 156
        Me.Label3.Text = "No. Operating Sprinklers Required for Suppression"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdDist_sprsuppression
        '
        Me.cmdDist_sprsuppression.Location = New System.Drawing.Point(263, 71)
        Me.cmdDist_sprsuppression.Name = "cmdDist_sprsuppression"
        Me.cmdDist_sprsuppression.Size = New System.Drawing.Size(68, 19)
        Me.cmdDist_sprsuppression.TabIndex = 14
        Me.cmdDist_sprsuppression.Text = "distribution"
        Me.cmdDist_sprsuppression.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(25, 73)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(128, 13)
        Me.Label2.TabIndex = 154
        Me.Label2.Text = "Probability of Suppression"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtSprSuppressProb
        '
        Me.txtSprSuppressProb.BackColor = System.Drawing.SystemColors.Info
        Me.txtSprSuppressProb.Location = New System.Drawing.Point(177, 70)
        Me.txtSprSuppressProb.Name = "txtSprSuppressProb"
        Me.txtSprSuppressProb.Size = New System.Drawing.Size(64, 20)
        Me.txtSprSuppressProb.TabIndex = 11
        Me.txtSprSuppressProb.Text = "1.00"
        Me.txtSprSuppressProb.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtSprSuppressProb, "Enter probability that the sprinkler system will suppress the fire given the syst" & _
                "em has activated")
        '
        'cmdDist_sprreliability
        '
        Me.cmdDist_sprreliability.Location = New System.Drawing.Point(263, 38)
        Me.cmdDist_sprreliability.Name = "cmdDist_sprreliability"
        Me.cmdDist_sprreliability.Size = New System.Drawing.Size(68, 19)
        Me.cmdDist_sprreliability.TabIndex = 13
        Me.cmdDist_sprreliability.Text = "distribution"
        Me.cmdDist_sprreliability.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(58, 41)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(95, 13)
        Me.Label1.TabIndex = 120
        Me.Label1.Text = "Sprinkler Reliability"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(370, 45)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(170, 23)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Add Quick Resp Sprinkler"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(370, 103)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(169, 23)
        Me.Button3.TabIndex = 4
        Me.Button3.Text = "Add Heat Detector"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(371, 74)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(170, 23)
        Me.Button4.TabIndex = 3
        Me.Button4.Text = "Add Ext Coverage Sprinkler"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'chkCalcSprinkRadialDist
        '
        Me.chkCalcSprinkRadialDist.AutoSize = True
        Me.chkCalcSprinkRadialDist.Location = New System.Drawing.Point(13, 193)
        Me.chkCalcSprinkRadialDist.Name = "chkCalcSprinkRadialDist"
        Me.chkCalcSprinkRadialDist.Size = New System.Drawing.Size(324, 17)
        Me.chkCalcSprinkRadialDist.TabIndex = 9
        Me.chkCalcSprinkRadialDist.Text = "Calculate fire to sensor radial distance (overrides sensor setting)"
        Me.chkCalcSprinkRadialDist.UseVisualStyleBackColor = True
        '
        'cmdRemoveAll
        '
        Me.cmdRemoveAll.Location = New System.Drawing.Point(464, 162)
        Me.cmdRemoveAll.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdRemoveAll.Name = "cmdRemoveAll"
        Me.cmdRemoveAll.Size = New System.Drawing.Size(74, 22)
        Me.cmdRemoveAll.TabIndex = 8
        Me.cmdRemoveAll.Text = "Remove All"
        Me.cmdRemoveAll.UseVisualStyleBackColor = True
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'frmSprinklerList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(552, 407)
        Me.Controls.Add(Me.cmdRemoveAll)
        Me.Controls.Add(Me.chkCalcSprinkRadialDist)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.cmdCopySprinkler)
        Me.Controls.Add(Me.cmdEditSprinkler)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdRemoveSprinkler)
        Me.Controls.Add(Me.cmdAddSprinkler)
        Me.Controls.Add(Me.lstSprinklers)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MinimizeBox = False
        Me.Name = "frmSprinklerList"
        Me.Text = "Sprinklers and Heat Detectors"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lstSprinklers As System.Windows.Forms.ListBox
    Friend WithEvents cmdAddSprinkler As System.Windows.Forms.Button
    Friend WithEvents cmdRemoveSprinkler As System.Windows.Forms.Button
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents cmdEditSprinkler As System.Windows.Forms.Button
    Friend WithEvents cmdCopySprinkler As System.Windows.Forms.Button
    Friend WithEvents txtSprReliability As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdDist_sprreliability As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents txtSprSuppressProb As System.Windows.Forms.TextBox
    Friend WithEvents cmdDist_sprsuppression As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmdNoSprink As System.Windows.Forms.Button
    Friend WithEvents txtSprCoolingCoeff As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents cmdDist_sprcooling As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents chkCalcSprinkRadialDist As System.Windows.Forms.CheckBox
    Friend WithEvents cmdRemoveAll As System.Windows.Forms.Button
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
End Class
