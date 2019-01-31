<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmSplash
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

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim Frame1 As System.Windows.Forms.GroupBox
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSplash))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.lblCopyright = New System.Windows.Forms.Label()
        Me.lblCompany = New System.Windows.Forms.Label()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.lblPlatform = New System.Windows.Forms.Label()
        Me.lblProductName = New System.Windows.Forms.Label()
        Me.lblCompanyProduct = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Frame1 = New System.Windows.Forms.GroupBox()
        Frame1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Frame1.BackColor = System.Drawing.Color.Black
        Frame1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Frame1.Controls.Add(Me.PictureBox1)
        Frame1.Controls.Add(Me.lblCopyright)
        Frame1.Controls.Add(Me.lblCompany)
        Frame1.Controls.Add(Me.lblVersion)
        Frame1.Controls.Add(Me.lblPlatform)
        Frame1.Controls.Add(Me.lblProductName)
        Frame1.Controls.Add(Me.lblCompanyProduct)
        Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Frame1.Location = New System.Drawing.Point(-9, -8)
        Frame1.Name = "Frame1"
        Frame1.Padding = New System.Windows.Forms.Padding(0)
        Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Frame1.Size = New System.Drawing.Size(521, 283)
        Frame1.TabIndex = 0
        Frame1.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = CType(resources.GetObject("PictureBox1.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PictureBox1.ErrorImage = Nothing
        Me.PictureBox1.InitialImage = Nothing
        Me.PictureBox1.Location = New System.Drawing.Point(26, 61)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(136, 197)
        Me.PictureBox1.TabIndex = 8
        Me.PictureBox1.TabStop = False
        '
        'lblCopyright
        '
        Me.lblCopyright.BackColor = System.Drawing.Color.Black
        Me.lblCopyright.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCopyright.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCopyright.ForeColor = System.Drawing.Color.Transparent
        Me.lblCopyright.Location = New System.Drawing.Point(227, 208)
        Me.lblCopyright.Name = "lblCopyright"
        Me.lblCopyright.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCopyright.Size = New System.Drawing.Size(107, 13)
        Me.lblCopyright.TabIndex = 3
        Me.lblCopyright.Text = "Copyright 2019"
        '
        'lblCompany
        '
        Me.lblCompany.BackColor = System.Drawing.Color.Black
        Me.lblCompany.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCompany.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompany.ForeColor = System.Drawing.Color.Transparent
        Me.lblCompany.Location = New System.Drawing.Point(227, 193)
        Me.lblCompany.Name = "lblCompany"
        Me.lblCompany.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCompany.Size = New System.Drawing.Size(233, 15)
        Me.lblCompany.TabIndex = 2
        Me.lblCompany.Text = "BRANZ Ltd"
        '
        'lblVersion
        '
        Me.lblVersion.AutoSize = True
        Me.lblVersion.BackColor = System.Drawing.Color.Black
        Me.lblVersion.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVersion.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVersion.ForeColor = System.Drawing.Color.Transparent
        Me.lblVersion.Location = New System.Drawing.Point(352, 98)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVersion.Size = New System.Drawing.Size(107, 19)
        Me.lblVersion.TabIndex = 4
        Me.lblVersion.Text = "Version 2019"
        Me.lblVersion.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblPlatform
        '
        Me.lblPlatform.AutoSize = True
        Me.lblPlatform.BackColor = System.Drawing.Color.Black
        Me.lblPlatform.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPlatform.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPlatform.ForeColor = System.Drawing.Color.Transparent
        Me.lblPlatform.Location = New System.Drawing.Point(226, 155)
        Me.lblPlatform.Name = "lblPlatform"
        Me.lblPlatform.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPlatform.Size = New System.Drawing.Size(218, 22)
        Me.lblPlatform.TabIndex = 5
        Me.lblPlatform.Text = "for Microsoft Windows"
        Me.lblPlatform.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblProductName
        '
        Me.lblProductName.AutoSize = True
        Me.lblProductName.BackColor = System.Drawing.Color.Black
        Me.lblProductName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblProductName.Font = New System.Drawing.Font("Arial Black", 32.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProductName.ForeColor = System.Drawing.Color.Firebrick
        Me.lblProductName.Location = New System.Drawing.Point(168, 66)
        Me.lblProductName.Name = "lblProductName"
        Me.lblProductName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblProductName.Size = New System.Drawing.Size(189, 60)
        Me.lblProductName.TabIndex = 7
        Me.lblProductName.Text = "B-RISK"
        '
        'lblCompanyProduct
        '
        Me.lblCompanyProduct.AutoSize = True
        Me.lblCompanyProduct.BackColor = System.Drawing.Color.Black
        Me.lblCompanyProduct.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCompanyProduct.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompanyProduct.ForeColor = System.Drawing.Color.Transparent
        Me.lblCompanyProduct.Location = New System.Drawing.Point(174, 47)
        Me.lblCompanyProduct.Name = "lblCompanyProduct"
        Me.lblCompanyProduct.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCompanyProduct.Size = New System.Drawing.Size(302, 19)
        Me.lblCompanyProduct.TabIndex = 6
        Me.lblCompanyProduct.Text = "design fire tool and room fire simulator"
        '
        'Timer1
        '
        Me.Timer1.Interval = 3000
        '
        'frmSplash
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(503, 262)
        Me.ControlBox = False
        Me.Controls.Add(Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(17, 94)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSplash"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Frame1.ResumeLayout(False)
        Frame1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents lblCompanyProduct As System.Windows.Forms.Label
    Public WithEvents lblProductName As System.Windows.Forms.Label
    Public WithEvents lblPlatform As System.Windows.Forms.Label
    Public WithEvents lblVersion As System.Windows.Forms.Label
    Public WithEvents lblCompany As System.Windows.Forms.Label
    Public WithEvents lblCopyright As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Public WithEvents Timer1 As System.Windows.Forms.Timer
#End Region
End Class