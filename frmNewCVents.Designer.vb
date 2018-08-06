<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNewCVents
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
        Me.components = New System.ComponentModel.Container()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.optCventAuto = New System.Windows.Forms.RadioButton()
        Me.optCventManual = New System.Windows.Forms.RadioButton()
        Me.cmdVentOptions = New System.Windows.Forms.Button()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.lstCVentIDlower = New System.Windows.Forms.ComboBox()
        Me.lstCVentIDUpper = New System.Windows.Forms.ComboBox()
        Me.lblCVentID = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.cmdDist_cvarea = New System.Windows.Forms.Button()
        Me.txtCVentArea = New System.Windows.Forms.TextBox()
        Me.lblVentArea = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtVentDescription = New System.Windows.Forms.TextBox()
        Me.lblVentOpenTime = New System.Windows.Forms.Label()
        Me.txtCVentOpenTime = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtCVentCloseTime = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtDischargeCoeff = New System.Windows.Forms.TextBox()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.cmdFireResistance = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        Me.Frame4.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.optCventAuto)
        Me.Frame1.Controls.Add(Me.optCventManual)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(12, 124)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(235, 87)
        Me.Frame1.TabIndex = 99
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Vent Opening Method"
        Me.Frame1.Visible = False
        '
        'optCventAuto
        '
        Me.optCventAuto.BackColor = System.Drawing.SystemColors.Control
        Me.optCventAuto.Cursor = System.Windows.Forms.Cursors.Default
        Me.optCventAuto.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optCventAuto.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optCventAuto.Location = New System.Drawing.Point(24, 49)
        Me.optCventAuto.Name = "optCventAuto"
        Me.optCventAuto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optCventAuto.Size = New System.Drawing.Size(208, 25)
        Me.optCventAuto.TabIndex = 88
        Me.optCventAuto.TabStop = True
        Me.optCventAuto.Text = "Automatic (smoke or heat detector)"
        Me.optCventAuto.UseVisualStyleBackColor = False
        '
        'optCventManual
        '
        Me.optCventManual.BackColor = System.Drawing.SystemColors.Control
        Me.optCventManual.Checked = True
        Me.optCventManual.Cursor = System.Windows.Forms.Cursors.Default
        Me.optCventManual.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optCventManual.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optCventManual.Location = New System.Drawing.Point(24, 24)
        Me.optCventManual.Name = "optCventManual"
        Me.optCventManual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optCventManual.Size = New System.Drawing.Size(73, 25)
        Me.optCventManual.TabIndex = 87
        Me.optCventManual.TabStop = True
        Me.optCventManual.Text = "Manual"
        Me.optCventManual.UseVisualStyleBackColor = False
        '
        'cmdVentOptions
        '
        Me.cmdVentOptions.BackColor = System.Drawing.SystemColors.Control
        Me.cmdVentOptions.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdVentOptions.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdVentOptions.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdVentOptions.Location = New System.Drawing.Point(253, 186)
        Me.cmdVentOptions.Name = "cmdVentOptions"
        Me.cmdVentOptions.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdVentOptions.Size = New System.Drawing.Size(137, 25)
        Me.cmdVentOptions.TabIndex = 89
        Me.cmdVentOptions.Text = "Vent Opening Options"
        Me.cmdVentOptions.UseVisualStyleBackColor = False
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.lstCVentIDlower)
        Me.Frame4.Controls.Add(Me.lstCVentIDUpper)
        Me.Frame4.Controls.Add(Me.lblCVentID)
        Me.Frame4.Controls.Add(Me.Label14)
        Me.Frame4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(12, 12)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(232, 106)
        Me.Frame4.TabIndex = 100
        Me.Frame4.TabStop = False
        '
        'lstCVentIDlower
        '
        Me.lstCVentIDlower.BackColor = System.Drawing.Color.White
        Me.lstCVentIDlower.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstCVentIDlower.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstCVentIDlower.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstCVentIDlower.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstCVentIDlower.Location = New System.Drawing.Point(136, 64)
        Me.lstCVentIDlower.Name = "lstCVentIDlower"
        Me.lstCVentIDlower.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstCVentIDlower.Size = New System.Drawing.Size(73, 22)
        Me.lstCVentIDlower.TabIndex = 93
        '
        'lstCVentIDUpper
        '
        Me.lstCVentIDUpper.BackColor = System.Drawing.Color.White
        Me.lstCVentIDUpper.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstCVentIDUpper.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstCVentIDUpper.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstCVentIDUpper.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstCVentIDUpper.Location = New System.Drawing.Point(136, 24)
        Me.lstCVentIDUpper.Name = "lstCVentIDUpper"
        Me.lstCVentIDUpper.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstCVentIDUpper.Size = New System.Drawing.Size(73, 22)
        Me.lstCVentIDUpper.TabIndex = 92
        '
        'lblCVentID
        '
        Me.lblCVentID.BackColor = System.Drawing.SystemColors.Control
        Me.lblCVentID.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCVentID.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCVentID.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCVentID.Location = New System.Drawing.Point(32, 64)
        Me.lblCVentID.Name = "lblCVentID"
        Me.lblCVentID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCVentID.Size = New System.Drawing.Size(97, 25)
        Me.lblCVentID.TabIndex = 96
        Me.lblCVentID.Text = "to lower room"
        Me.lblCVentID.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.SystemColors.Control
        Me.Label14.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label14.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label14.Location = New System.Drawing.Point(8, 24)
        Me.Label14.Name = "Label14"
        Me.Label14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label14.Size = New System.Drawing.Size(121, 17)
        Me.Label14.TabIndex = 95
        Me.Label14.Text = "Connecting upper room"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(473, 217)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(75, 23)
        Me.cmdSave.TabIndex = 147
        Me.cmdSave.Text = "Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 3
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 119.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.cmdDist_cvarea, 2, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.txtCVentArea, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.lblVentArea, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label4, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.txtVentDescription, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.lblVentOpenTime, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.txtCVentOpenTime, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label10, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.txtCVentCloseTime, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.txtDischargeCoeff, 1, 4)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(253, 21)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 5
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(303, 148)
        Me.TableLayoutPanel1.TabIndex = 148
        '
        'cmdDist_cvarea
        '
        Me.cmdDist_cvarea.Location = New System.Drawing.Point(187, 33)
        Me.cmdDist_cvarea.Name = "cmdDist_cvarea"
        Me.cmdDist_cvarea.Size = New System.Drawing.Size(93, 22)
        Me.cmdDist_cvarea.TabIndex = 20
        Me.cmdDist_cvarea.Text = "distribution"
        Me.cmdDist_cvarea.UseVisualStyleBackColor = True
        '
        'txtCVentArea
        '
        Me.txtCVentArea.AcceptsReturn = True
        Me.txtCVentArea.BackColor = System.Drawing.Color.White
        Me.txtCVentArea.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCVentArea.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCVentArea.ForeColor = System.Drawing.Color.Black
        Me.txtCVentArea.Location = New System.Drawing.Point(132, 33)
        Me.txtCVentArea.MaxLength = 0
        Me.txtCVentArea.Name = "txtCVentArea"
        Me.txtCVentArea.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCVentArea.Size = New System.Drawing.Size(49, 20)
        Me.txtCVentArea.TabIndex = 4
        Me.txtCVentArea.Text = "1"
        '
        'lblVentArea
        '
        Me.lblVentArea.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblVentArea.BackColor = System.Drawing.SystemColors.Control
        Me.lblVentArea.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVentArea.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVentArea.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentArea.Location = New System.Drawing.Point(45, 30)
        Me.lblVentArea.Name = "lblVentArea"
        Me.lblVentArea.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVentArea.Size = New System.Drawing.Size(81, 30)
        Me.lblVentArea.TabIndex = 133
        Me.lblVentArea.Text = "Vent Area (m2)"
        Me.lblVentArea.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(65, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(61, 14)
        Me.Label4.TabIndex = 136
        Me.Label4.Text = "Description"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtVentDescription
        '
        Me.txtVentDescription.AcceptsReturn = True
        Me.txtVentDescription.BackColor = System.Drawing.SystemColors.Window
        Me.TableLayoutPanel1.SetColumnSpan(Me.txtVentDescription, 2)
        Me.txtVentDescription.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtVentDescription.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVentDescription.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVentDescription.Location = New System.Drawing.Point(132, 3)
        Me.txtVentDescription.MaxLength = 23
        Me.txtVentDescription.Name = "txtVentDescription"
        Me.txtVentDescription.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtVentDescription.Size = New System.Drawing.Size(151, 20)
        Me.txtVentDescription.TabIndex = 3
        Me.txtVentDescription.Text = "vent label"
        '
        'lblVentOpenTime
        '
        Me.lblVentOpenTime.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.lblVentOpenTime.AutoSize = True
        Me.lblVentOpenTime.BackColor = System.Drawing.Color.Transparent
        Me.lblVentOpenTime.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVentOpenTime.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVentOpenTime.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblVentOpenTime.Location = New System.Drawing.Point(3, 68)
        Me.lblVentOpenTime.Name = "lblVentOpenTime"
        Me.lblVentOpenTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVentOpenTime.Size = New System.Drawing.Size(123, 14)
        Me.lblVentOpenTime.TabIndex = 130
        Me.lblVentOpenTime.Text = "Vent OpeningTime (sec)"
        Me.lblVentOpenTime.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtCVentOpenTime
        '
        Me.txtCVentOpenTime.AcceptsReturn = True
        Me.txtCVentOpenTime.BackColor = System.Drawing.Color.White
        Me.txtCVentOpenTime.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCVentOpenTime.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCVentOpenTime.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCVentOpenTime.Location = New System.Drawing.Point(132, 63)
        Me.txtCVentOpenTime.MaxLength = 0
        Me.txtCVentOpenTime.Name = "txtCVentOpenTime"
        Me.txtCVentOpenTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCVentOpenTime.Size = New System.Drawing.Size(49, 20)
        Me.txtCVentOpenTime.TabIndex = 8
        Me.txtCVentOpenTime.Text = "0"
        '
        'Label10
        '
        Me.Label10.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label10.AutoSize = True
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(5, 98)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(121, 14)
        Me.Label10.TabIndex = 134
        Me.Label10.Text = "Vent Closing Time (sec)"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtCVentCloseTime
        '
        Me.txtCVentCloseTime.AcceptsReturn = True
        Me.txtCVentCloseTime.BackColor = System.Drawing.SystemColors.Window
        Me.txtCVentCloseTime.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCVentCloseTime.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCVentCloseTime.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCVentCloseTime.Location = New System.Drawing.Point(132, 93)
        Me.txtCVentCloseTime.MaxLength = 0
        Me.txtCVentCloseTime.Name = "txtCVentCloseTime"
        Me.txtCVentCloseTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCVentCloseTime.Size = New System.Drawing.Size(49, 20)
        Me.txtCVentCloseTime.TabIndex = 9
        Me.txtCVentCloseTime.Text = "0"
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 120)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(118, 28)
        Me.Label1.TabIndex = 137
        Me.Label1.Text = "Discharge Coefficent (-)"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtDischargeCoeff
        '
        Me.txtDischargeCoeff.Location = New System.Drawing.Point(132, 123)
        Me.txtDischargeCoeff.Name = "txtDischargeCoeff"
        Me.txtDischargeCoeff.Size = New System.Drawing.Size(49, 20)
        Me.txtDischargeCoeff.TabIndex = 138
        Me.txtDischargeCoeff.Text = "0.60"
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'cmdFireResistance
        '
        Me.ErrorProvider1.SetIconAlignment(Me.cmdFireResistance, System.Windows.Forms.ErrorIconAlignment.TopLeft)
        Me.cmdFireResistance.Location = New System.Drawing.Point(396, 188)
        Me.cmdFireResistance.Name = "cmdFireResistance"
        Me.cmdFireResistance.Size = New System.Drawing.Size(152, 23)
        Me.cmdFireResistance.TabIndex = 149
        Me.cmdFireResistance.TabStop = False
        Me.cmdFireResistance.Text = "Fire Resistance Settings"
        Me.cmdFireResistance.UseVisualStyleBackColor = True
        '
        'frmNewCVents
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(562, 248)
        Me.Controls.Add(Me.cmdVentOptions)
        Me.Controls.Add(Me.cmdFireResistance)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.Frame4)
        Me.Controls.Add(Me.Frame1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "frmNewCVents"
        Me.Text = "New Vent"
        Me.Frame1.ResumeLayout(False)
        Me.Frame4.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents cmdVentOptions As System.Windows.Forms.Button
    Public WithEvents optCventAuto As System.Windows.Forms.RadioButton
    Public WithEvents optCventManual As System.Windows.Forms.RadioButton
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents lstCVentIDlower As System.Windows.Forms.ComboBox
    Public WithEvents lstCVentIDUpper As System.Windows.Forms.ComboBox
    Public WithEvents lblCVentID As System.Windows.Forms.Label
    Public WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents cmdDist_cvarea As System.Windows.Forms.Button
    Public WithEvents txtCVentArea As System.Windows.Forms.TextBox
    Public WithEvents lblVentArea As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents txtVentDescription As System.Windows.Forms.TextBox
    Public WithEvents lblVentOpenTime As System.Windows.Forms.Label
    Public WithEvents txtCVentOpenTime As System.Windows.Forms.TextBox
    Public WithEvents Label10 As System.Windows.Forms.Label
    Public WithEvents txtCVentCloseTime As System.Windows.Forms.TextBox
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents cmdFireResistance As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtDischargeCoeff As System.Windows.Forms.TextBox
End Class
