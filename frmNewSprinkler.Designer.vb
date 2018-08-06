<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNewSprinkler
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
        Me.txtValue8 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtValue9 = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.txtValue11 = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtValue12 = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtValue7 = New System.Windows.Forms.TextBox
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown
        Me.Label10 = New System.Windows.Forms.Label
        Me.cmdSave = New System.Windows.Forms.Button
        Me.txtValue10 = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.cmdDist_sprr = New System.Windows.Forms.Button
        Me.cmdDist_sprz = New System.Windows.Forms.Button
        Me.cmdDist_sprdensity = New System.Windows.Forms.Button
        Me.Label7 = New System.Windows.Forms.Label
        Me.cmdDist_cfactor = New System.Windows.Forms.Button
        Me.cmdDist_acttemp = New System.Windows.Forms.Button
        Me.cmdDist_rti = New System.Windows.Forms.Button
        Me.Label31 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label33 = New System.Windows.Forms.Label
        Me.Label36 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label30 = New System.Windows.Forms.Label
        Me.Label37 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtValue8
        '
        Me.txtValue8.Location = New System.Drawing.Point(124, 49)
        Me.txtValue8.Name = "txtValue8"
        Me.txtValue8.Size = New System.Drawing.Size(62, 20)
        Me.txtValue8.TabIndex = 2
        Me.txtValue8.Tag = "rti"
        Me.txtValue8.Text = "50"
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(93, 52)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(25, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "RTI"
        '
        'Label4
        '
        Me.Label4.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(78, 104)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 13)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "cfactor"
        '
        'txtValue9
        '
        Me.txtValue9.Location = New System.Drawing.Point(124, 101)
        Me.txtValue9.Name = "txtValue9"
        Me.txtValue9.Size = New System.Drawing.Size(62, 20)
        Me.txtValue9.TabIndex = 4
        Me.txtValue9.Tag = "cfactor"
        Me.txtValue9.Text = "0.4"
        '
        'Label5
        '
        Me.Label5.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(47, 182)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(71, 13)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "x - coordinate"
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(124, 179)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(62, 20)
        Me.TextBox3.TabIndex = 7
        Me.TextBox3.Tag = "x"
        Me.TextBox3.Text = "0"
        Me.ToolTip1.SetToolTip(Me.TextBox3, "only used if 'Auto Populate Room Contents' is ON")
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(124, 205)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(62, 20)
        Me.TextBox4.TabIndex = 8
        Me.TextBox4.Tag = "y"
        Me.TextBox4.Text = "0"
        Me.ToolTip1.SetToolTip(Me.TextBox4, "only used if 'Auto Populate Room Contents' is ON")
        '
        'txtValue11
        '
        Me.txtValue11.Location = New System.Drawing.Point(124, 231)
        Me.txtValue11.Name = "txtValue11"
        Me.txtValue11.Size = New System.Drawing.Size(62, 20)
        Me.txtValue11.TabIndex = 9
        Me.txtValue11.Tag = "z"
        Me.txtValue11.Text = "0"
        '
        'Label6
        '
        Me.Label6.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(47, 208)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(71, 13)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "y - coordinate"
        '
        'txtValue12
        '
        Me.txtValue12.Location = New System.Drawing.Point(124, 127)
        Me.txtValue12.Name = "txtValue12"
        Me.txtValue12.Size = New System.Drawing.Size(62, 20)
        Me.txtValue12.TabIndex = 5
        Me.txtValue12.Tag = "density"
        Me.txtValue12.Text = "4.2"
        '
        'Label8
        '
        Me.Label8.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(21, 130)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(97, 13)
        Me.Label8.TabIndex = 13
        Me.Label8.Text = "water spray density"
        '
        'txtValue7
        '
        Me.txtValue7.Location = New System.Drawing.Point(124, 153)
        Me.txtValue7.Name = "txtValue7"
        Me.txtValue7.Size = New System.Drawing.Size(62, 20)
        Me.txtValue7.TabIndex = 6
        Me.txtValue7.Tag = "r"
        Me.txtValue7.Text = "0"
        Me.ToolTip1.SetToolTip(Me.txtValue7, "only used if 'Auto Populate Room Contents' is OFF ")
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(192, 293)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 16
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Location = New System.Drawing.Point(124, 23)
        Me.NumericUpDown1.Maximum = New Decimal(New Integer() {12, 0, 0, 0})
        Me.NumericUpDown1.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(62, 20)
        Me.NumericUpDown1.TabIndex = 0
        Me.NumericUpDown1.Tag = "room"
        Me.NumericUpDown1.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label10
        '
        Me.Label10.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(88, 26)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(30, 13)
        Me.Label10.TabIndex = 18
        Me.Label10.Text = "room"
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(280, 293)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(75, 23)
        Me.cmdSave.TabIndex = 17
        Me.cmdSave.Text = "Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'txtValue10
        '
        Me.txtValue10.Location = New System.Drawing.Point(124, 75)
        Me.txtValue10.Name = "txtValue10"
        Me.txtValue10.Size = New System.Drawing.Size(62, 20)
        Me.txtValue10.TabIndex = 3
        Me.txtValue10.Tag = "acttemp"
        Me.txtValue10.Text = "74"
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(39, 78)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 13)
        Me.Label1.TabIndex = 21
        Me.Label1.Text = "activation temp"
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.AutoSize = True
        Me.TableLayoutPanel1.ColumnCount = 4
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle)
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle)
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle)
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 85.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.cmdDist_sprr, 3, 6)
        Me.TableLayoutPanel1.Controls.Add(Me.cmdDist_sprz, 3, 9)
        Me.TableLayoutPanel1.Controls.Add(Me.cmdDist_sprdensity, 3, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.Label7, 2, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.NumericUpDown1, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.cmdDist_cfactor, 3, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.Label10, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.cmdDist_acttemp, 3, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.txtValue11, 1, 9)
        Me.TableLayoutPanel1.Controls.Add(Me.cmdDist_rti, 3, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label6, 0, 8)
        Me.TableLayoutPanel1.Controls.Add(Me.txtValue10, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.txtValue7, 1, 6)
        Me.TableLayoutPanel1.Controls.Add(Me.txtValue8, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label5, 0, 7)
        Me.TableLayoutPanel1.Controls.Add(Me.TextBox3, 1, 7)
        Me.TableLayoutPanel1.Controls.Add(Me.Label8, 0, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.Label3, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.txtValue12, 1, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.txtValue9, 1, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.Label4, 0, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.TextBox4, 1, 8)
        Me.TableLayoutPanel1.Controls.Add(Me.Label31, 2, 6)
        Me.TableLayoutPanel1.Controls.Add(Me.Label2, 2, 7)
        Me.TableLayoutPanel1.Controls.Add(Me.Label11, 2, 8)
        Me.TableLayoutPanel1.Controls.Add(Me.Label12, 2, 9)
        Me.TableLayoutPanel1.Controls.Add(Me.Label33, 2, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label36, 2, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label13, 2, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label30, 0, 6)
        Me.TableLayoutPanel1.Controls.Add(Me.Label37, 0, 9)
        Me.TableLayoutPanel1.Controls.Add(Me.Label9, 2, 4)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(21, 12)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 10
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle)
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle)
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle)
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle)
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle)
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle)
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle)
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle)
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle)
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(334, 260)
        Me.TableLayoutPanel1.TabIndex = 22
        '
        'cmdDist_sprr
        '
        Me.cmdDist_sprr.Location = New System.Drawing.Point(252, 153)
        Me.cmdDist_sprr.Name = "cmdDist_sprr"
        Me.cmdDist_sprr.Size = New System.Drawing.Size(68, 19)
        Me.cmdDist_sprr.TabIndex = 14
        Me.cmdDist_sprr.Text = "distribution"
        Me.cmdDist_sprr.UseVisualStyleBackColor = True
        '
        'cmdDist_sprz
        '
        Me.cmdDist_sprz.Location = New System.Drawing.Point(252, 231)
        Me.cmdDist_sprz.Name = "cmdDist_sprz"
        Me.cmdDist_sprz.Size = New System.Drawing.Size(68, 19)
        Me.cmdDist_sprz.TabIndex = 15
        Me.cmdDist_sprz.Text = "distribution"
        Me.cmdDist_sprz.UseVisualStyleBackColor = True
        '
        'cmdDist_sprdensity
        '
        Me.cmdDist_sprdensity.Location = New System.Drawing.Point(252, 127)
        Me.cmdDist_sprdensity.Name = "cmdDist_sprdensity"
        Me.cmdDist_sprdensity.Size = New System.Drawing.Size(68, 19)
        Me.cmdDist_sprdensity.TabIndex = 13
        Me.cmdDist_sprdensity.Text = "distribution"
        Me.cmdDist_sprdensity.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(192, 124)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(54, 26)
        Me.Label7.TabIndex = 103
        Me.Label7.Text = "mm/min"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmdDist_cfactor
        '
        Me.cmdDist_cfactor.Location = New System.Drawing.Point(252, 101)
        Me.cmdDist_cfactor.Name = "cmdDist_cfactor"
        Me.cmdDist_cfactor.Size = New System.Drawing.Size(68, 19)
        Me.cmdDist_cfactor.TabIndex = 12
        Me.cmdDist_cfactor.Text = "distribution"
        Me.cmdDist_cfactor.UseVisualStyleBackColor = True
        '
        'cmdDist_acttemp
        '
        Me.cmdDist_acttemp.Location = New System.Drawing.Point(252, 75)
        Me.cmdDist_acttemp.Name = "cmdDist_acttemp"
        Me.cmdDist_acttemp.Size = New System.Drawing.Size(68, 19)
        Me.cmdDist_acttemp.TabIndex = 11
        Me.cmdDist_acttemp.Text = "distribution"
        Me.cmdDist_acttemp.UseVisualStyleBackColor = True
        '
        'cmdDist_rti
        '
        Me.cmdDist_rti.Location = New System.Drawing.Point(252, 49)
        Me.cmdDist_rti.Name = "cmdDist_rti"
        Me.cmdDist_rti.Size = New System.Drawing.Size(68, 19)
        Me.cmdDist_rti.TabIndex = 10
        Me.cmdDist_rti.Text = "distribution"
        Me.cmdDist_rti.UseVisualStyleBackColor = True
        '
        'Label31
        '
        Me.Label31.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label31.AutoSize = True
        Me.Label31.Location = New System.Drawing.Point(192, 150)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(54, 26)
        Me.Label31.TabIndex = 79
        Me.Label31.Text = "m"
        Me.Label31.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(192, 176)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(54, 26)
        Me.Label2.TabIndex = 80
        Me.Label2.Text = "m"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label11
        '
        Me.Label11.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(192, 202)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(54, 26)
        Me.Label11.TabIndex = 81
        Me.Label11.Text = "m"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label12
        '
        Me.Label12.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(192, 228)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(54, 32)
        Me.Label12.TabIndex = 82
        Me.Label12.Text = "m"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label33
        '
        Me.Label33.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label33.AutoSize = True
        Me.Label33.Location = New System.Drawing.Point(192, 46)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(54, 26)
        Me.Label33.TabIndex = 87
        Me.Label33.Text = "(ms)^1/2"
        Me.Label33.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label36
        '
        Me.Label36.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label36.AutoSize = True
        Me.Label36.Location = New System.Drawing.Point(192, 72)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(54, 26)
        Me.Label36.TabIndex = 102
        Me.Label36.Text = "deg C"
        Me.Label36.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label13
        '
        Me.Label13.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(192, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(54, 20)
        Me.Label13.TabIndex = 103
        Me.Label13.Text = "Units"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label30
        '
        Me.Label30.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label30.AutoSize = True
        Me.Label30.Location = New System.Drawing.Point(43, 156)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(75, 13)
        Me.Label30.TabIndex = 139
        Me.Label30.Text = "radial distance"
        Me.Label30.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label37
        '
        Me.Label37.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label37.AutoSize = True
        Me.Label37.Location = New System.Drawing.Point(3, 228)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(115, 32)
        Me.Label37.TabIndex = 140
        Me.Label37.Text = "Distance Below Ceiling"
        Me.Label37.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label9
        '
        Me.Label9.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(192, 104)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(54, 13)
        Me.Label9.TabIndex = 200
        Me.Label9.Text = "(m/s)^1/2"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'frmNewSprinkler
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(383, 328)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdCancel)
        Me.Name = "frmNewSprinkler"
        Me.Text = "New Sprinkler"
        Me.TopMost = True
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtValue8 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtValue9 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents txtValue11 As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtValue12 As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtValue7 As System.Windows.Forms.TextBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents NumericUpDown1 As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents txtValue10 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cmdDist_rti As System.Windows.Forms.Button
    Friend WithEvents cmdDist_acttemp As System.Windows.Forms.Button
    Friend WithEvents cmdDist_cfactor As System.Windows.Forms.Button
    Friend WithEvents cmdDist_sprdensity As System.Windows.Forms.Button
    Friend WithEvents cmdDist_sprr As System.Windows.Forms.Button
    Friend WithEvents cmdDist_sprz As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
End Class
