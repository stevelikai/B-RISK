<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFireResistance
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.txtMaxOpening = New System.Windows.Forms.TextBox
        Me.txtMaxOpeningTime = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cmdDist_integrity = New System.Windows.Forms.Button
        Me.cmdDist_maxopening = New System.Windows.Forms.Button
        Me.cmdDist_maxopeningtime = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtFRintegrity = New System.Windows.Forms.TextBox
        Me.txtFRgastemp = New System.Windows.Forms.TextBox
        Me.cmdDist_gastemp = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.cmdClose = New System.Windows.Forms.Button
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.chkFireBarrierModel = New System.Windows.Forms.CheckBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cboFRcriterion = New System.Windows.Forms.ComboBox
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 3
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle)
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle)
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle)
        Me.TableLayoutPanel1.Controls.Add(Me.txtMaxOpening, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.txtMaxOpeningTime, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.cmdDist_integrity, 2, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.cmdDist_maxopening, 2, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.cmdDist_maxopeningtime, 2, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label2, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label3, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.txtFRintegrity, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.txtFRgastemp, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.cmdDist_gastemp, 2, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label4, 0, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(14, 148)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 5
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle)
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle)
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle)
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle)
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(375, 121)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'txtMaxOpening
        '
        Me.txtMaxOpening.Location = New System.Drawing.Point(166, 61)
        Me.txtMaxOpening.Name = "txtMaxOpening"
        Me.txtMaxOpening.Size = New System.Drawing.Size(67, 20)
        Me.txtMaxOpening.TabIndex = 6
        Me.txtMaxOpening.Text = "100"
        '
        'txtMaxOpeningTime
        '
        Me.txtMaxOpeningTime.Location = New System.Drawing.Point(166, 90)
        Me.txtMaxOpeningTime.Name = "txtMaxOpeningTime"
        Me.txtMaxOpeningTime.Size = New System.Drawing.Size(67, 20)
        Me.txtMaxOpeningTime.TabIndex = 8
        Me.txtMaxOpeningTime.Text = "10"
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 37)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(151, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Fire Resistance - Integrity (min)"
        '
        'cmdDist_integrity
        '
        Me.cmdDist_integrity.Location = New System.Drawing.Point(239, 32)
        Me.cmdDist_integrity.Name = "cmdDist_integrity"
        Me.cmdDist_integrity.Size = New System.Drawing.Size(75, 23)
        Me.cmdDist_integrity.TabIndex = 5
        Me.cmdDist_integrity.Text = "Distribution"
        Me.cmdDist_integrity.UseVisualStyleBackColor = True
        '
        'cmdDist_maxopening
        '
        Me.cmdDist_maxopening.Location = New System.Drawing.Point(239, 61)
        Me.cmdDist_maxopening.Name = "cmdDist_maxopening"
        Me.cmdDist_maxopening.Size = New System.Drawing.Size(75, 23)
        Me.cmdDist_maxopening.TabIndex = 7
        Me.cmdDist_maxopening.Text = "Distribution"
        Me.cmdDist_maxopening.UseVisualStyleBackColor = True
        '
        'cmdDist_maxopeningtime
        '
        Me.cmdDist_maxopeningtime.Location = New System.Drawing.Point(239, 90)
        Me.cmdDist_maxopeningtime.Name = "cmdDist_maxopeningtime"
        Me.cmdDist_maxopeningtime.Size = New System.Drawing.Size(75, 23)
        Me.cmdDist_maxopeningtime.TabIndex = 9
        Me.cmdDist_maxopeningtime.Text = "Distribution"
        Me.cmdDist_maxopeningtime.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 66)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(111, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Maximum Opening (%)"
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 98)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(157, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Time to Reach Max Area (s)"
        '
        'txtFRintegrity
        '
        Me.txtFRintegrity.Location = New System.Drawing.Point(166, 32)
        Me.txtFRintegrity.Name = "txtFRintegrity"
        Me.txtFRintegrity.Size = New System.Drawing.Size(67, 20)
        Me.txtFRintegrity.TabIndex = 4
        Me.txtFRintegrity.Text = "30"
        '
        'txtFRgastemp
        '
        Me.txtFRgastemp.Location = New System.Drawing.Point(166, 3)
        Me.txtFRgastemp.Name = "txtFRgastemp"
        Me.txtFRgastemp.Size = New System.Drawing.Size(67, 20)
        Me.txtFRgastemp.TabIndex = 2
        Me.txtFRgastemp.Text = "200"
        '
        'cmdDist_gastemp
        '
        Me.cmdDist_gastemp.Location = New System.Drawing.Point(239, 3)
        Me.cmdDist_gastemp.Name = "cmdDist_gastemp"
        Me.cmdDist_gastemp.Size = New System.Drawing.Size(75, 23)
        Me.cmdDist_gastemp.TabIndex = 3
        Me.cmdDist_gastemp.Text = "Distribution"
        Me.cmdDist_gastemp.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(105, 13)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Gas Temperature (C)"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(314, 287)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 10
        Me.cmdClose.Text = "Save"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(236, 287)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 11
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'chkFireBarrierModel
        '
        Me.chkFireBarrierModel.AutoSize = True
        Me.chkFireBarrierModel.Location = New System.Drawing.Point(14, 24)
        Me.chkFireBarrierModel.Name = "chkFireBarrierModel"
        Me.chkFireBarrierModel.Size = New System.Drawing.Size(297, 17)
        Me.chkFireBarrierModel.TabIndex = 0
        Me.chkFireBarrierModel.Text = "Progessively open this vent upon reaching failure criterion"
        Me.chkFireBarrierModel.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cboFRcriterion)
        Me.GroupBox1.Location = New System.Drawing.Point(14, 59)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(377, 73)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Select Failure Criterion"
        '
        'cboFRcriterion
        '
        Me.cboFRcriterion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFRcriterion.FormattingEnabled = True
        Me.cboFRcriterion.Items.AddRange(New Object() {"Upper layer gas temperature", "Heat Load (using emissive power of fire gases)", "Normalised heat load (using Harmathy's furnace eqn)", "Ceiling normalised heat load", "Upper wall normalised heat load"})
        Me.cboFRcriterion.Location = New System.Drawing.Point(30, 29)
        Me.cboFRcriterion.Name = "cboFRcriterion"
        Me.cboFRcriterion.Size = New System.Drawing.Size(323, 21)
        Me.cboFRcriterion.TabIndex = 1
        '
        'frmFireResistance
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(410, 321)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.chkFireBarrierModel)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmFireResistance"
        Me.ShowInTaskbar = False
        Me.Text = "Fire Resistance and Opening Characteristics"
        Me.TopMost = True
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents txtFRintegrity As System.Windows.Forms.TextBox
    Friend WithEvents txtMaxOpening As System.Windows.Forms.TextBox
    Friend WithEvents txtMaxOpeningTime As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdDist_integrity As System.Windows.Forms.Button
    Friend WithEvents cmdDist_maxopening As System.Windows.Forms.Button
    Friend WithEvents cmdDist_maxopeningtime As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents chkFireBarrierModel As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtFRgastemp As System.Windows.Forms.TextBox
    Friend WithEvents cmdDist_gastemp As System.Windows.Forms.Button
    Friend WithEvents cboFRcriterion As System.Windows.Forms.ComboBox
End Class
