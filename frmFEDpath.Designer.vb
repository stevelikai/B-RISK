<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFEDpath
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
        Me.txtFED_start_3 = New System.Windows.Forms.TextBox
        Me.txtFED_start_2 = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.lstFED_room_1 = New System.Windows.Forms.ComboBox
        Me.txtFED_start_1 = New System.Windows.Forms.TextBox
        Me.txtFED_end_1 = New System.Windows.Forms.TextBox
        Me.lstFED_room_2 = New System.Windows.Forms.ComboBox
        Me.lstFED_room_3 = New System.Windows.Forms.ComboBox
        Me.txtFED_end_2 = New System.Windows.Forms.TextBox
        Me.txtFED_end_3 = New System.Windows.Forms.TextBox
        Me.cmdClose = New System.Windows.Forms.Button
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.cmdUpdateFED = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 3
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle)
        Me.TableLayoutPanel1.Controls.Add(Me.txtFED_start_3, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.txtFED_start_2, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label2, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label3, 2, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.lstFED_room_1, 2, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.txtFED_start_1, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.txtFED_end_1, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.lstFED_room_2, 2, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.lstFED_room_3, 2, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.txtFED_end_2, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.txtFED_end_3, 1, 3)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(102, 34)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 4
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle)
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle)
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle)
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(374, 99)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'txtFED_start_3
        '
        Me.txtFED_start_3.Location = New System.Drawing.Point(3, 70)
        Me.txtFED_start_3.Name = "txtFED_start_3"
        Me.txtFED_start_3.Size = New System.Drawing.Size(100, 20)
        Me.txtFED_start_3.TabIndex = 11
        Me.txtFED_start_3.Text = "0"
        '
        'txtFED_start_2
        '
        Me.txtFED_start_2.Location = New System.Drawing.Point(3, 43)
        Me.txtFED_start_2.Name = "txtFED_start_2"
        Me.txtFED_start_2.Size = New System.Drawing.Size(100, 20)
        Me.txtFED_start_2.TabIndex = 10
        Me.txtFED_start_2.Text = "0"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(97, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Start Time (sec)"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(136, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(92, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "End Time (sec)"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(269, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(39, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Room"
        '
        'lstFED_room_1
        '
        Me.lstFED_room_1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstFED_room_1.FormattingEnabled = True
        Me.lstFED_room_1.Location = New System.Drawing.Point(269, 16)
        Me.lstFED_room_1.Name = "lstFED_room_1"
        Me.lstFED_room_1.Size = New System.Drawing.Size(101, 21)
        Me.lstFED_room_1.TabIndex = 3
        '
        'txtFED_start_1
        '
        Me.txtFED_start_1.Location = New System.Drawing.Point(3, 16)
        Me.txtFED_start_1.Name = "txtFED_start_1"
        Me.txtFED_start_1.Size = New System.Drawing.Size(100, 20)
        Me.txtFED_start_1.TabIndex = 4
        Me.txtFED_start_1.Text = "0"
        '
        'txtFED_end_1
        '
        Me.txtFED_end_1.Location = New System.Drawing.Point(136, 16)
        Me.txtFED_end_1.Name = "txtFED_end_1"
        Me.txtFED_end_1.Size = New System.Drawing.Size(100, 20)
        Me.txtFED_end_1.TabIndex = 5
        Me.txtFED_end_1.Text = "0"
        '
        'lstFED_room_2
        '
        Me.lstFED_room_2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstFED_room_2.FormattingEnabled = True
        Me.lstFED_room_2.Location = New System.Drawing.Point(269, 43)
        Me.lstFED_room_2.Name = "lstFED_room_2"
        Me.lstFED_room_2.Size = New System.Drawing.Size(101, 21)
        Me.lstFED_room_2.TabIndex = 6
        '
        'lstFED_room_3
        '
        Me.lstFED_room_3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstFED_room_3.FormattingEnabled = True
        Me.lstFED_room_3.Location = New System.Drawing.Point(269, 70)
        Me.lstFED_room_3.Name = "lstFED_room_3"
        Me.lstFED_room_3.Size = New System.Drawing.Size(101, 21)
        Me.lstFED_room_3.TabIndex = 7
        '
        'txtFED_end_2
        '
        Me.txtFED_end_2.Location = New System.Drawing.Point(136, 43)
        Me.txtFED_end_2.Name = "txtFED_end_2"
        Me.txtFED_end_2.Size = New System.Drawing.Size(100, 20)
        Me.txtFED_end_2.TabIndex = 8
        Me.txtFED_end_2.Text = "0"
        '
        'txtFED_end_3
        '
        Me.txtFED_end_3.Location = New System.Drawing.Point(136, 70)
        Me.txtFED_end_3.Name = "txtFED_end_3"
        Me.txtFED_end_3.Size = New System.Drawing.Size(100, 20)
        Me.txtFED_end_3.TabIndex = 9
        Me.txtFED_end_3.Text = "0"
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(401, 153)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 1
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'cmdUpdateFED
        '
        Me.cmdUpdateFED.Location = New System.Drawing.Point(298, 153)
        Me.cmdUpdateFED.Name = "cmdUpdateFED"
        Me.cmdUpdateFED.Size = New System.Drawing.Size(97, 23)
        Me.cmdUpdateFED.TabIndex = 2
        Me.cmdUpdateFED.Text = "Update FED's"
        Me.cmdUpdateFED.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(26, 53)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(58, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Segment 1"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(26, 80)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(58, 13)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Segment 2"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(26, 107)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(58, 13)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Segment 3"
        '
        'frmFEDpath
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(496, 192)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cmdUpdateFED)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frmFEDpath"
        Me.Text = "Critical Egress Path for FED Calculations"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lstFED_room_1 As System.Windows.Forms.ComboBox
    Friend WithEvents txtFED_start_1 As System.Windows.Forms.TextBox
    Friend WithEvents txtFED_end_1 As System.Windows.Forms.TextBox
    Friend WithEvents txtFED_start_3 As System.Windows.Forms.TextBox
    Friend WithEvents txtFED_start_2 As System.Windows.Forms.TextBox
    Friend WithEvents lstFED_room_2 As System.Windows.Forms.ComboBox
    Friend WithEvents lstFED_room_3 As System.Windows.Forms.ComboBox
    Friend WithEvents txtFED_end_2 As System.Windows.Forms.TextBox
    Friend WithEvents txtFED_end_3 As System.Windows.Forms.TextBox
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents cmdUpdateFED As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
End Class
