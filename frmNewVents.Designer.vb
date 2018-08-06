<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNewVents
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
        Me.Label18 = New System.Windows.Forms.Label()
        Me.lstRoomFace = New System.Windows.Forms.ComboBox()
        Me.cmdOpenOptions = New System.Windows.Forms.Button()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblVentOpenTime = New System.Windows.Forms.Label()
        Me.lblVentSillHeight = New System.Windows.Forms.Label()
        Me.lblVentWidth = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.lstRoom1 = New System.Windows.Forms.ComboBox()
        Me.lstRoom2 = New System.Windows.Forms.ComboBox()
        Me.txtVentOpenTime = New System.Windows.Forms.TextBox()
        Me.txtVentHeight = New System.Windows.Forms.TextBox()
        Me.txtVentWidth = New System.Windows.Forms.TextBox()
        Me.txtVentSillHeight = New System.Windows.Forms.TextBox()
        Me.cmdVentGeometry = New System.Windows.Forms.Button()
        Me.txtVentCloseTime = New System.Windows.Forms.TextBox()
        Me.cmdGlassProperties = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.txtHOreliability = New System.Windows.Forms.TextBox()
        Me.cmdDist_vheight = New System.Windows.Forms.Button()
        Me.cmdDist_vwidth = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtVentProb = New System.Windows.Forms.TextBox()
        Me.cmdDist_vprob = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtVentDescription = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtBalconyWidth = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtVentCD = New System.Windows.Forms.TextBox()
        Me.txtDownstand = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtOffset = New System.Windows.Forms.TextBox()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.optGlassAutoBreak = New System.Windows.Forms.RadioButton()
        Me.optGlassManualOpen = New System.Windows.Forms.RadioButton()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.opt3Dunchanbalcony = New System.Windows.Forms.RadioButton()
        Me.opt3Dchanbalcony = New System.Windows.Forms.RadioButton()
        Me.opt3Dadhered = New System.Windows.Forms.RadioButton()
        Me.opt2Dbalcony = New System.Windows.Forms.RadioButton()
        Me.opt2Dadhered = New System.Windows.Forms.RadioButton()
        Me.optStandard = New System.Windows.Forms.RadioButton()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdFireResistance = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Frame2.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label18
        '
        Me.Label18.BackColor = System.Drawing.SystemColors.Control
        Me.Label18.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label18.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label18.Location = New System.Drawing.Point(197, 15)
        Me.Label18.Name = "Label18"
        Me.Label18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label18.Size = New System.Drawing.Size(45, 19)
        Me.Label18.TabIndex = 140
        Me.Label18.Text = "face"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lstRoomFace
        '
        Me.lstRoomFace.BackColor = System.Drawing.Color.White
        Me.lstRoomFace.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstRoomFace.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstRoomFace.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstRoomFace.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstRoomFace.Items.AddRange(New Object() {"Front", "Right", "Rear", "Left"})
        Me.lstRoomFace.Location = New System.Drawing.Point(243, 12)
        Me.lstRoomFace.Name = "lstRoomFace"
        Me.lstRoomFace.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstRoomFace.Size = New System.Drawing.Size(61, 22)
        Me.lstRoomFace.TabIndex = 1
        '
        'cmdOpenOptions
        '
        Me.cmdOpenOptions.Location = New System.Drawing.Point(382, 423)
        Me.cmdOpenOptions.Name = "cmdOpenOptions"
        Me.cmdOpenOptions.Size = New System.Drawing.Size(153, 26)
        Me.cmdOpenOptions.TabIndex = 90
        Me.cmdOpenOptions.Text = "Vent Open/Close Options"
        Me.cmdOpenOptions.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(18, 15)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(97, 17)
        Me.Label8.TabIndex = 127
        Me.Label8.Text = "Connecting room"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(319, 13)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(54, 19)
        Me.Label6.TabIndex = 128
        Me.Label6.Text = "to room"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblVentOpenTime
        '
        Me.lblVentOpenTime.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.lblVentOpenTime.AutoSize = True
        Me.lblVentOpenTime.BackColor = System.Drawing.Color.Transparent
        Me.lblVentOpenTime.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVentOpenTime.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVentOpenTime.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblVentOpenTime.Location = New System.Drawing.Point(22, 158)
        Me.lblVentOpenTime.Name = "lblVentOpenTime"
        Me.lblVentOpenTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVentOpenTime.Size = New System.Drawing.Size(123, 14)
        Me.lblVentOpenTime.TabIndex = 130
        Me.lblVentOpenTime.Text = "Vent OpeningTime (sec)"
        Me.lblVentOpenTime.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblVentSillHeight
        '
        Me.lblVentSillHeight.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblVentSillHeight.BackColor = System.Drawing.SystemColors.Control
        Me.lblVentSillHeight.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVentSillHeight.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVentSillHeight.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentSillHeight.Location = New System.Drawing.Point(48, 90)
        Me.lblVentSillHeight.Name = "lblVentSillHeight"
        Me.lblVentSillHeight.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVentSillHeight.Size = New System.Drawing.Size(97, 30)
        Me.lblVentSillHeight.TabIndex = 131
        Me.lblVentSillHeight.Text = "Vent Sill Height (m)"
        Me.lblVentSillHeight.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblVentWidth
        '
        Me.lblVentWidth.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblVentWidth.BackColor = System.Drawing.SystemColors.Control
        Me.lblVentWidth.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVentWidth.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVentWidth.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentWidth.Location = New System.Drawing.Point(64, 30)
        Me.lblVentWidth.Name = "lblVentWidth"
        Me.lblVentWidth.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVentWidth.Size = New System.Drawing.Size(81, 30)
        Me.lblVentWidth.TabIndex = 133
        Me.lblVentWidth.Text = "Vent Width (m)"
        Me.lblVentWidth.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label10
        '
        Me.Label10.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label10.AutoSize = True
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(24, 188)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(121, 14)
        Me.Label10.TabIndex = 134
        Me.Label10.Text = "Vent Closing Time (sec)"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lstRoom1
        '
        Me.lstRoom1.BackColor = System.Drawing.SystemColors.Window
        Me.lstRoom1.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstRoom1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstRoom1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstRoom1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstRoom1.Location = New System.Drawing.Point(121, 12)
        Me.lstRoom1.Name = "lstRoom1"
        Me.lstRoom1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstRoom1.Size = New System.Drawing.Size(70, 22)
        Me.lstRoom1.TabIndex = 0
        Me.lstRoom1.Tag = "  "
        '
        'lstRoom2
        '
        Me.lstRoom2.BackColor = System.Drawing.SystemColors.Window
        Me.lstRoom2.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstRoom2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstRoom2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstRoom2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstRoom2.Location = New System.Drawing.Point(379, 12)
        Me.lstRoom2.Name = "lstRoom2"
        Me.lstRoom2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstRoom2.Size = New System.Drawing.Size(67, 22)
        Me.lstRoom2.TabIndex = 2
        '
        'txtVentOpenTime
        '
        Me.txtVentOpenTime.AcceptsReturn = True
        Me.txtVentOpenTime.BackColor = System.Drawing.Color.White
        Me.txtVentOpenTime.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtVentOpenTime.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVentOpenTime.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVentOpenTime.Location = New System.Drawing.Point(151, 153)
        Me.txtVentOpenTime.MaxLength = 0
        Me.txtVentOpenTime.Name = "txtVentOpenTime"
        Me.txtVentOpenTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtVentOpenTime.Size = New System.Drawing.Size(49, 20)
        Me.txtVentOpenTime.TabIndex = 8
        Me.txtVentOpenTime.Text = "0"
        '
        'txtVentHeight
        '
        Me.txtVentHeight.AcceptsReturn = True
        Me.txtVentHeight.BackColor = System.Drawing.Color.White
        Me.txtVentHeight.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtVentHeight.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVentHeight.ForeColor = System.Drawing.Color.Black
        Me.txtVentHeight.Location = New System.Drawing.Point(151, 63)
        Me.txtVentHeight.MaxLength = 0
        Me.txtVentHeight.Name = "txtVentHeight"
        Me.txtVentHeight.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtVentHeight.Size = New System.Drawing.Size(49, 20)
        Me.txtVentHeight.TabIndex = 5
        Me.txtVentHeight.Text = "1"
        '
        'txtVentWidth
        '
        Me.txtVentWidth.AcceptsReturn = True
        Me.txtVentWidth.BackColor = System.Drawing.Color.White
        Me.txtVentWidth.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtVentWidth.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVentWidth.ForeColor = System.Drawing.Color.Black
        Me.txtVentWidth.Location = New System.Drawing.Point(151, 33)
        Me.txtVentWidth.MaxLength = 0
        Me.txtVentWidth.Name = "txtVentWidth"
        Me.txtVentWidth.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtVentWidth.Size = New System.Drawing.Size(49, 20)
        Me.txtVentWidth.TabIndex = 4
        Me.txtVentWidth.Text = "1"
        '
        'txtVentSillHeight
        '
        Me.txtVentSillHeight.AcceptsReturn = True
        Me.txtVentSillHeight.BackColor = System.Drawing.Color.White
        Me.txtVentSillHeight.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtVentSillHeight.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVentSillHeight.ForeColor = System.Drawing.Color.Black
        Me.txtVentSillHeight.Location = New System.Drawing.Point(151, 93)
        Me.txtVentSillHeight.MaxLength = 0
        Me.txtVentSillHeight.Name = "txtVentSillHeight"
        Me.txtVentSillHeight.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtVentSillHeight.Size = New System.Drawing.Size(49, 20)
        Me.txtVentSillHeight.TabIndex = 6
        Me.txtVentSillHeight.Text = "0"
        '
        'cmdVentGeometry
        '
        Me.cmdVentGeometry.BackColor = System.Drawing.SystemColors.Control
        Me.cmdVentGeometry.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdVentGeometry.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdVentGeometry.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdVentGeometry.Location = New System.Drawing.Point(21, 426)
        Me.cmdVentGeometry.Name = "cmdVentGeometry"
        Me.cmdVentGeometry.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdVentGeometry.Size = New System.Drawing.Size(122, 25)
        Me.cmdVentGeometry.TabIndex = 124
        Me.cmdVentGeometry.TabStop = False
        Me.cmdVentGeometry.Text = "Room-Vent Geometry"
        Me.cmdVentGeometry.UseVisualStyleBackColor = False
        Me.cmdVentGeometry.Visible = False
        '
        'txtVentCloseTime
        '
        Me.txtVentCloseTime.AcceptsReturn = True
        Me.txtVentCloseTime.BackColor = System.Drawing.SystemColors.Window
        Me.txtVentCloseTime.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtVentCloseTime.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVentCloseTime.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVentCloseTime.Location = New System.Drawing.Point(151, 183)
        Me.txtVentCloseTime.MaxLength = 0
        Me.txtVentCloseTime.Name = "txtVentCloseTime"
        Me.txtVentCloseTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtVentCloseTime.Size = New System.Drawing.Size(49, 20)
        Me.txtVentCloseTime.TabIndex = 9
        Me.txtVentCloseTime.Text = "0"
        '
        'cmdGlassProperties
        '
        Me.cmdGlassProperties.BackColor = System.Drawing.SystemColors.Control
        Me.cmdGlassProperties.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdGlassProperties.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdGlassProperties.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdGlassProperties.Location = New System.Drawing.Point(16, 89)
        Me.cmdGlassProperties.Name = "cmdGlassProperties"
        Me.cmdGlassProperties.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdGlassProperties.Size = New System.Drawing.Size(144, 25)
        Me.cmdGlassProperties.TabIndex = 80
        Me.cmdGlassProperties.Text = "Vent - Glass Properties"
        Me.cmdGlassProperties.UseVisualStyleBackColor = False
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 3
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 119.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Label11, 0, 11)
        Me.TableLayoutPanel1.Controls.Add(Me.Button1, 2, 11)
        Me.TableLayoutPanel1.Controls.Add(Me.txtHOreliability, 1, 11)
        Me.TableLayoutPanel1.Controls.Add(Me.cmdDist_vheight, 2, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.cmdDist_vwidth, 2, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.txtVentWidth, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.lblVentWidth, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.txtVentHeight, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.txtVentSillHeight, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label10, 0, 6)
        Me.TableLayoutPanel1.Controls.Add(Me.lblVentOpenTime, 0, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.txtVentOpenTime, 1, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.txtVentCloseTime, 1, 6)
        Me.TableLayoutPanel1.Controls.Add(Me.lblVentSillHeight, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label7, 0, 10)
        Me.TableLayoutPanel1.Controls.Add(Me.txtVentProb, 1, 10)
        Me.TableLayoutPanel1.Controls.Add(Me.cmdDist_vprob, 2, 10)
        Me.TableLayoutPanel1.Controls.Add(Me.Label2, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.txtVentDescription, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label4, 0, 8)
        Me.TableLayoutPanel1.Controls.Add(Me.txtBalconyWidth, 1, 8)
        Me.TableLayoutPanel1.Controls.Add(Me.Label5, 0, 7)
        Me.TableLayoutPanel1.Controls.Add(Me.txtVentCD, 1, 7)
        Me.TableLayoutPanel1.Controls.Add(Me.txtDownstand, 1, 9)
        Me.TableLayoutPanel1.Controls.Add(Me.Label9, 0, 9)
        Me.TableLayoutPanel1.Controls.Add(Me.Label3, 0, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.txtOffset, 1, 4)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(21, 49)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 15
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(303, 368)
        Me.TableLayoutPanel1.TabIndex = 141
        '
        'Label11
        '
        Me.Label11.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(3, 338)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(142, 13)
        Me.Label11.TabIndex = 199
        Me.Label11.Text = "Hold Open Device Reliability"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(206, 333)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(91, 22)
        Me.Button1.TabIndex = 50
        Me.Button1.Text = "distribution"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'txtHOreliability
        '
        Me.txtHOreliability.Location = New System.Drawing.Point(151, 333)
        Me.txtHOreliability.Name = "txtHOreliability"
        Me.txtHOreliability.Size = New System.Drawing.Size(49, 20)
        Me.txtHOreliability.TabIndex = 14
        Me.txtHOreliability.Text = "1"
        '
        'cmdDist_vheight
        '
        Me.cmdDist_vheight.Location = New System.Drawing.Point(206, 63)
        Me.cmdDist_vheight.Name = "cmdDist_vheight"
        Me.cmdDist_vheight.Size = New System.Drawing.Size(93, 22)
        Me.cmdDist_vheight.TabIndex = 30
        Me.cmdDist_vheight.Text = "distribution"
        Me.cmdDist_vheight.UseVisualStyleBackColor = True
        '
        'cmdDist_vwidth
        '
        Me.cmdDist_vwidth.Location = New System.Drawing.Point(206, 33)
        Me.cmdDist_vwidth.Name = "cmdDist_vwidth"
        Me.cmdDist_vwidth.Size = New System.Drawing.Size(93, 22)
        Me.cmdDist_vwidth.TabIndex = 20
        Me.cmdDist_vwidth.Text = "distribution"
        Me.cmdDist_vwidth.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(64, 60)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(81, 30)
        Me.Label1.TabIndex = 135
        Me.Label1.Text = "Vent Height (m)"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label7
        '
        Me.Label7.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(19, 308)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(126, 13)
        Me.Label7.TabIndex = 198
        Me.Label7.Text = "Prob. Vent Initially Closed"
        '
        'txtVentProb
        '
        Me.txtVentProb.Location = New System.Drawing.Point(151, 303)
        Me.txtVentProb.Name = "txtVentProb"
        Me.txtVentProb.Size = New System.Drawing.Size(49, 20)
        Me.txtVentProb.TabIndex = 13
        Me.txtVentProb.Text = "0"
        '
        'cmdDist_vprob
        '
        Me.cmdDist_vprob.Location = New System.Drawing.Point(206, 303)
        Me.cmdDist_vprob.Name = "cmdDist_vprob"
        Me.cmdDist_vprob.Size = New System.Drawing.Size(91, 22)
        Me.cmdDist_vprob.TabIndex = 40
        Me.cmdDist_vprob.Text = "distribution"
        Me.cmdDist_vprob.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(84, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(61, 14)
        Me.Label2.TabIndex = 136
        Me.Label2.Text = "Description"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtVentDescription
        '
        Me.txtVentDescription.AcceptsReturn = True
        Me.txtVentDescription.BackColor = System.Drawing.SystemColors.Window
        Me.TableLayoutPanel1.SetColumnSpan(Me.txtVentDescription, 2)
        Me.txtVentDescription.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtVentDescription.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVentDescription.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVentDescription.Location = New System.Drawing.Point(151, 3)
        Me.txtVentDescription.MaxLength = 23
        Me.txtVentDescription.Name = "txtVentDescription"
        Me.txtVentDescription.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtVentDescription.Size = New System.Drawing.Size(151, 20)
        Me.txtVentDescription.TabIndex = 3
        Me.txtVentDescription.Text = "vent label"
        '
        'Label4
        '
        Me.Label4.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(29, 248)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(116, 13)
        Me.Label4.TabIndex = 203
        Me.Label4.Text = "Projection Distance (m)"
        '
        'txtBalconyWidth
        '
        Me.txtBalconyWidth.Location = New System.Drawing.Point(151, 243)
        Me.txtBalconyWidth.Name = "txtBalconyWidth"
        Me.txtBalconyWidth.Size = New System.Drawing.Size(49, 20)
        Me.txtBalconyWidth.TabIndex = 11
        Me.txtBalconyWidth.Tag = "0"
        Me.txtBalconyWidth.Text = "0"
        Me.ToolTip1.SetToolTip(Me.txtBalconyWidth, "The horizontal distance between the compartment opening and the spill edge of the" &
        " balcony or soffit.")
        '
        'Label5
        '
        Me.Label5.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(47, 218)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(98, 13)
        Me.Label5.TabIndex = 197
        Me.Label5.Text = "Discharge Coeff. (-)"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtVentCD
        '
        Me.txtVentCD.Location = New System.Drawing.Point(151, 213)
        Me.txtVentCD.Name = "txtVentCD"
        Me.txtVentCD.Size = New System.Drawing.Size(49, 20)
        Me.txtVentCD.TabIndex = 10
        Me.txtVentCD.Text = "0.68"
        '
        'txtDownstand
        '
        Me.txtDownstand.Location = New System.Drawing.Point(151, 273)
        Me.txtDownstand.Name = "txtDownstand"
        Me.txtDownstand.Size = New System.Drawing.Size(49, 20)
        Me.txtDownstand.TabIndex = 12
        Me.txtDownstand.Text = "0"
        Me.txtDownstand.Visible = False
        '
        'Label9
        '
        Me.Label9.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(67, 278)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(78, 13)
        Me.Label9.TabIndex = 206
        Me.Label9.Text = "Downstand (m)"
        Me.Label9.Visible = False
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 128)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(137, 13)
        Me.Label3.TabIndex = 208
        Me.Label3.Text = "Horizontal Offset in Wall (m)"
        '
        'txtOffset
        '
        Me.txtOffset.Location = New System.Drawing.Point(151, 123)
        Me.txtOffset.Name = "txtOffset"
        Me.txtOffset.Size = New System.Drawing.Size(49, 20)
        Me.txtOffset.TabIndex = 7
        Me.txtOffset.Text = "0"
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(541, 426)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(75, 23)
        Me.cmdSave.TabIndex = 91
        Me.cmdSave.Text = "Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.optGlassAutoBreak)
        Me.Frame2.Controls.Add(Me.cmdGlassProperties)
        Me.Frame2.Controls.Add(Me.optGlassManualOpen)
        Me.Frame2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(333, 288)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(280, 129)
        Me.Frame2.TabIndex = 143
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Glass Fracture Model"
        '
        'optGlassAutoBreak
        '
        Me.optGlassAutoBreak.BackColor = System.Drawing.SystemColors.Control
        Me.optGlassAutoBreak.Cursor = System.Windows.Forms.Cursors.Default
        Me.optGlassAutoBreak.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optGlassAutoBreak.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optGlassAutoBreak.Location = New System.Drawing.Point(16, 51)
        Me.optGlassAutoBreak.Name = "optGlassAutoBreak"
        Me.optGlassAutoBreak.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optGlassAutoBreak.Size = New System.Drawing.Size(267, 26)
        Me.optGlassAutoBreak.TabIndex = 71
        Me.optGlassAutoBreak.TabStop = True
        Me.optGlassAutoBreak.Text = "Auto Break Glass  (use glass fracture model)"
        Me.optGlassAutoBreak.UseVisualStyleBackColor = False
        '
        'optGlassManualOpen
        '
        Me.optGlassManualOpen.BackColor = System.Drawing.SystemColors.Control
        Me.optGlassManualOpen.Checked = True
        Me.optGlassManualOpen.Cursor = System.Windows.Forms.Cursors.Default
        Me.optGlassManualOpen.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optGlassManualOpen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optGlassManualOpen.Location = New System.Drawing.Point(16, 16)
        Me.optGlassManualOpen.Name = "optGlassManualOpen"
        Me.optGlassManualOpen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optGlassManualOpen.Size = New System.Drawing.Size(201, 33)
        Me.optGlassManualOpen.TabIndex = 70
        Me.optGlassManualOpen.TabStop = True
        Me.optGlassManualOpen.Text = "Use Manual Vent Opening Settings"
        Me.optGlassManualOpen.UseVisualStyleBackColor = False
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.opt3Dunchanbalcony)
        Me.GroupBox1.Controls.Add(Me.opt3Dchanbalcony)
        Me.GroupBox1.Controls.Add(Me.opt3Dadhered)
        Me.GroupBox1.Controls.Add(Me.opt2Dbalcony)
        Me.GroupBox1.Controls.Add(Me.opt2Dadhered)
        Me.GroupBox1.Controls.Add(Me.optStandard)
        Me.GroupBox1.Location = New System.Drawing.Point(333, 52)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(279, 228)
        Me.GroupBox1.TabIndex = 144
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Vent Type"
        '
        'opt3Dunchanbalcony
        '
        Me.opt3Dunchanbalcony.AutoSize = True
        Me.opt3Dunchanbalcony.Location = New System.Drawing.Point(16, 185)
        Me.opt3Dunchanbalcony.Name = "opt3Dunchanbalcony"
        Me.opt3Dunchanbalcony.Size = New System.Drawing.Size(181, 17)
        Me.opt3Dunchanbalcony.TabIndex = 65
        Me.opt3Dunchanbalcony.Text = "3D Unchannelled Balcony Plume"
        Me.opt3Dunchanbalcony.UseVisualStyleBackColor = True
        '
        'opt3Dchanbalcony
        '
        Me.opt3Dchanbalcony.AutoSize = True
        Me.opt3Dchanbalcony.Location = New System.Drawing.Point(16, 155)
        Me.opt3Dchanbalcony.Name = "opt3Dchanbalcony"
        Me.opt3Dchanbalcony.Size = New System.Drawing.Size(168, 17)
        Me.opt3Dchanbalcony.TabIndex = 64
        Me.opt3Dchanbalcony.Text = "3D Channelled Balcony Plume"
        Me.opt3Dchanbalcony.UseVisualStyleBackColor = True
        '
        'opt3Dadhered
        '
        Me.opt3Dadhered.AutoSize = True
        Me.opt3Dadhered.Location = New System.Drawing.Point(16, 93)
        Me.opt3Dadhered.Name = "opt3Dadhered"
        Me.opt3Dadhered.Size = New System.Drawing.Size(114, 17)
        Me.opt3Dadhered.TabIndex = 62
        Me.opt3Dadhered.Text = "3D Adhered Plume"
        Me.opt3Dadhered.UseVisualStyleBackColor = True
        '
        'opt2Dbalcony
        '
        Me.opt2Dbalcony.AutoSize = True
        Me.opt2Dbalcony.Location = New System.Drawing.Point(16, 125)
        Me.opt2Dbalcony.Name = "opt2Dbalcony"
        Me.opt2Dbalcony.Size = New System.Drawing.Size(112, 17)
        Me.opt2Dbalcony.TabIndex = 63
        Me.opt2Dbalcony.Text = "2D Balcony Plume"
        Me.opt2Dbalcony.UseVisualStyleBackColor = True
        '
        'opt2Dadhered
        '
        Me.opt2Dadhered.AutoSize = True
        Me.opt2Dadhered.Location = New System.Drawing.Point(16, 65)
        Me.opt2Dadhered.Name = "opt2Dadhered"
        Me.opt2Dadhered.Size = New System.Drawing.Size(114, 17)
        Me.opt2Dadhered.TabIndex = 61
        Me.opt2Dadhered.Text = "2D Adhered Plume"
        Me.opt2Dadhered.UseVisualStyleBackColor = True
        '
        'optStandard
        '
        Me.optStandard.AutoSize = True
        Me.optStandard.Checked = True
        Me.optStandard.Location = New System.Drawing.Point(16, 33)
        Me.optStandard.Name = "optStandard"
        Me.optStandard.Size = New System.Drawing.Size(68, 17)
        Me.optStandard.TabIndex = 60
        Me.optStandard.TabStop = True
        Me.optStandard.Text = "Standard"
        Me.optStandard.UseVisualStyleBackColor = True
        '
        'cmdFireResistance
        '
        Me.cmdFireResistance.Location = New System.Drawing.Point(464, 11)
        Me.cmdFireResistance.Name = "cmdFireResistance"
        Me.cmdFireResistance.Size = New System.Drawing.Size(152, 23)
        Me.cmdFireResistance.TabIndex = 145
        Me.cmdFireResistance.TabStop = False
        Me.cmdFireResistance.Text = "Fire Resistance Settings"
        Me.cmdFireResistance.UseVisualStyleBackColor = True
        '
        'frmNewVents
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(637, 460)
        Me.Controls.Add(Me.cmdFireResistance)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.lstRoomFace)
        Me.Controls.Add(Me.cmdOpenOptions)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.lstRoom1)
        Me.Controls.Add(Me.lstRoom2)
        Me.Controls.Add(Me.cmdVentGeometry)
        Me.MaximizeBox = False
        Me.Name = "frmNewVents"
        Me.ShowInTaskbar = False
        Me.Text = "New Vent"
        Me.TopMost = True
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.Frame2.ResumeLayout(False)
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents Label18 As System.Windows.Forms.Label
    Public WithEvents lstRoomFace As System.Windows.Forms.ComboBox
    Friend WithEvents cmdOpenOptions As System.Windows.Forms.Button
    Public WithEvents Label8 As System.Windows.Forms.Label
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents lblVentOpenTime As System.Windows.Forms.Label
    Public WithEvents lblVentSillHeight As System.Windows.Forms.Label
    Public WithEvents lblVentWidth As System.Windows.Forms.Label
    Public WithEvents Label10 As System.Windows.Forms.Label
    Public WithEvents lstRoom1 As System.Windows.Forms.ComboBox
    Public WithEvents lstRoom2 As System.Windows.Forms.ComboBox
    Public WithEvents txtVentOpenTime As System.Windows.Forms.TextBox
    Public WithEvents txtVentHeight As System.Windows.Forms.TextBox
    Public WithEvents txtVentWidth As System.Windows.Forms.TextBox
    Public WithEvents txtVentSillHeight As System.Windows.Forms.TextBox
    Public WithEvents cmdVentGeometry As System.Windows.Forms.Button
    Public WithEvents txtVentCloseTime As System.Windows.Forms.TextBox
    Public WithEvents cmdGlassProperties As System.Windows.Forms.Button
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents txtVentDescription As System.Windows.Forms.TextBox
    Public WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdDist_vheight As System.Windows.Forms.Button
    Friend WithEvents cmdDist_vwidth As System.Windows.Forms.Button
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents optGlassAutoBreak As System.Windows.Forms.RadioButton
    Public WithEvents optGlassManualOpen As System.Windows.Forms.RadioButton
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtVentCD As System.Windows.Forms.TextBox
    Friend WithEvents txtVentProb As System.Windows.Forms.TextBox
    Friend WithEvents cmdDist_vprob As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents opt2Dadhered As System.Windows.Forms.RadioButton
    Friend WithEvents optStandard As System.Windows.Forms.RadioButton
    Friend WithEvents opt2Dbalcony As System.Windows.Forms.RadioButton
    Friend WithEvents opt3Dadhered As System.Windows.Forms.RadioButton
    Friend WithEvents opt3Dchanbalcony As System.Windows.Forms.RadioButton
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtBalconyWidth As System.Windows.Forms.TextBox
    Friend WithEvents txtDownstand As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents opt3Dunchanbalcony As System.Windows.Forms.RadioButton
    Friend WithEvents txtOffset As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents txtHOreliability As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents cmdFireResistance As System.Windows.Forms.Button
End Class
