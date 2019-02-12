<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNewRoom
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.txtID = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.cmdDist_width = New System.Windows.Forms.Button()
        Me.cmdDist_length = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtWidth = New System.Windows.Forms.TextBox()
        Me.txtElevation = New System.Windows.Forms.TextBox()
        Me.txtLength = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextAbsX = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lblLength = New System.Windows.Forms.Label()
        Me.txtMaxHeight = New System.Windows.Forms.TextBox()
        Me.txtMinHeight = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.Label36 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.TextAbsY = New System.Windows.Forms.TextBox()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtWallSubstrateThickness = New System.Windows.Forms.TextBox()
        Me.lblWallSubstrateThickness = New System.Windows.Forms.Label()
        Me.txtWallSurfaceThickness = New System.Windows.Forms.TextBox()
        Me.lblWallSurfaceThickness = New System.Windows.Forms.Label()
        Me.cmdPickWSubstrate = New System.Windows.Forms.Button()
        Me.cmdPickWsurface = New System.Windows.Forms.Button()
        Me.lblWallSubstrate = New System.Windows.Forms.Label()
        Me.lblWallSurface = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.txtCeilingSubstrateThickness = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txtCeilingSurfaceThickness = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.lblCeilingSubstrate = New System.Windows.Forms.Label()
        Me.lblCeilingSurface = New System.Windows.Forms.Label()
        Me.cmdPickCsubstrate = New System.Windows.Forms.Button()
        Me.cmdPickCsurface = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.txtFloorSubstrateThickness = New System.Windows.Forms.TextBox()
        Me.cmdPickFsubstrate = New System.Windows.Forms.Button()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.cmdPickFsurface = New System.Windows.Forms.Button()
        Me.txtFloorSurfaceThickness = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.lblFloorSurface = New System.Windows.Forms.Label()
        Me.lblFloorSubstrate = New System.Windows.Forms.Label()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.AutoSize = True
        Me.TableLayoutPanel1.ColumnCount = 4
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 114.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.txtID, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label7, 2, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.Label3, 0, 10)
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label10, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.cmdDist_width, 3, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.cmdDist_length, 3, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label6, 0, 8)
        Me.TableLayoutPanel1.Controls.Add(Me.txtWidth, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.txtElevation, 1, 6)
        Me.TableLayoutPanel1.Controls.Add(Me.txtLength, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label5, 0, 7)
        Me.TableLayoutPanel1.Controls.Add(Me.TextAbsX, 1, 7)
        Me.TableLayoutPanel1.Controls.Add(Me.Label8, 0, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.lblLength, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.txtMaxHeight, 1, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.txtMinHeight, 1, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.Label4, 0, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.Label31, 2, 6)
        Me.TableLayoutPanel1.Controls.Add(Me.Label2, 2, 7)
        Me.TableLayoutPanel1.Controls.Add(Me.Label11, 2, 8)
        Me.TableLayoutPanel1.Controls.Add(Me.Label33, 2, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label36, 2, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label13, 2, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label30, 0, 6)
        Me.TableLayoutPanel1.Controls.Add(Me.Label9, 2, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.txtDescription, 1, 10)
        Me.TableLayoutPanel1.Controls.Add(Me.TextAbsY, 1, 8)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(12, 12)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 13
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(328, 299)
        Me.TableLayoutPanel1.TabIndex = 23
        '
        'txtID
        '
        Me.txtID.Location = New System.Drawing.Point(107, 23)
        Me.txtID.Name = "txtID"
        Me.txtID.ReadOnly = True
        Me.txtID.Size = New System.Drawing.Size(62, 20)
        Me.txtID.TabIndex = 203
        Me.txtID.TabStop = False
        Me.txtID.Tag = "room"
        Me.txtID.Text = "1"
        '
        'Label7
        '
        Me.Label7.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(175, 124)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(36, 26)
        Me.Label7.TabIndex = 103
        Me.Label7.Text = "m"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(38, 234)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(63, 13)
        Me.Label3.TabIndex = 201
        Me.Label3.Text = "Description "
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(66, 78)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(35, 13)
        Me.Label1.TabIndex = 21
        Me.Label1.Text = "Width"
        '
        'Label10
        '
        Me.Label10.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(71, 26)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(30, 13)
        Me.Label10.TabIndex = 18
        Me.Label10.Text = "room"
        '
        'cmdDist_width
        '
        Me.cmdDist_width.Location = New System.Drawing.Point(217, 75)
        Me.cmdDist_width.Name = "cmdDist_width"
        Me.cmdDist_width.Size = New System.Drawing.Size(68, 19)
        Me.cmdDist_width.TabIndex = 11
        Me.cmdDist_width.Text = "distribution"
        Me.cmdDist_width.UseVisualStyleBackColor = True
        '
        'cmdDist_length
        '
        Me.cmdDist_length.Location = New System.Drawing.Point(217, 49)
        Me.cmdDist_length.Name = "cmdDist_length"
        Me.cmdDist_length.Size = New System.Drawing.Size(68, 19)
        Me.cmdDist_length.TabIndex = 10
        Me.cmdDist_length.Text = "distribution"
        Me.cmdDist_length.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 208)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(98, 13)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "abs y - coordinate *"
        '
        'txtWidth
        '
        Me.txtWidth.Location = New System.Drawing.Point(107, 75)
        Me.txtWidth.Name = "txtWidth"
        Me.txtWidth.Size = New System.Drawing.Size(62, 20)
        Me.txtWidth.TabIndex = 2
        Me.txtWidth.Tag = "width"
        Me.txtWidth.Text = "2.4"
        '
        'txtElevation
        '
        Me.txtElevation.Location = New System.Drawing.Point(107, 153)
        Me.txtElevation.Name = "txtElevation"
        Me.txtElevation.Size = New System.Drawing.Size(62, 20)
        Me.txtElevation.TabIndex = 5
        Me.txtElevation.Tag = "elevation"
        Me.txtElevation.Text = "0"
        '
        'txtLength
        '
        Me.txtLength.Location = New System.Drawing.Point(107, 49)
        Me.txtLength.Name = "txtLength"
        Me.txtLength.Size = New System.Drawing.Size(62, 20)
        Me.txtLength.TabIndex = 1
        Me.txtLength.Tag = "length"
        Me.txtLength.Text = "3.6"
        '
        'Label5
        '
        Me.Label5.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 182)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(98, 13)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "abs x - coordinate *"
        '
        'TextAbsX
        '
        Me.TextAbsX.Location = New System.Drawing.Point(107, 179)
        Me.TextAbsX.Name = "TextAbsX"
        Me.TextAbsX.Size = New System.Drawing.Size(62, 20)
        Me.TextAbsX.TabIndex = 6
        Me.TextAbsX.Tag = "absx"
        Me.TextAbsX.Text = "0"
        '
        'Label8
        '
        Me.Label8.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(16, 130)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(85, 13)
        Me.Label8.TabIndex = 13
        Me.Label8.Text = "Maximum Height"
        '
        'lblLength
        '
        Me.lblLength.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.lblLength.AutoSize = True
        Me.lblLength.Location = New System.Drawing.Point(61, 52)
        Me.lblLength.Name = "lblLength"
        Me.lblLength.Size = New System.Drawing.Size(40, 13)
        Me.lblLength.TabIndex = 3
        Me.lblLength.Text = "Length"
        '
        'txtMaxHeight
        '
        Me.txtMaxHeight.Location = New System.Drawing.Point(107, 127)
        Me.txtMaxHeight.Name = "txtMaxHeight"
        Me.txtMaxHeight.Size = New System.Drawing.Size(62, 20)
        Me.txtMaxHeight.TabIndex = 4
        Me.txtMaxHeight.Tag = "maxheight"
        Me.txtMaxHeight.Text = "2.4"
        '
        'txtMinHeight
        '
        Me.txtMinHeight.Location = New System.Drawing.Point(107, 101)
        Me.txtMinHeight.Name = "txtMinHeight"
        Me.txtMinHeight.Size = New System.Drawing.Size(62, 20)
        Me.txtMinHeight.TabIndex = 3
        Me.txtMinHeight.Tag = "minheight"
        Me.txtMinHeight.Text = "2.4"
        '
        'Label4
        '
        Me.Label4.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(19, 104)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(82, 13)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Minimum Height"
        '
        'Label31
        '
        Me.Label31.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label31.AutoSize = True
        Me.Label31.Location = New System.Drawing.Point(175, 150)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(36, 26)
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
        Me.Label2.Location = New System.Drawing.Point(175, 176)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(36, 26)
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
        Me.Label11.Location = New System.Drawing.Point(175, 202)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(36, 26)
        Me.Label11.TabIndex = 81
        Me.Label11.Text = "m"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label33
        '
        Me.Label33.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label33.AutoSize = True
        Me.Label33.Location = New System.Drawing.Point(175, 46)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(36, 26)
        Me.Label33.TabIndex = 87
        Me.Label33.Text = "m"
        Me.Label33.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label36
        '
        Me.Label36.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label36.AutoSize = True
        Me.Label36.Location = New System.Drawing.Point(175, 72)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(36, 26)
        Me.Label36.TabIndex = 102
        Me.Label36.Text = "m"
        Me.Label36.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label13
        '
        Me.Label13.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(175, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(36, 20)
        Me.Label13.TabIndex = 103
        Me.Label13.Text = "Units"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label30
        '
        Me.Label30.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label30.AutoSize = True
        Me.Label30.Location = New System.Drawing.Point(50, 156)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(51, 13)
        Me.Label30.TabIndex = 139
        Me.Label30.Text = "Elevation"
        Me.Label30.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label9
        '
        Me.Label9.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(175, 104)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(15, 13)
        Me.Label9.TabIndex = 200
        Me.Label9.Text = "m"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtDescription
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.txtDescription, 3)
        Me.txtDescription.Location = New System.Drawing.Point(107, 231)
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(205, 20)
        Me.txtDescription.TabIndex = 8
        Me.txtDescription.Tag = "description"
        Me.txtDescription.Text = "Room A"
        '
        'TextAbsY
        '
        Me.TextAbsY.Location = New System.Drawing.Point(107, 205)
        Me.TextAbsY.Name = "TextAbsY"
        Me.TextAbsY.Size = New System.Drawing.Size(62, 20)
        Me.TextAbsY.TabIndex = 7
        Me.TextAbsY.Tag = "absy"
        Me.TextAbsY.Text = "0"
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(667, 324)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(75, 23)
        Me.cmdSave.TabIndex = 9
        Me.cmdSave.Text = "Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(586, 324)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 12
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(22, 322)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(130, 13)
        Me.Label12.TabIndex = 24
        Me.Label12.Text = "* no effect on calculations"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtWallSubstrateThickness)
        Me.GroupBox1.Controls.Add(Me.lblWallSubstrateThickness)
        Me.GroupBox1.Controls.Add(Me.txtWallSurfaceThickness)
        Me.GroupBox1.Controls.Add(Me.lblWallSurfaceThickness)
        Me.GroupBox1.Controls.Add(Me.cmdPickWSubstrate)
        Me.GroupBox1.Controls.Add(Me.cmdPickWsurface)
        Me.GroupBox1.Controls.Add(Me.lblWallSubstrate)
        Me.GroupBox1.Controls.Add(Me.lblWallSurface)
        Me.GroupBox1.Location = New System.Drawing.Point(346, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(397, 99)
        Me.GroupBox1.TabIndex = 25
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Wall Materials"
        '
        'txtWallSubstrateThickness
        '
        Me.txtWallSubstrateThickness.Location = New System.Drawing.Point(208, 56)
        Me.txtWallSubstrateThickness.Name = "txtWallSubstrateThickness"
        Me.txtWallSubstrateThickness.Size = New System.Drawing.Size(35, 20)
        Me.txtWallSubstrateThickness.TabIndex = 7
        Me.txtWallSubstrateThickness.Text = "150"
        '
        'lblWallSubstrateThickness
        '
        Me.lblWallSubstrateThickness.AutoSize = True
        Me.lblWallSubstrateThickness.Location = New System.Drawing.Point(121, 60)
        Me.lblWallSubstrateThickness.Name = "lblWallSubstrateThickness"
        Me.lblWallSubstrateThickness.Size = New System.Drawing.Size(81, 13)
        Me.lblWallSubstrateThickness.TabIndex = 6
        Me.lblWallSubstrateThickness.Text = "Thickness (mm)"
        '
        'txtWallSurfaceThickness
        '
        Me.txtWallSurfaceThickness.Location = New System.Drawing.Point(208, 27)
        Me.txtWallSurfaceThickness.Name = "txtWallSurfaceThickness"
        Me.txtWallSurfaceThickness.Size = New System.Drawing.Size(35, 20)
        Me.txtWallSurfaceThickness.TabIndex = 5
        Me.txtWallSurfaceThickness.Text = "13"
        '
        'lblWallSurfaceThickness
        '
        Me.lblWallSurfaceThickness.AutoSize = True
        Me.lblWallSurfaceThickness.Location = New System.Drawing.Point(121, 30)
        Me.lblWallSurfaceThickness.Name = "lblWallSurfaceThickness"
        Me.lblWallSurfaceThickness.Size = New System.Drawing.Size(81, 13)
        Me.lblWallSurfaceThickness.TabIndex = 4
        Me.lblWallSurfaceThickness.Text = "Thickness (mm)"
        '
        'cmdPickWSubstrate
        '
        Me.cmdPickWSubstrate.Location = New System.Drawing.Point(11, 58)
        Me.cmdPickWSubstrate.Name = "cmdPickWSubstrate"
        Me.cmdPickWSubstrate.Size = New System.Drawing.Size(104, 23)
        Me.cmdPickWSubstrate.TabIndex = 3
        Me.cmdPickWSubstrate.Text = "Pick substrate mat"
        Me.cmdPickWSubstrate.UseVisualStyleBackColor = True
        '
        'cmdPickWsurface
        '
        Me.cmdPickWsurface.Location = New System.Drawing.Point(11, 29)
        Me.cmdPickWsurface.Name = "cmdPickWsurface"
        Me.cmdPickWsurface.Size = New System.Drawing.Size(104, 23)
        Me.cmdPickWsurface.TabIndex = 2
        Me.cmdPickWsurface.Text = "Pick surface mat"
        Me.cmdPickWsurface.UseVisualStyleBackColor = True
        '
        'lblWallSubstrate
        '
        Me.lblWallSubstrate.Location = New System.Drawing.Point(249, 60)
        Me.lblWallSubstrate.Name = "lblWallSubstrate"
        Me.lblWallSubstrate.Size = New System.Drawing.Size(142, 21)
        Me.lblWallSubstrate.TabIndex = 1
        Me.lblWallSubstrate.Text = "None"
        '
        'lblWallSurface
        '
        Me.lblWallSurface.Location = New System.Drawing.Point(249, 30)
        Me.lblWallSurface.Name = "lblWallSurface"
        Me.lblWallSurface.Size = New System.Drawing.Size(142, 16)
        Me.lblWallSurface.TabIndex = 0
        Me.lblWallSurface.Text = "None"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtCeilingSubstrateThickness)
        Me.GroupBox2.Controls.Add(Me.Label14)
        Me.GroupBox2.Controls.Add(Me.txtCeilingSurfaceThickness)
        Me.GroupBox2.Controls.Add(Me.Label15)
        Me.GroupBox2.Controls.Add(Me.lblCeilingSubstrate)
        Me.GroupBox2.Controls.Add(Me.lblCeilingSurface)
        Me.GroupBox2.Controls.Add(Me.cmdPickCsubstrate)
        Me.GroupBox2.Controls.Add(Me.cmdPickCsurface)
        Me.GroupBox2.Location = New System.Drawing.Point(346, 113)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(397, 99)
        Me.GroupBox2.TabIndex = 26
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Ceiling Materials"
        '
        'txtCeilingSubstrateThickness
        '
        Me.txtCeilingSubstrateThickness.Location = New System.Drawing.Point(208, 52)
        Me.txtCeilingSubstrateThickness.Name = "txtCeilingSubstrateThickness"
        Me.txtCeilingSubstrateThickness.Size = New System.Drawing.Size(35, 20)
        Me.txtCeilingSubstrateThickness.TabIndex = 14
        Me.txtCeilingSubstrateThickness.Text = "150"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(121, 56)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(81, 13)
        Me.Label14.TabIndex = 13
        Me.Label14.Text = "Thickness (mm)"
        '
        'txtCeilingSurfaceThickness
        '
        Me.txtCeilingSurfaceThickness.Location = New System.Drawing.Point(208, 23)
        Me.txtCeilingSurfaceThickness.Name = "txtCeilingSurfaceThickness"
        Me.txtCeilingSurfaceThickness.Size = New System.Drawing.Size(35, 20)
        Me.txtCeilingSurfaceThickness.TabIndex = 12
        Me.txtCeilingSurfaceThickness.Text = "13"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(121, 26)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(81, 13)
        Me.Label15.TabIndex = 11
        Me.Label15.Text = "Thickness (mm)"
        '
        'lblCeilingSubstrate
        '
        Me.lblCeilingSubstrate.Location = New System.Drawing.Point(249, 56)
        Me.lblCeilingSubstrate.Name = "lblCeilingSubstrate"
        Me.lblCeilingSubstrate.Size = New System.Drawing.Size(142, 21)
        Me.lblCeilingSubstrate.TabIndex = 10
        Me.lblCeilingSubstrate.Text = "None"
        '
        'lblCeilingSurface
        '
        Me.lblCeilingSurface.Location = New System.Drawing.Point(249, 26)
        Me.lblCeilingSurface.Name = "lblCeilingSurface"
        Me.lblCeilingSurface.Size = New System.Drawing.Size(142, 16)
        Me.lblCeilingSurface.TabIndex = 9
        Me.lblCeilingSurface.Text = "None"
        '
        'cmdPickCsubstrate
        '
        Me.cmdPickCsubstrate.Location = New System.Drawing.Point(11, 55)
        Me.cmdPickCsubstrate.Name = "cmdPickCsubstrate"
        Me.cmdPickCsubstrate.Size = New System.Drawing.Size(104, 23)
        Me.cmdPickCsubstrate.TabIndex = 8
        Me.cmdPickCsubstrate.Text = "Pick substrate mat"
        Me.cmdPickCsubstrate.UseVisualStyleBackColor = True
        '
        'cmdPickCsurface
        '
        Me.cmdPickCsurface.Location = New System.Drawing.Point(11, 26)
        Me.cmdPickCsurface.Name = "cmdPickCsurface"
        Me.cmdPickCsurface.Size = New System.Drawing.Size(104, 23)
        Me.cmdPickCsurface.TabIndex = 8
        Me.cmdPickCsurface.Text = "Pick surface mat"
        Me.cmdPickCsurface.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.txtFloorSubstrateThickness)
        Me.GroupBox3.Controls.Add(Me.cmdPickFsubstrate)
        Me.GroupBox3.Controls.Add(Me.Label18)
        Me.GroupBox3.Controls.Add(Me.cmdPickFsurface)
        Me.GroupBox3.Controls.Add(Me.txtFloorSurfaceThickness)
        Me.GroupBox3.Controls.Add(Me.Label19)
        Me.GroupBox3.Controls.Add(Me.lblFloorSurface)
        Me.GroupBox3.Controls.Add(Me.lblFloorSubstrate)
        Me.GroupBox3.Location = New System.Drawing.Point(346, 218)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(396, 99)
        Me.GroupBox3.TabIndex = 26
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Floor Materials"
        '
        'txtFloorSubstrateThickness
        '
        Me.txtFloorSubstrateThickness.Location = New System.Drawing.Point(208, 51)
        Me.txtFloorSubstrateThickness.Name = "txtFloorSubstrateThickness"
        Me.txtFloorSubstrateThickness.Size = New System.Drawing.Size(35, 20)
        Me.txtFloorSubstrateThickness.TabIndex = 20
        Me.txtFloorSubstrateThickness.Text = "150"
        '
        'cmdPickFsubstrate
        '
        Me.cmdPickFsubstrate.Location = New System.Drawing.Point(11, 54)
        Me.cmdPickFsubstrate.Name = "cmdPickFsubstrate"
        Me.cmdPickFsubstrate.Size = New System.Drawing.Size(104, 23)
        Me.cmdPickFsubstrate.TabIndex = 9
        Me.cmdPickFsubstrate.Text = "Pick substrate mat"
        Me.cmdPickFsubstrate.UseVisualStyleBackColor = True
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(121, 55)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(81, 13)
        Me.Label18.TabIndex = 19
        Me.Label18.Text = "Thickness (mm)"
        '
        'cmdPickFsurface
        '
        Me.cmdPickFsurface.Location = New System.Drawing.Point(11, 25)
        Me.cmdPickFsurface.Name = "cmdPickFsurface"
        Me.cmdPickFsurface.Size = New System.Drawing.Size(104, 23)
        Me.cmdPickFsurface.TabIndex = 10
        Me.cmdPickFsurface.Text = "Pick surface mat"
        Me.cmdPickFsurface.UseVisualStyleBackColor = True
        '
        'txtFloorSurfaceThickness
        '
        Me.txtFloorSurfaceThickness.Location = New System.Drawing.Point(208, 22)
        Me.txtFloorSurfaceThickness.Name = "txtFloorSurfaceThickness"
        Me.txtFloorSurfaceThickness.Size = New System.Drawing.Size(35, 20)
        Me.txtFloorSurfaceThickness.TabIndex = 18
        Me.txtFloorSurfaceThickness.Text = "13"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(121, 25)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(81, 13)
        Me.Label19.TabIndex = 17
        Me.Label19.Text = "Thickness (mm)"
        '
        'lblFloorSurface
        '
        Me.lblFloorSurface.Location = New System.Drawing.Point(249, 25)
        Me.lblFloorSurface.Name = "lblFloorSurface"
        Me.lblFloorSurface.Size = New System.Drawing.Size(142, 16)
        Me.lblFloorSurface.TabIndex = 15
        Me.lblFloorSurface.Text = "None"
        '
        'lblFloorSubstrate
        '
        Me.lblFloorSubstrate.Location = New System.Drawing.Point(249, 55)
        Me.lblFloorSubstrate.Name = "lblFloorSubstrate"
        Me.lblFloorSubstrate.Size = New System.Drawing.Size(142, 21)
        Me.lblFloorSubstrate.TabIndex = 16
        Me.lblFloorSubstrate.Text = "None"
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(357, 324)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(154, 26)
        Me.Button1.TabIndex = 27
        Me.Button1.Text = "Apply materials to all rooms"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'frmNewRoom
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(761, 359)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "frmNewRoom"
        Me.Text = "frmNewRoom"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents cmdDist_width As System.Windows.Forms.Button
    Friend WithEvents cmdDist_length As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtWidth As System.Windows.Forms.TextBox
    Friend WithEvents txtElevation As System.Windows.Forms.TextBox
    Friend WithEvents txtLength As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextAbsX As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents lblLength As System.Windows.Forms.Label
    Friend WithEvents txtMaxHeight As System.Windows.Forms.TextBox
    Friend WithEvents txtMinHeight As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextAbsY As System.Windows.Forms.TextBox
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtID As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents Label12 As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents txtWallSurfaceThickness As TextBox
    Friend WithEvents lblWallSurfaceThickness As Label
    Friend WithEvents cmdPickWSubstrate As Button
    Friend WithEvents cmdPickWsurface As Button
    Friend WithEvents lblWallSubstrate As Label
    Friend WithEvents lblWallSurface As Label
    Friend WithEvents txtWallSubstrateThickness As TextBox
    Friend WithEvents lblWallSubstrateThickness As Label
    Friend WithEvents cmdPickCsubstrate As Button
    Friend WithEvents cmdPickCsurface As Button
    Friend WithEvents cmdPickFsubstrate As Button
    Friend WithEvents cmdPickFsurface As Button
    Friend WithEvents txtCeilingSubstrateThickness As TextBox
    Friend WithEvents Label14 As Label
    Friend WithEvents txtCeilingSurfaceThickness As TextBox
    Friend WithEvents Label15 As Label
    Friend WithEvents lblCeilingSubstrate As Label
    Friend WithEvents lblCeilingSurface As Label
    Friend WithEvents txtFloorSubstrateThickness As TextBox
    Friend WithEvents Label18 As Label
    Friend WithEvents txtFloorSurfaceThickness As TextBox
    Friend WithEvents Label19 As Label
    Friend WithEvents lblFloorSurface As Label
    Friend WithEvents lblFloorSubstrate As Label
    Friend WithEvents ErrorProvider1 As ErrorProvider
    Friend WithEvents Button1 As Button
End Class
