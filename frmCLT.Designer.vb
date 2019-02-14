<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCLT
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCLT))
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.optCLTOFF = New System.Windows.Forms.RadioButton()
        Me.optCLTON = New System.Windows.Forms.RadioButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.numeric_wallareapercent = New System.Windows.Forms.NumericUpDown()
        Me.numeric_ceilareapercent = New System.Windows.Forms.NumericUpDown()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtCharTemp = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtLamellaDepth = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtCritFlux = New System.Windows.Forms.TextBox()
        Me.txtFlameFlux = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtCLTLoG = New System.Windows.Forms.TextBox()
        Me.txtCLTigtemp = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtCLTcalibration = New System.Windows.Forms.TextBox()
        Me.Panel_integralmodel = New System.Windows.Forms.Panel()
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel()
        Me.lblDebondTemp = New System.Windows.Forms.Label()
        Me.txtDebondTemp = New System.Windows.Forms.TextBox()
        Me.RB_dynamic = New System.Windows.Forms.RadioButton()
        Me.RB_Integral = New System.Windows.Forms.RadioButton()
        Me.RB_Kinetic = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.TXT_MoistureContent = New System.Windows.Forms.TextBox()
        Me.panel_kinetic = New System.Windows.Forms.TableLayoutPanel()
        Me.txtA_hemi = New System.Windows.Forms.TextBox()
        Me.txtA_cell = New System.Windows.Forms.TextBox()
        Me.txtA_lignin = New System.Windows.Forms.TextBox()
        Me.txtA_water = New System.Windows.Forms.TextBox()
        Me.txtE_hemi = New System.Windows.Forms.TextBox()
        Me.txtE_cell = New System.Windows.Forms.TextBox()
        Me.txtE_lignin = New System.Windows.Forms.TextBox()
        Me.txtE_water = New System.Windows.Forms.TextBox()
        Me.txtReact_hemi = New System.Windows.Forms.TextBox()
        Me.txtReact_cell = New System.Windows.Forms.TextBox()
        Me.txtReact_lignin = New System.Windows.Forms.TextBox()
        Me.txtReact_water = New System.Windows.Forms.TextBox()
        Me.txtInit_hemi = New System.Windows.Forms.TextBox()
        Me.txtInit_cell = New System.Windows.Forms.TextBox()
        Me.txtInit_lignin = New System.Windows.Forms.TextBox()
        Me.txtInit_water = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.TextBox_charyield_hemi = New System.Windows.Forms.TextBox()
        Me.TextBox_charyield_cell = New System.Windows.Forms.TextBox()
        Me.TextBox_charyield_lignin = New System.Windows.Forms.TextBox()
        Me.ComboBox1_k = New System.Windows.Forms.ComboBox()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numeric_wallareapercent, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numeric_ceilareapercent, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel_integralmodel.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.panel_kinetic.SuspendLayout()
        Me.SuspendLayout()
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(424, 582)
        Me.cmdClose.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(112, 35)
        Me.cmdClose.TabIndex = 0
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'optCLTOFF
        '
        Me.optCLTOFF.AutoSize = True
        Me.optCLTOFF.Checked = True
        Me.optCLTOFF.Location = New System.Drawing.Point(50, 582)
        Me.optCLTOFF.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.optCLTOFF.Name = "optCLTOFF"
        Me.optCLTOFF.Size = New System.Drawing.Size(148, 24)
        Me.optCLTOFF.TabIndex = 1
        Me.optCLTOFF.TabStop = True
        Me.optCLTOFF.Text = "CLT model is off"
        Me.optCLTOFF.UseVisualStyleBackColor = True
        '
        'optCLTON
        '
        Me.optCLTON.AutoSize = True
        Me.optCLTON.Location = New System.Drawing.Point(242, 582)
        Me.optCLTON.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.optCLTON.Name = "optCLTON"
        Me.optCLTON.Size = New System.Drawing.Size(147, 24)
        Me.optCLTON.TabIndex = 2
        Me.optCLTON.Text = "CLT model is on"
        Me.optCLTON.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 9)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(0, 20)
        Me.Label1.TabIndex = 3
        '
        'TextBox1
        '
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox1.Location = New System.Drawing.Point(54, 92)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(462, 222)
        Me.TextBox1.TabIndex = 4
        Me.TextBox1.Text = resources.GetString("TextBox1.Text")
        '
        'numeric_wallareapercent
        '
        Me.numeric_wallareapercent.Location = New System.Drawing.Point(334, 22)
        Me.numeric_wallareapercent.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.numeric_wallareapercent.Name = "numeric_wallareapercent"
        Me.numeric_wallareapercent.Size = New System.Drawing.Size(88, 26)
        Me.numeric_wallareapercent.TabIndex = 5
        Me.numeric_wallareapercent.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'numeric_ceilareapercent
        '
        Me.numeric_ceilareapercent.Location = New System.Drawing.Point(334, 68)
        Me.numeric_ceilareapercent.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.numeric_ceilareapercent.Name = "numeric_ceilareapercent"
        Me.numeric_ceilareapercent.Size = New System.Drawing.Size(88, 26)
        Me.numeric_ceilareapercent.TabIndex = 6
        Me.numeric_ceilareapercent.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(134, 25)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(192, 20)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Contribuing wall area (%)  "
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(116, 68)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(206, 20)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Contribuing ceiling area (%) "
        '
        'txtCharTemp
        '
        Me.txtCharTemp.Location = New System.Drawing.Point(334, 115)
        Me.txtCharTemp.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtCharTemp.Name = "txtCharTemp"
        Me.txtCharTemp.Size = New System.Drawing.Size(86, 26)
        Me.txtCharTemp.TabIndex = 9
        Me.txtCharTemp.Text = "300"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(165, 120)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(159, 20)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Char temperature (C)"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(156, 225)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(162, 20)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "Lamella thickness (m)"
        '
        'txtLamellaDepth
        '
        Me.txtLamellaDepth.Location = New System.Drawing.Point(334, 220)
        Me.txtLamellaDepth.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtLamellaDepth.Name = "txtLamellaDepth"
        Me.txtLamellaDepth.Size = New System.Drawing.Size(86, 26)
        Me.txtLamellaDepth.TabIndex = 13
        Me.txtLamellaDepth.Text = "0.020"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(105, 29)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(183, 20)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "Critical heat flux (kW/m2)"
        '
        'txtCritFlux
        '
        Me.txtCritFlux.Location = New System.Drawing.Point(310, 29)
        Me.txtCritFlux.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtCritFlux.Name = "txtCritFlux"
        Me.txtCritFlux.Size = New System.Drawing.Size(88, 26)
        Me.txtCritFlux.TabIndex = 15
        Me.txtCritFlux.Text = "16"
        '
        'txtFlameFlux
        '
        Me.txtFlameFlux.Location = New System.Drawing.Point(312, 75)
        Me.txtFlameFlux.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtFlameFlux.Name = "txtFlameFlux"
        Me.txtFlameFlux.Size = New System.Drawing.Size(86, 26)
        Me.txtFlameFlux.TabIndex = 16
        Me.txtFlameFlux.Text = "17"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(146, 80)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(144, 20)
        Me.Label7.TabIndex = 17
        Me.Label7.Text = "Flame flux (kW/m2)"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(54, 128)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(236, 20)
        Me.Label8.TabIndex = 18
        Me.Label8.Text = "Latent heat of gasification (kJ/g)"
        '
        'txtCLTLoG
        '
        Me.txtCLTLoG.Location = New System.Drawing.Point(312, 123)
        Me.txtCLTLoG.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtCLTLoG.Name = "txtCLTLoG"
        Me.txtCLTLoG.Size = New System.Drawing.Size(86, 26)
        Me.txtCLTLoG.TabIndex = 19
        Me.txtCLTLoG.Text = "1.6"
        '
        'txtCLTigtemp
        '
        Me.txtCLTigtemp.Location = New System.Drawing.Point(312, 174)
        Me.txtCLTigtemp.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtCLTigtemp.Name = "txtCLTigtemp"
        Me.txtCLTigtemp.Size = New System.Drawing.Size(86, 26)
        Me.txtCLTigtemp.TabIndex = 20
        Me.txtCLTigtemp.Text = "384"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(120, 178)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(177, 20)
        Me.Label9.TabIndex = 21
        Me.Label9.Text = "Ignition temperature (C)"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(165, 229)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(129, 20)
        Me.Label10.TabIndex = 22
        Me.Label10.Text = "Calibration factor"
        '
        'txtCLTcalibration
        '
        Me.txtCLTcalibration.Location = New System.Drawing.Point(310, 225)
        Me.txtCLTcalibration.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtCLTcalibration.Name = "txtCLTcalibration"
        Me.txtCLTcalibration.Size = New System.Drawing.Size(88, 26)
        Me.txtCLTcalibration.TabIndex = 23
        Me.txtCLTcalibration.Text = "1"
        '
        'Panel_integralmodel
        '
        Me.Panel_integralmodel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel_integralmodel.Controls.Add(Me.txtCLTcalibration)
        Me.Panel_integralmodel.Controls.Add(Me.Label10)
        Me.Panel_integralmodel.Controls.Add(Me.Label9)
        Me.Panel_integralmodel.Controls.Add(Me.txtCLTigtemp)
        Me.Panel_integralmodel.Controls.Add(Me.txtCLTLoG)
        Me.Panel_integralmodel.Controls.Add(Me.Label8)
        Me.Panel_integralmodel.Controls.Add(Me.Label7)
        Me.Panel_integralmodel.Controls.Add(Me.txtFlameFlux)
        Me.Panel_integralmodel.Controls.Add(Me.txtCritFlux)
        Me.Panel_integralmodel.Controls.Add(Me.Label6)
        Me.Panel_integralmodel.Location = New System.Drawing.Point(574, 345)
        Me.Panel_integralmodel.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Panel_integralmodel.Name = "Panel_integralmodel"
        Me.Panel_integralmodel.Size = New System.Drawing.Size(449, 279)
        Me.Panel_integralmodel.TabIndex = 24
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(375, 46)
        Me.LinkLabel1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(117, 20)
        Me.LinkLabel1.TabIndex = 25
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "See Reference"
        '
        'lblDebondTemp
        '
        Me.lblDebondTemp.AutoSize = True
        Me.lblDebondTemp.Location = New System.Drawing.Point(46, 168)
        Me.lblDebondTemp.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblDebondTemp.Name = "lblDebondTemp"
        Me.lblDebondTemp.Size = New System.Drawing.Size(269, 20)
        Me.lblDebondTemp.TabIndex = 26
        Me.lblDebondTemp.Text = "Adhesive debonding temperature (C)"
        '
        'txtDebondTemp
        '
        Me.txtDebondTemp.Location = New System.Drawing.Point(334, 169)
        Me.txtDebondTemp.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtDebondTemp.Name = "txtDebondTemp"
        Me.txtDebondTemp.Size = New System.Drawing.Size(85, 26)
        Me.txtDebondTemp.TabIndex = 27
        Me.txtDebondTemp.Text = "200"
        '
        'RB_dynamic
        '
        Me.RB_dynamic.AutoSize = True
        Me.RB_dynamic.Checked = True
        Me.RB_dynamic.Location = New System.Drawing.Point(28, 43)
        Me.RB_dynamic.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.RB_dynamic.Name = "RB_dynamic"
        Me.RB_dynamic.Size = New System.Drawing.Size(339, 24)
        Me.RB_dynamic.TabIndex = 0
        Me.RB_dynamic.TabStop = True
        Me.RB_dynamic.Text = "Global equivalence charring model (default)"
        Me.RB_dynamic.UseVisualStyleBackColor = True
        '
        'RB_Integral
        '
        Me.RB_Integral.AutoSize = True
        Me.RB_Integral.Location = New System.Drawing.Point(28, 360)
        Me.RB_Integral.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.RB_Integral.Name = "RB_Integral"
        Me.RB_Integral.Size = New System.Drawing.Size(240, 24)
        Me.RB_Integral.TabIndex = 1
        Me.RB_Integral.Text = "Integral wood pyrolysis model"
        Me.RB_Integral.UseVisualStyleBackColor = True
        '
        'RB_Kinetic
        '
        Me.RB_Kinetic.AutoSize = True
        Me.RB_Kinetic.Location = New System.Drawing.Point(28, 446)
        Me.RB_Kinetic.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.RB_Kinetic.Name = "RB_Kinetic"
        Me.RB_Kinetic.Size = New System.Drawing.Size(233, 24)
        Me.RB_Kinetic.TabIndex = 2
        Me.RB_Kinetic.Text = "Kinetic wood pyrolysis model"
        Me.RB_Kinetic.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ComboBox1_k)
        Me.GroupBox1.Controls.Add(Me.LinkLabel1)
        Me.GroupBox1.Controls.Add(Me.RB_Kinetic)
        Me.GroupBox1.Controls.Add(Me.RB_dynamic)
        Me.GroupBox1.Controls.Add(Me.RB_Integral)
        Me.GroupBox1.Controls.Add(Me.TextBox1)
        Me.GroupBox1.Location = New System.Drawing.Point(21, 22)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox1.Size = New System.Drawing.Size(544, 529)
        Me.GroupBox1.TabIndex = 30
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Select wood pyrolysis model"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label20)
        Me.GroupBox2.Controls.Add(Me.TXT_MoistureContent)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.txtLamellaDepth)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.txtDebondTemp)
        Me.GroupBox2.Controls.Add(Me.numeric_wallareapercent)
        Me.GroupBox2.Controls.Add(Me.lblDebondTemp)
        Me.GroupBox2.Controls.Add(Me.numeric_ceilareapercent)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.txtCharTemp)
        Me.GroupBox2.Location = New System.Drawing.Point(726, 9)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox2.Size = New System.Drawing.Size(452, 311)
        Me.GroupBox2.TabIndex = 31
        Me.GroupBox2.TabStop = False
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(156, 275)
        Me.Label20.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(156, 20)
        Me.Label20.TabIndex = 29
        Me.Label20.Text = "Moisture content (%)"
        '
        'TXT_MoistureContent
        '
        Me.TXT_MoistureContent.Location = New System.Drawing.Point(333, 271)
        Me.TXT_MoistureContent.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TXT_MoistureContent.Name = "TXT_MoistureContent"
        Me.TXT_MoistureContent.Size = New System.Drawing.Size(86, 26)
        Me.TXT_MoistureContent.TabIndex = 28
        Me.TXT_MoistureContent.Text = "10"
        '
        'panel_kinetic
        '
        Me.panel_kinetic.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.[Single]
        Me.panel_kinetic.ColumnCount = 5
        Me.panel_kinetic.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 120.0!))
        Me.panel_kinetic.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.panel_kinetic.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.panel_kinetic.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.panel_kinetic.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.panel_kinetic.Controls.Add(Me.txtA_hemi, 1, 1)
        Me.panel_kinetic.Controls.Add(Me.txtA_cell, 2, 1)
        Me.panel_kinetic.Controls.Add(Me.txtA_lignin, 3, 1)
        Me.panel_kinetic.Controls.Add(Me.txtA_water, 4, 1)
        Me.panel_kinetic.Controls.Add(Me.txtE_hemi, 1, 2)
        Me.panel_kinetic.Controls.Add(Me.txtE_cell, 2, 2)
        Me.panel_kinetic.Controls.Add(Me.txtE_lignin, 3, 2)
        Me.panel_kinetic.Controls.Add(Me.txtE_water, 4, 2)
        Me.panel_kinetic.Controls.Add(Me.txtReact_hemi, 1, 3)
        Me.panel_kinetic.Controls.Add(Me.txtReact_cell, 2, 3)
        Me.panel_kinetic.Controls.Add(Me.txtReact_lignin, 3, 3)
        Me.panel_kinetic.Controls.Add(Me.txtReact_water, 4, 3)
        Me.panel_kinetic.Controls.Add(Me.txtInit_hemi, 1, 4)
        Me.panel_kinetic.Controls.Add(Me.txtInit_cell, 2, 4)
        Me.panel_kinetic.Controls.Add(Me.txtInit_lignin, 3, 4)
        Me.panel_kinetic.Controls.Add(Me.txtInit_water, 4, 4)
        Me.panel_kinetic.Controls.Add(Me.Label16, 0, 1)
        Me.panel_kinetic.Controls.Add(Me.Label12, 1, 0)
        Me.panel_kinetic.Controls.Add(Me.Label17, 0, 2)
        Me.panel_kinetic.Controls.Add(Me.Label13, 2, 0)
        Me.panel_kinetic.Controls.Add(Me.Label18, 0, 3)
        Me.panel_kinetic.Controls.Add(Me.Label14, 3, 0)
        Me.panel_kinetic.Controls.Add(Me.Label19, 0, 4)
        Me.panel_kinetic.Controls.Add(Me.Label15, 4, 0)
        Me.panel_kinetic.Controls.Add(Me.Label11, 0, 5)
        Me.panel_kinetic.Controls.Add(Me.TextBox_charyield_hemi, 1, 5)
        Me.panel_kinetic.Controls.Add(Me.TextBox_charyield_cell, 2, 5)
        Me.panel_kinetic.Controls.Add(Me.TextBox_charyield_lignin, 3, 5)
        Me.panel_kinetic.Location = New System.Drawing.Point(574, 345)
        Me.panel_kinetic.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.panel_kinetic.Name = "panel_kinetic"
        Me.panel_kinetic.RowCount = 6
        Me.panel_kinetic.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.panel_kinetic.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.panel_kinetic.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.panel_kinetic.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.panel_kinetic.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.panel_kinetic.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.panel_kinetic.Size = New System.Drawing.Size(603, 280)
        Me.panel_kinetic.TabIndex = 26
        '
        'txtA_hemi
        '
        Me.txtA_hemi.Location = New System.Drawing.Point(126, 47)
        Me.txtA_hemi.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtA_hemi.Name = "txtA_hemi"
        Me.txtA_hemi.Size = New System.Drawing.Size(108, 26)
        Me.txtA_hemi.TabIndex = 5
        Me.txtA_hemi.Text = "1.64E+05"
        '
        'txtA_cell
        '
        Me.txtA_cell.Location = New System.Drawing.Point(246, 47)
        Me.txtA_cell.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtA_cell.Name = "txtA_cell"
        Me.txtA_cell.Size = New System.Drawing.Size(108, 26)
        Me.txtA_cell.TabIndex = 6
        Me.txtA_cell.Text = "1.98E+05"
        '
        'txtA_lignin
        '
        Me.txtA_lignin.Location = New System.Drawing.Point(366, 47)
        Me.txtA_lignin.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtA_lignin.Name = "txtA_lignin"
        Me.txtA_lignin.Size = New System.Drawing.Size(108, 26)
        Me.txtA_lignin.TabIndex = 7
        Me.txtA_lignin.Text = "1.52E+05"
        '
        'txtA_water
        '
        Me.txtA_water.Location = New System.Drawing.Point(486, 47)
        Me.txtA_water.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtA_water.Name = "txtA_water"
        Me.txtA_water.Size = New System.Drawing.Size(108, 26)
        Me.txtA_water.TabIndex = 8
        Me.txtA_water.Text = "1.0E+05"
        '
        'txtE_hemi
        '
        Me.txtE_hemi.Location = New System.Drawing.Point(126, 94)
        Me.txtE_hemi.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtE_hemi.Name = "txtE_hemi"
        Me.txtE_hemi.Size = New System.Drawing.Size(108, 26)
        Me.txtE_hemi.TabIndex = 9
        Me.txtE_hemi.Text = "3.25E+13"
        '
        'txtE_cell
        '
        Me.txtE_cell.Location = New System.Drawing.Point(246, 94)
        Me.txtE_cell.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtE_cell.Name = "txtE_cell"
        Me.txtE_cell.Size = New System.Drawing.Size(108, 26)
        Me.txtE_cell.TabIndex = 10
        Me.txtE_cell.Text = "3.51E+14"
        '
        'txtE_lignin
        '
        Me.txtE_lignin.Location = New System.Drawing.Point(366, 94)
        Me.txtE_lignin.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtE_lignin.Name = "txtE_lignin"
        Me.txtE_lignin.Size = New System.Drawing.Size(108, 26)
        Me.txtE_lignin.TabIndex = 11
        Me.txtE_lignin.Text = "8.41E+13"
        '
        'txtE_water
        '
        Me.txtE_water.Location = New System.Drawing.Point(486, 94)
        Me.txtE_water.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtE_water.Name = "txtE_water"
        Me.txtE_water.Size = New System.Drawing.Size(108, 26)
        Me.txtE_water.TabIndex = 12
        Me.txtE_water.Text = "1.0E+13"
        '
        'txtReact_hemi
        '
        Me.txtReact_hemi.Location = New System.Drawing.Point(126, 141)
        Me.txtReact_hemi.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtReact_hemi.Name = "txtReact_hemi"
        Me.txtReact_hemi.Size = New System.Drawing.Size(108, 26)
        Me.txtReact_hemi.TabIndex = 13
        Me.txtReact_hemi.Text = "2.1"
        '
        'txtReact_cell
        '
        Me.txtReact_cell.Location = New System.Drawing.Point(246, 141)
        Me.txtReact_cell.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtReact_cell.Name = "txtReact_cell"
        Me.txtReact_cell.Size = New System.Drawing.Size(108, 26)
        Me.txtReact_cell.TabIndex = 14
        Me.txtReact_cell.Text = "1.1"
        '
        'txtReact_lignin
        '
        Me.txtReact_lignin.Location = New System.Drawing.Point(366, 141)
        Me.txtReact_lignin.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtReact_lignin.Name = "txtReact_lignin"
        Me.txtReact_lignin.Size = New System.Drawing.Size(108, 26)
        Me.txtReact_lignin.TabIndex = 15
        Me.txtReact_lignin.Text = "5.0"
        '
        'txtReact_water
        '
        Me.txtReact_water.Location = New System.Drawing.Point(486, 141)
        Me.txtReact_water.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtReact_water.Name = "txtReact_water"
        Me.txtReact_water.Size = New System.Drawing.Size(108, 26)
        Me.txtReact_water.TabIndex = 16
        Me.txtReact_water.Text = "1.0"
        '
        'txtInit_hemi
        '
        Me.txtInit_hemi.Location = New System.Drawing.Point(126, 188)
        Me.txtInit_hemi.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtInit_hemi.Name = "txtInit_hemi"
        Me.txtInit_hemi.Size = New System.Drawing.Size(108, 26)
        Me.txtInit_hemi.TabIndex = 17
        Me.txtInit_hemi.Text = "0.37"
        '
        'txtInit_cell
        '
        Me.txtInit_cell.Location = New System.Drawing.Point(246, 188)
        Me.txtInit_cell.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtInit_cell.Name = "txtInit_cell"
        Me.txtInit_cell.Size = New System.Drawing.Size(108, 26)
        Me.txtInit_cell.TabIndex = 18
        Me.txtInit_cell.Text = "0.44"
        '
        'txtInit_lignin
        '
        Me.txtInit_lignin.Location = New System.Drawing.Point(366, 188)
        Me.txtInit_lignin.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtInit_lignin.Name = "txtInit_lignin"
        Me.txtInit_lignin.Size = New System.Drawing.Size(108, 26)
        Me.txtInit_lignin.TabIndex = 19
        Me.txtInit_lignin.Text = "0.09"
        '
        'txtInit_water
        '
        Me.txtInit_water.Location = New System.Drawing.Point(486, 188)
        Me.txtInit_water.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtInit_water.Name = "txtInit_water"
        Me.txtInit_water.Size = New System.Drawing.Size(108, 26)
        Me.txtInit_water.TabIndex = 20
        Me.txtInit_water.Text = "0.10"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(5, 42)
        Me.Label16.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(108, 40)
        Me.Label16.TabIndex = 21
        Me.Label16.Text = "Activation energy (J/mol)"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(126, 1)
        Me.Label12.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(106, 40)
        Me.Label12.TabIndex = 1
        Me.Label12.Text = "Hemicellulose"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(5, 89)
        Me.Label17.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(112, 40)
        Me.Label17.TabIndex = 22
        Me.Label17.Text = "Pre-exp factor (1/s)"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(246, 1)
        Me.Label13.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(73, 20)
        Me.Label13.TabIndex = 2
        Me.Label13.Text = "Cellulose"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(5, 136)
        Me.Label18.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(77, 40)
        Me.Label18.TabIndex = 23
        Me.Label18.Text = "Reaction order (-)"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(366, 1)
        Me.Label14.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(51, 20)
        Me.Label14.TabIndex = 3
        Me.Label14.Text = "Lignin"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(5, 183)
        Me.Label19.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(97, 40)
        Me.Label19.TabIndex = 24
        Me.Label19.Text = "Initial comp. fraction (-)"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(486, 1)
        Me.Label15.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(52, 20)
        Me.Label15.TabIndex = 4
        Me.Label15.Text = "Water"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(5, 230)
        Me.Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(97, 20)
        Me.Label11.TabIndex = 25
        Me.Label11.Text = "Char yield (-)"
        '
        'TextBox_charyield_hemi
        '
        Me.TextBox_charyield_hemi.Location = New System.Drawing.Point(126, 235)
        Me.TextBox_charyield_hemi.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TextBox_charyield_hemi.Name = "TextBox_charyield_hemi"
        Me.TextBox_charyield_hemi.Size = New System.Drawing.Size(108, 26)
        Me.TextBox_charyield_hemi.TabIndex = 26
        Me.TextBox_charyield_hemi.Text = "0.40"
        '
        'TextBox_charyield_cell
        '
        Me.TextBox_charyield_cell.Location = New System.Drawing.Point(246, 235)
        Me.TextBox_charyield_cell.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TextBox_charyield_cell.Name = "TextBox_charyield_cell"
        Me.TextBox_charyield_cell.Size = New System.Drawing.Size(108, 26)
        Me.TextBox_charyield_cell.TabIndex = 27
        Me.TextBox_charyield_cell.Text = "0.40"
        '
        'TextBox_charyield_lignin
        '
        Me.TextBox_charyield_lignin.Location = New System.Drawing.Point(366, 235)
        Me.TextBox_charyield_lignin.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TextBox_charyield_lignin.Name = "TextBox_charyield_lignin"
        Me.TextBox_charyield_lignin.Size = New System.Drawing.Size(108, 26)
        Me.TextBox_charyield_lignin.TabIndex = 28
        Me.TextBox_charyield_lignin.Text = "0.40"
        '
        'ComboBox1_k
        '
        Me.ComboBox1_k.FormattingEnabled = True
        Me.ComboBox1_k.Items.AddRange(New Object() {"Hankalin thermal conductivity", "Eurocode 5 thermal conductivity", "Hybrid", "Constant thermal conductivity (0.2)"})
        Me.ComboBox1_k.Location = New System.Drawing.Point(221, 494)
        Me.ComboBox1_k.MaxDropDownItems = 4
        Me.ComboBox1_k.Name = "ComboBox1_k"
        Me.ComboBox1_k.Size = New System.Drawing.Size(316, 28)
        Me.ComboBox1_k.TabIndex = 27
        '
        'frmCLT
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1190, 648)
        Me.Controls.Add(Me.panel_kinetic)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Panel_integralmodel)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.optCLTON)
        Me.Controls.Add(Me.optCLTOFF)
        Me.Controls.Add(Me.cmdClose)
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "frmCLT"
        Me.ShowInTaskbar = False
        Me.Text = "CLT Model"
        Me.TopMost = True
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numeric_wallareapercent, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numeric_ceilareapercent, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel_integralmodel.ResumeLayout(False)
        Me.Panel_integralmodel.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.panel_kinetic.ResumeLayout(False)
        Me.panel_kinetic.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents optCLTON As System.Windows.Forms.RadioButton
    Friend WithEvents optCLTOFF As System.Windows.Forms.RadioButton
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents numeric_ceilareapercent As System.Windows.Forms.NumericUpDown
    Friend WithEvents numeric_wallareapercent As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtCharTemp As System.Windows.Forms.TextBox
    Friend WithEvents txtLamellaDepth As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents txtCritFlux As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents txtFlameFlux As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents txtCLTigtemp As TextBox
    Friend WithEvents txtCLTLoG As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents txtCLTcalibration As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents Panel_integralmodel As Panel
    Friend WithEvents LinkLabel1 As LinkLabel
    Friend WithEvents txtDebondTemp As TextBox
    Friend WithEvents lblDebondTemp As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents RB_Kinetic As RadioButton
    Friend WithEvents RB_dynamic As RadioButton
    Friend WithEvents RB_Integral As RadioButton
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents panel_kinetic As TableLayoutPanel
    Friend WithEvents txtA_hemi As TextBox
    Friend WithEvents txtA_cell As TextBox
    Friend WithEvents txtA_lignin As TextBox
    Friend WithEvents txtA_water As TextBox
    Friend WithEvents txtE_hemi As TextBox
    Friend WithEvents txtE_cell As TextBox
    Friend WithEvents txtE_lignin As TextBox
    Friend WithEvents txtE_water As TextBox
    Friend WithEvents txtReact_hemi As TextBox
    Friend WithEvents txtReact_cell As TextBox
    Friend WithEvents txtReact_lignin As TextBox
    Friend WithEvents txtReact_water As TextBox
    Friend WithEvents txtInit_hemi As TextBox
    Friend WithEvents txtInit_cell As TextBox
    Friend WithEvents txtInit_lignin As TextBox
    Friend WithEvents txtInit_water As TextBox
    Friend WithEvents Label16 As Label
    Friend WithEvents Label12 As Label
    Friend WithEvents Label17 As Label
    Friend WithEvents Label13 As Label
    Friend WithEvents Label18 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents Label19 As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents TextBox_charyield_hemi As TextBox
    Friend WithEvents TextBox_charyield_cell As TextBox
    Friend WithEvents TextBox_charyield_lignin As TextBox
    Friend WithEvents Label20 As Label
    Friend WithEvents TXT_MoistureContent As TextBox
    Friend WithEvents ComboBox1_k As ComboBox
End Class
