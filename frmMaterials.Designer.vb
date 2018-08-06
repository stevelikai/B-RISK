<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMaterials
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
        Dim Material_DescriptionLabel As System.Windows.Forms.Label
        Dim Thermal_ConductivityLabel As System.Windows.Forms.Label
        Dim Specific_HeatLabel As System.Windows.Forms.Label
        Dim DensityLabel As System.Windows.Forms.Label
        Dim EmissivityLabel As System.Windows.Forms.Label
        Dim Min_Temp_for_SpreadLabel As System.Windows.Forms.Label
        Dim Flame_Spread_ParameterLabel As System.Windows.Forms.Label
        Dim Cone_Data_FileLabel As System.Windows.Forms.Label
        Dim Heat_of_CombustionLabel As System.Windows.Forms.Label
        Dim CommentsLabel As System.Windows.Forms.Label
        Dim Soot_YieldLabel As System.Windows.Forms.Label
        Dim CO2_YieldLabel As System.Windows.Forms.Label
        Dim H2O_YieldLabel As System.Windows.Forms.Label
        Dim Calibration_FactorLabel As System.Windows.Forms.Label
        Dim Comments2Label As System.Windows.Forms.Label
        Dim Material_DescriptionLabel1 As System.Windows.Forms.Label
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMaterials))
        Me.Table1BindingNavigator = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.BindingNavigatorAddNewItem = New System.Windows.Forms.ToolStripButton()
        Me.Table1BindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.ThermalDataSet = New BRISK.thermalDataSet()
        Me.BindingNavigatorCountItem = New System.Windows.Forms.ToolStripLabel()
        Me.BindingNavigatorDeleteItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMoveFirstItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMovePreviousItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorSeparator = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorPositionItem = New System.Windows.Forms.ToolStripTextBox()
        Me.BindingNavigatorSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorMoveNextItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMoveLastItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.Table1BindingNavigatorSaveItem = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton_CancelItem = New System.Windows.Forms.ToolStripButton()
        Me.Thermal_ConductivityTextBox = New System.Windows.Forms.TextBox()
        Me.Specific_HeatTextBox = New System.Windows.Forms.TextBox()
        Me.DensityTextBox = New System.Windows.Forms.TextBox()
        Me.EmissivityTextBox = New System.Windows.Forms.TextBox()
        Me.Min_Temp_for_SpreadTextBox = New System.Windows.Forms.TextBox()
        Me.Flame_Spread_ParameterTextBox = New System.Windows.Forms.TextBox()
        Me.Cone_Data_FileTextBox = New System.Windows.Forms.TextBox()
        Me.Heat_of_CombustionTextBox = New System.Windows.Forms.TextBox()
        Me.CommentsTextBox = New System.Windows.Forms.TextBox()
        Me.Soot_YieldTextBox = New System.Windows.Forms.TextBox()
        Me.CO2_YieldTextBox = New System.Windows.Forms.TextBox()
        Me.H2O_YieldTextBox = New System.Windows.Forms.TextBox()
        Me.Comments2TextBox = New System.Windows.Forms.TextBox()
        Me.Table1TableAdapter = New BRISK.thermalDataSetTableAdapters.Table1TableAdapter()
        Me.TableAdapterManager = New BRISK.thermalDataSetTableAdapters.TableAdapterManager()
        Me.Material_DescriptionComboBox = New System.Windows.Forms.ComboBox()
        Me.Material_DescriptionTextBox = New System.Windows.Forms.TextBox()
        Me.Calibration_FactorTextBox = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Material_DescriptionLabel = New System.Windows.Forms.Label()
        Thermal_ConductivityLabel = New System.Windows.Forms.Label()
        Specific_HeatLabel = New System.Windows.Forms.Label()
        DensityLabel = New System.Windows.Forms.Label()
        EmissivityLabel = New System.Windows.Forms.Label()
        Min_Temp_for_SpreadLabel = New System.Windows.Forms.Label()
        Flame_Spread_ParameterLabel = New System.Windows.Forms.Label()
        Cone_Data_FileLabel = New System.Windows.Forms.Label()
        Heat_of_CombustionLabel = New System.Windows.Forms.Label()
        CommentsLabel = New System.Windows.Forms.Label()
        Soot_YieldLabel = New System.Windows.Forms.Label()
        CO2_YieldLabel = New System.Windows.Forms.Label()
        H2O_YieldLabel = New System.Windows.Forms.Label()
        Calibration_FactorLabel = New System.Windows.Forms.Label()
        Comments2Label = New System.Windows.Forms.Label()
        Material_DescriptionLabel1 = New System.Windows.Forms.Label()
        CType(Me.Table1BindingNavigator, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Table1BindingNavigator.SuspendLayout()
        CType(Me.Table1BindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ThermalDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Material_DescriptionLabel
        '
        Material_DescriptionLabel.AutoSize = True
        Material_DescriptionLabel.Location = New System.Drawing.Point(40, 132)
        Material_DescriptionLabel.Name = "Material_DescriptionLabel"
        Material_DescriptionLabel.Size = New System.Drawing.Size(103, 13)
        Material_DescriptionLabel.TabIndex = 3
        Material_DescriptionLabel.Text = "Material Description:"
        '
        'Thermal_ConductivityLabel
        '
        Thermal_ConductivityLabel.AutoSize = True
        Thermal_ConductivityLabel.Location = New System.Drawing.Point(40, 158)
        Thermal_ConductivityLabel.Name = "Thermal_ConductivityLabel"
        Thermal_ConductivityLabel.Size = New System.Drawing.Size(109, 13)
        Thermal_ConductivityLabel.TabIndex = 5
        Thermal_ConductivityLabel.Text = "Thermal Conductivity:"
        '
        'Specific_HeatLabel
        '
        Specific_HeatLabel.AutoSize = True
        Specific_HeatLabel.Location = New System.Drawing.Point(40, 184)
        Specific_HeatLabel.Name = "Specific_HeatLabel"
        Specific_HeatLabel.Size = New System.Drawing.Size(74, 13)
        Specific_HeatLabel.TabIndex = 7
        Specific_HeatLabel.Text = "Specific Heat:"
        '
        'DensityLabel
        '
        DensityLabel.AutoSize = True
        DensityLabel.Location = New System.Drawing.Point(40, 210)
        DensityLabel.Name = "DensityLabel"
        DensityLabel.Size = New System.Drawing.Size(45, 13)
        DensityLabel.TabIndex = 9
        DensityLabel.Text = "Density:"
        '
        'EmissivityLabel
        '
        EmissivityLabel.AutoSize = True
        EmissivityLabel.Location = New System.Drawing.Point(40, 236)
        EmissivityLabel.Name = "EmissivityLabel"
        EmissivityLabel.Size = New System.Drawing.Size(55, 13)
        EmissivityLabel.TabIndex = 11
        EmissivityLabel.Text = "Emissivity:"
        '
        'Min_Temp_for_SpreadLabel
        '
        Min_Temp_for_SpreadLabel.AutoSize = True
        Min_Temp_for_SpreadLabel.Location = New System.Drawing.Point(40, 262)
        Min_Temp_for_SpreadLabel.Name = "Min_Temp_for_SpreadLabel"
        Min_Temp_for_SpreadLabel.Size = New System.Drawing.Size(109, 13)
        Min_Temp_for_SpreadLabel.TabIndex = 13
        Min_Temp_for_SpreadLabel.Text = "Min Temp for Spread:"
        '
        'Flame_Spread_ParameterLabel
        '
        Flame_Spread_ParameterLabel.AutoSize = True
        Flame_Spread_ParameterLabel.Location = New System.Drawing.Point(40, 288)
        Flame_Spread_ParameterLabel.Name = "Flame_Spread_ParameterLabel"
        Flame_Spread_ParameterLabel.Size = New System.Drawing.Size(126, 13)
        Flame_Spread_ParameterLabel.TabIndex = 15
        Flame_Spread_ParameterLabel.Text = "Flame Spread Parameter:"
        '
        'Cone_Data_FileLabel
        '
        Cone_Data_FileLabel.AutoSize = True
        Cone_Data_FileLabel.Location = New System.Drawing.Point(40, 314)
        Cone_Data_FileLabel.Name = "Cone_Data_FileLabel"
        Cone_Data_FileLabel.Size = New System.Drawing.Size(80, 13)
        Cone_Data_FileLabel.TabIndex = 17
        Cone_Data_FileLabel.Text = "Cone Data File:"
        '
        'Heat_of_CombustionLabel
        '
        Heat_of_CombustionLabel.AutoSize = True
        Heat_of_CombustionLabel.Location = New System.Drawing.Point(40, 340)
        Heat_of_CombustionLabel.Name = "Heat_of_CombustionLabel"
        Heat_of_CombustionLabel.Size = New System.Drawing.Size(103, 13)
        Heat_of_CombustionLabel.TabIndex = 19
        Heat_of_CombustionLabel.Text = "Heat of Combustion:"
        '
        'CommentsLabel
        '
        CommentsLabel.AutoSize = True
        CommentsLabel.Location = New System.Drawing.Point(40, 533)
        CommentsLabel.Name = "CommentsLabel"
        CommentsLabel.Size = New System.Drawing.Size(59, 13)
        CommentsLabel.TabIndex = 21
        CommentsLabel.Text = "Comments:"
        CommentsLabel.Visible = False
        '
        'Soot_YieldLabel
        '
        Soot_YieldLabel.AutoSize = True
        Soot_YieldLabel.Location = New System.Drawing.Point(40, 366)
        Soot_YieldLabel.Name = "Soot_YieldLabel"
        Soot_YieldLabel.Size = New System.Drawing.Size(58, 13)
        Soot_YieldLabel.TabIndex = 23
        Soot_YieldLabel.Text = "Soot Yield:"
        '
        'CO2_YieldLabel
        '
        CO2_YieldLabel.AutoSize = True
        CO2_YieldLabel.Location = New System.Drawing.Point(40, 392)
        CO2_YieldLabel.Name = "CO2_YieldLabel"
        CO2_YieldLabel.Size = New System.Drawing.Size(57, 13)
        CO2_YieldLabel.TabIndex = 25
        CO2_YieldLabel.Text = "CO2 Yield:"
        '
        'H2O_YieldLabel
        '
        H2O_YieldLabel.AutoSize = True
        H2O_YieldLabel.Location = New System.Drawing.Point(40, 418)
        H2O_YieldLabel.Name = "H2O_YieldLabel"
        H2O_YieldLabel.Size = New System.Drawing.Size(58, 13)
        H2O_YieldLabel.TabIndex = 27
        H2O_YieldLabel.Text = "H2O Yield:"
        '
        'Calibration_FactorLabel
        '
        Calibration_FactorLabel.AutoSize = True
        Calibration_FactorLabel.Location = New System.Drawing.Point(40, 511)
        Calibration_FactorLabel.Name = "Calibration_FactorLabel"
        Calibration_FactorLabel.Size = New System.Drawing.Size(92, 13)
        Calibration_FactorLabel.TabIndex = 29
        Calibration_FactorLabel.Text = "Calibration Factor:"
        Calibration_FactorLabel.Visible = False
        '
        'Comments2Label
        '
        Comments2Label.AutoSize = True
        Comments2Label.Location = New System.Drawing.Point(39, 454)
        Comments2Label.Name = "Comments2Label"
        Comments2Label.Size = New System.Drawing.Size(59, 13)
        Comments2Label.TabIndex = 31
        Comments2Label.Text = "Comments:"
        '
        'Material_DescriptionLabel1
        '
        Material_DescriptionLabel1.AutoSize = True
        Material_DescriptionLabel1.Location = New System.Drawing.Point(13, 22)
        Material_DescriptionLabel1.Name = "Material_DescriptionLabel1"
        Material_DescriptionLabel1.Size = New System.Drawing.Size(79, 13)
        Material_DescriptionLabel1.TabIndex = 33
        Material_DescriptionLabel1.Text = "Find a Material:"
        '
        'Table1BindingNavigator
        '
        Me.Table1BindingNavigator.AddNewItem = Me.BindingNavigatorAddNewItem
        Me.Table1BindingNavigator.BindingSource = Me.Table1BindingSource
        Me.Table1BindingNavigator.CountItem = Me.BindingNavigatorCountItem
        Me.Table1BindingNavigator.DeleteItem = Me.BindingNavigatorDeleteItem
        Me.Table1BindingNavigator.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BindingNavigatorMoveFirstItem, Me.BindingNavigatorMovePreviousItem, Me.BindingNavigatorSeparator, Me.BindingNavigatorPositionItem, Me.BindingNavigatorCountItem, Me.BindingNavigatorSeparator1, Me.BindingNavigatorMoveNextItem, Me.BindingNavigatorMoveLastItem, Me.BindingNavigatorSeparator2, Me.BindingNavigatorAddNewItem, Me.BindingNavigatorDeleteItem, Me.Table1BindingNavigatorSaveItem, Me.ToolStripButton_CancelItem})
        Me.Table1BindingNavigator.Location = New System.Drawing.Point(0, 0)
        Me.Table1BindingNavigator.MoveFirstItem = Me.BindingNavigatorMoveFirstItem
        Me.Table1BindingNavigator.MoveLastItem = Me.BindingNavigatorMoveLastItem
        Me.Table1BindingNavigator.MoveNextItem = Me.BindingNavigatorMoveNextItem
        Me.Table1BindingNavigator.MovePreviousItem = Me.BindingNavigatorMovePreviousItem
        Me.Table1BindingNavigator.Name = "Table1BindingNavigator"
        Me.Table1BindingNavigator.PositionItem = Me.BindingNavigatorPositionItem
        Me.Table1BindingNavigator.Size = New System.Drawing.Size(359, 25)
        Me.Table1BindingNavigator.TabIndex = 0
        Me.Table1BindingNavigator.Text = "BindingNavigator1"
        '
        'BindingNavigatorAddNewItem
        '
        Me.BindingNavigatorAddNewItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorAddNewItem.Image = CType(resources.GetObject("BindingNavigatorAddNewItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorAddNewItem.Name = "BindingNavigatorAddNewItem"
        Me.BindingNavigatorAddNewItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorAddNewItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorAddNewItem.Text = "Add new"
        '
        'Table1BindingSource
        '
        Me.Table1BindingSource.DataMember = "Table1"
        Me.Table1BindingSource.DataSource = Me.ThermalDataSet
        '
        'ThermalDataSet
        '
        Me.ThermalDataSet.DataSetName = "thermalDataSet"
        Me.ThermalDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'BindingNavigatorCountItem
        '
        Me.BindingNavigatorCountItem.Name = "BindingNavigatorCountItem"
        Me.BindingNavigatorCountItem.Size = New System.Drawing.Size(35, 22)
        Me.BindingNavigatorCountItem.Text = "of {0}"
        Me.BindingNavigatorCountItem.ToolTipText = "Total number of items"
        '
        'BindingNavigatorDeleteItem
        '
        Me.BindingNavigatorDeleteItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorDeleteItem.Image = CType(resources.GetObject("BindingNavigatorDeleteItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorDeleteItem.Name = "BindingNavigatorDeleteItem"
        Me.BindingNavigatorDeleteItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorDeleteItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorDeleteItem.Text = "Delete"
        '
        'BindingNavigatorMoveFirstItem
        '
        Me.BindingNavigatorMoveFirstItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveFirstItem.Image = CType(resources.GetObject("BindingNavigatorMoveFirstItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveFirstItem.Name = "BindingNavigatorMoveFirstItem"
        Me.BindingNavigatorMoveFirstItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveFirstItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveFirstItem.Text = "Move first"
        '
        'BindingNavigatorMovePreviousItem
        '
        Me.BindingNavigatorMovePreviousItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMovePreviousItem.Image = CType(resources.GetObject("BindingNavigatorMovePreviousItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMovePreviousItem.Name = "BindingNavigatorMovePreviousItem"
        Me.BindingNavigatorMovePreviousItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMovePreviousItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMovePreviousItem.Text = "Move previous"
        '
        'BindingNavigatorSeparator
        '
        Me.BindingNavigatorSeparator.Name = "BindingNavigatorSeparator"
        Me.BindingNavigatorSeparator.Size = New System.Drawing.Size(6, 25)
        '
        'BindingNavigatorPositionItem
        '
        Me.BindingNavigatorPositionItem.AccessibleName = "Position"
        Me.BindingNavigatorPositionItem.AutoSize = False
        Me.BindingNavigatorPositionItem.Name = "BindingNavigatorPositionItem"
        Me.BindingNavigatorPositionItem.Size = New System.Drawing.Size(50, 21)
        Me.BindingNavigatorPositionItem.Text = "0"
        Me.BindingNavigatorPositionItem.ToolTipText = "Current position"
        '
        'BindingNavigatorSeparator1
        '
        Me.BindingNavigatorSeparator1.Name = "BindingNavigatorSeparator1"
        Me.BindingNavigatorSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'BindingNavigatorMoveNextItem
        '
        Me.BindingNavigatorMoveNextItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveNextItem.Image = CType(resources.GetObject("BindingNavigatorMoveNextItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveNextItem.Name = "BindingNavigatorMoveNextItem"
        Me.BindingNavigatorMoveNextItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveNextItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveNextItem.Text = "Move next"
        '
        'BindingNavigatorMoveLastItem
        '
        Me.BindingNavigatorMoveLastItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveLastItem.Image = CType(resources.GetObject("BindingNavigatorMoveLastItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveLastItem.Name = "BindingNavigatorMoveLastItem"
        Me.BindingNavigatorMoveLastItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveLastItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveLastItem.Text = "Move last"
        '
        'BindingNavigatorSeparator2
        '
        Me.BindingNavigatorSeparator2.Name = "BindingNavigatorSeparator2"
        Me.BindingNavigatorSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'Table1BindingNavigatorSaveItem
        '
        Me.Table1BindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.Table1BindingNavigatorSaveItem.Image = CType(resources.GetObject("Table1BindingNavigatorSaveItem.Image"), System.Drawing.Image)
        Me.Table1BindingNavigatorSaveItem.Name = "Table1BindingNavigatorSaveItem"
        Me.Table1BindingNavigatorSaveItem.Size = New System.Drawing.Size(23, 22)
        Me.Table1BindingNavigatorSaveItem.Text = "Save Data"
        '
        'ToolStripButton_CancelItem
        '
        Me.ToolStripButton_CancelItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton_CancelItem.Image = CType(resources.GetObject("ToolStripButton_CancelItem.Image"), System.Drawing.Image)
        Me.ToolStripButton_CancelItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton_CancelItem.Name = "ToolStripButton_CancelItem"
        Me.ToolStripButton_CancelItem.Size = New System.Drawing.Size(47, 22)
        Me.ToolStripButton_CancelItem.Text = "Cancel"
        '
        'Thermal_ConductivityTextBox
        '
        Me.Thermal_ConductivityTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Thermal Conductivity", True))
        Me.Thermal_ConductivityTextBox.Location = New System.Drawing.Point(172, 155)
        Me.Thermal_ConductivityTextBox.Name = "Thermal_ConductivityTextBox"
        Me.Thermal_ConductivityTextBox.Size = New System.Drawing.Size(100, 20)
        Me.Thermal_ConductivityTextBox.TabIndex = 6
        '
        'Specific_HeatTextBox
        '
        Me.Specific_HeatTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Specific Heat", True))
        Me.Specific_HeatTextBox.Location = New System.Drawing.Point(172, 181)
        Me.Specific_HeatTextBox.Name = "Specific_HeatTextBox"
        Me.Specific_HeatTextBox.Size = New System.Drawing.Size(100, 20)
        Me.Specific_HeatTextBox.TabIndex = 8
        '
        'DensityTextBox
        '
        Me.DensityTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Density", True))
        Me.DensityTextBox.Location = New System.Drawing.Point(172, 207)
        Me.DensityTextBox.Name = "DensityTextBox"
        Me.DensityTextBox.Size = New System.Drawing.Size(100, 20)
        Me.DensityTextBox.TabIndex = 10
        '
        'EmissivityTextBox
        '
        Me.EmissivityTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Emissivity", True))
        Me.EmissivityTextBox.Location = New System.Drawing.Point(172, 233)
        Me.EmissivityTextBox.Name = "EmissivityTextBox"
        Me.EmissivityTextBox.Size = New System.Drawing.Size(100, 20)
        Me.EmissivityTextBox.TabIndex = 12
        '
        'Min_Temp_for_SpreadTextBox
        '
        Me.Min_Temp_for_SpreadTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Min Temp for Spread", True))
        Me.Min_Temp_for_SpreadTextBox.Location = New System.Drawing.Point(172, 259)
        Me.Min_Temp_for_SpreadTextBox.Name = "Min_Temp_for_SpreadTextBox"
        Me.Min_Temp_for_SpreadTextBox.Size = New System.Drawing.Size(100, 20)
        Me.Min_Temp_for_SpreadTextBox.TabIndex = 14
        '
        'Flame_Spread_ParameterTextBox
        '
        Me.Flame_Spread_ParameterTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Flame Spread Parameter", True))
        Me.Flame_Spread_ParameterTextBox.Location = New System.Drawing.Point(172, 285)
        Me.Flame_Spread_ParameterTextBox.Name = "Flame_Spread_ParameterTextBox"
        Me.Flame_Spread_ParameterTextBox.Size = New System.Drawing.Size(100, 20)
        Me.Flame_Spread_ParameterTextBox.TabIndex = 16
        '
        'Cone_Data_FileTextBox
        '
        Me.Cone_Data_FileTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Cone Data File", True))
        Me.Cone_Data_FileTextBox.Location = New System.Drawing.Point(172, 311)
        Me.Cone_Data_FileTextBox.Name = "Cone_Data_FileTextBox"
        Me.Cone_Data_FileTextBox.Size = New System.Drawing.Size(155, 20)
        Me.Cone_Data_FileTextBox.TabIndex = 18
        '
        'Heat_of_CombustionTextBox
        '
        Me.Heat_of_CombustionTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Heat of Combustion", True))
        Me.Heat_of_CombustionTextBox.Location = New System.Drawing.Point(172, 337)
        Me.Heat_of_CombustionTextBox.Name = "Heat_of_CombustionTextBox"
        Me.Heat_of_CombustionTextBox.Size = New System.Drawing.Size(100, 20)
        Me.Heat_of_CombustionTextBox.TabIndex = 20
        '
        'CommentsTextBox
        '
        Me.CommentsTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Comments", True))
        Me.CommentsTextBox.Location = New System.Drawing.Point(172, 530)
        Me.CommentsTextBox.Multiline = True
        Me.CommentsTextBox.Name = "CommentsTextBox"
        Me.CommentsTextBox.Size = New System.Drawing.Size(100, 20)
        Me.CommentsTextBox.TabIndex = 22
        Me.CommentsTextBox.Visible = False
        '
        'Soot_YieldTextBox
        '
        Me.Soot_YieldTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Soot Yield", True))
        Me.Soot_YieldTextBox.Location = New System.Drawing.Point(172, 363)
        Me.Soot_YieldTextBox.Name = "Soot_YieldTextBox"
        Me.Soot_YieldTextBox.Size = New System.Drawing.Size(100, 20)
        Me.Soot_YieldTextBox.TabIndex = 24
        '
        'CO2_YieldTextBox
        '
        Me.CO2_YieldTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "CO2 Yield", True))
        Me.CO2_YieldTextBox.Location = New System.Drawing.Point(172, 389)
        Me.CO2_YieldTextBox.Name = "CO2_YieldTextBox"
        Me.CO2_YieldTextBox.Size = New System.Drawing.Size(100, 20)
        Me.CO2_YieldTextBox.TabIndex = 26
        '
        'H2O_YieldTextBox
        '
        Me.H2O_YieldTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "H2O Yield", True))
        Me.H2O_YieldTextBox.Location = New System.Drawing.Point(172, 415)
        Me.H2O_YieldTextBox.Name = "H2O_YieldTextBox"
        Me.H2O_YieldTextBox.Size = New System.Drawing.Size(100, 20)
        Me.H2O_YieldTextBox.TabIndex = 28
        '
        'Comments2TextBox
        '
        Me.Comments2TextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Comments2", True))
        Me.Comments2TextBox.Location = New System.Drawing.Point(172, 441)
        Me.Comments2TextBox.Multiline = True
        Me.Comments2TextBox.Name = "Comments2TextBox"
        Me.Comments2TextBox.Size = New System.Drawing.Size(155, 83)
        Me.Comments2TextBox.TabIndex = 32
        '
        'Table1TableAdapter
        '
        Me.Table1TableAdapter.ClearBeforeFill = True
        '
        'TableAdapterManager
        '
        Me.TableAdapterManager.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager.Table1TableAdapter = Me.Table1TableAdapter
        Me.TableAdapterManager.UpdateOrder = BRISK.thermalDataSetTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        '
        'Material_DescriptionComboBox
        '
        Me.Material_DescriptionComboBox.CausesValidation = False
        Me.Material_DescriptionComboBox.DataSource = Me.Table1BindingSource
        Me.Material_DescriptionComboBox.DisplayMember = "Material Description"
        Me.Material_DescriptionComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Material_DescriptionComboBox.FormattingEnabled = True
        Me.Material_DescriptionComboBox.Location = New System.Drawing.Point(122, 19)
        Me.Material_DescriptionComboBox.Name = "Material_DescriptionComboBox"
        Me.Material_DescriptionComboBox.Size = New System.Drawing.Size(176, 21)
        Me.Material_DescriptionComboBox.TabIndex = 34
        '
        'Material_DescriptionTextBox
        '
        Me.Material_DescriptionTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Material Description", True))
        Me.Material_DescriptionTextBox.Location = New System.Drawing.Point(172, 129)
        Me.Material_DescriptionTextBox.Name = "Material_DescriptionTextBox"
        Me.Material_DescriptionTextBox.Size = New System.Drawing.Size(155, 20)
        Me.Material_DescriptionTextBox.TabIndex = 4
        '
        'Calibration_FactorTextBox
        '
        Me.Calibration_FactorTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Calibration Factor", True))
        Me.Calibration_FactorTextBox.Location = New System.Drawing.Point(172, 508)
        Me.Calibration_FactorTextBox.Name = "Calibration_FactorTextBox"
        Me.Calibration_FactorTextBox.Size = New System.Drawing.Size(100, 20)
        Me.Calibration_FactorTextBox.TabIndex = 30
        Me.Calibration_FactorTextBox.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Material_DescriptionLabel1)
        Me.GroupBox1.Controls.Add(Me.Material_DescriptionComboBox)
        Me.GroupBox1.Location = New System.Drawing.Point(29, 43)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(315, 57)
        Me.GroupBox1.TabIndex = 35
        Me.GroupBox1.TabStop = False
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(252, 535)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 36
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(279, 156)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(38, 13)
        Me.Label1.TabIndex = 37
        Me.Label1.Text = "W/mK"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(279, 183)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(36, 13)
        Me.Label2.TabIndex = 38
        Me.Label2.Text = "J/kgK"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(279, 207)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(38, 13)
        Me.Label3.TabIndex = 39
        Me.Label3.Text = "kg/m3"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(279, 233)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(10, 13)
        Me.Label4.TabIndex = 40
        Me.Label4.Text = "-"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(279, 259)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(14, 13)
        Me.Label5.TabIndex = 41
        Me.Label5.Text = "C"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(279, 285)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(10, 13)
        Me.Label6.TabIndex = 42
        Me.Label6.Text = "-"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(279, 365)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(24, 13)
        Me.Label7.TabIndex = 43
        Me.Label7.Text = "g/g"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(279, 389)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(24, 13)
        Me.Label8.TabIndex = 44
        Me.Label8.Text = "g/g"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(279, 415)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(24, 13)
        Me.Label9.TabIndex = 45
        Me.Label9.Text = "g/g"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(279, 340)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(29, 13)
        Me.Label10.TabIndex = 46
        Me.Label10.Text = "kJ/g"
        '
        'frmMaterials
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(359, 579)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Material_DescriptionLabel)
        Me.Controls.Add(Me.Material_DescriptionTextBox)
        Me.Controls.Add(Thermal_ConductivityLabel)
        Me.Controls.Add(Me.Thermal_ConductivityTextBox)
        Me.Controls.Add(Specific_HeatLabel)
        Me.Controls.Add(Me.Specific_HeatTextBox)
        Me.Controls.Add(DensityLabel)
        Me.Controls.Add(Me.DensityTextBox)
        Me.Controls.Add(EmissivityLabel)
        Me.Controls.Add(Me.EmissivityTextBox)
        Me.Controls.Add(Min_Temp_for_SpreadLabel)
        Me.Controls.Add(Me.Min_Temp_for_SpreadTextBox)
        Me.Controls.Add(Flame_Spread_ParameterLabel)
        Me.Controls.Add(Me.Flame_Spread_ParameterTextBox)
        Me.Controls.Add(Cone_Data_FileLabel)
        Me.Controls.Add(Me.Cone_Data_FileTextBox)
        Me.Controls.Add(Heat_of_CombustionLabel)
        Me.Controls.Add(Me.Heat_of_CombustionTextBox)
        Me.Controls.Add(CommentsLabel)
        Me.Controls.Add(Me.CommentsTextBox)
        Me.Controls.Add(Soot_YieldLabel)
        Me.Controls.Add(Me.Soot_YieldTextBox)
        Me.Controls.Add(CO2_YieldLabel)
        Me.Controls.Add(Me.CO2_YieldTextBox)
        Me.Controls.Add(H2O_YieldLabel)
        Me.Controls.Add(Me.H2O_YieldTextBox)
        Me.Controls.Add(Calibration_FactorLabel)
        Me.Controls.Add(Me.Calibration_FactorTextBox)
        Me.Controls.Add(Comments2Label)
        Me.Controls.Add(Me.Comments2TextBox)
        Me.Controls.Add(Me.Table1BindingNavigator)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frmMaterials"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Materials Database"
        CType(Me.Table1BindingNavigator, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Table1BindingNavigator.ResumeLayout(False)
        Me.Table1BindingNavigator.PerformLayout()
        CType(Me.Table1BindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ThermalDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ThermalDataSet As thermalDataSet
    Friend WithEvents Table1BindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Table1TableAdapter As thermalDataSetTableAdapters.Table1TableAdapter
    Friend WithEvents TableAdapterManager As thermalDataSetTableAdapters.TableAdapterManager
    Friend WithEvents Table1BindingNavigator As System.Windows.Forms.BindingNavigator
    Friend WithEvents BindingNavigatorAddNewItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorCountItem As System.Windows.Forms.ToolStripLabel
    Friend WithEvents BindingNavigatorDeleteItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMoveFirstItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMovePreviousItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorSeparator As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BindingNavigatorPositionItem As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents BindingNavigatorSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BindingNavigatorMoveNextItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMoveLastItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents Table1BindingNavigatorSaveItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents Thermal_ConductivityTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Specific_HeatTextBox As System.Windows.Forms.TextBox
    Friend WithEvents DensityTextBox As System.Windows.Forms.TextBox
    Friend WithEvents EmissivityTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Min_Temp_for_SpreadTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Flame_Spread_ParameterTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Cone_Data_FileTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Heat_of_CombustionTextBox As System.Windows.Forms.TextBox
    Friend WithEvents CommentsTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Soot_YieldTextBox As System.Windows.Forms.TextBox
    Friend WithEvents CO2_YieldTextBox As System.Windows.Forms.TextBox
    Friend WithEvents H2O_YieldTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Comments2TextBox As System.Windows.Forms.TextBox
    Friend WithEvents Material_DescriptionComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents ToolStripButton_CancelItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents Material_DescriptionTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Calibration_FactorTextBox As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
End Class
