<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFireObjectDB
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
        Dim Energy_YieldLabel As System.Windows.Forms.Label
        Dim HCN_YieldLabel As System.Windows.Forms.Label
        Dim CO2_YieldLabel As System.Windows.Forms.Label
        Dim Fire_HeightLabel As System.Windows.Forms.Label
        Dim Object_TypeLabel As System.Windows.Forms.Label
        Dim DescriptionLabel As System.Windows.Forms.Label
        Dim lblTimeHeat_Release As System.Windows.Forms.Label
        Dim Soot_YieldLabel As System.Windows.Forms.Label
        Dim Max_HRRLabel As System.Windows.Forms.Label
        Dim Growth_RateLabel As System.Windows.Forms.Label
        Dim FuelLabel As System.Windows.Forms.Label
        Dim DescriptionLabel1 As System.Windows.Forms.Label
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFireObjectDB))
        Me.FireDataSet1 = New fireDataSet
        Me.Fire_DataBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.Fire_DataTableAdapter = New fireDataSetTableAdapters.Fire_DataTableAdapter
        Me.TableAdapterManager = New fireDataSetTableAdapters.TableAdapterManager
        Me.Fire_DataBindingNavigator = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.BindingNavigatorAddNewItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorCountItem = New System.Windows.Forms.ToolStripLabel
        Me.BindingNavigatorDeleteItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorMoveFirstItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorMovePreviousItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorSeparator = New System.Windows.Forms.ToolStripSeparator
        Me.BindingNavigatorPositionItem = New System.Windows.Forms.ToolStripTextBox
        Me.BindingNavigatorSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.BindingNavigatorMoveNextItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorMoveLastItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorSeparator2 = New System.Windows.Forms.ToolStripSeparator
        Me.Fire_DataBindingNavigatorSaveItem = New System.Windows.Forms.ToolStripButton
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.ToolStripButton2 = New System.Windows.Forms.ToolStripButton
        Me.ToolStripLabel1 = New System.Windows.Forms.ToolStripLabel
        Me.Energy_YieldTextBox = New System.Windows.Forms.TextBox
        Me.HCN_YieldTextBox = New System.Windows.Forms.TextBox
        Me.CO2_YieldTextBox = New System.Windows.Forms.TextBox
        Me.Fire_HeightTextBox = New System.Windows.Forms.TextBox
        Me.Object_TypeComboBox = New System.Windows.Forms.ComboBox
        Me.DescriptionTextBox = New System.Windows.Forms.TextBox
        Me.Time___Heat_ReleaseTextBox = New System.Windows.Forms.TextBox
        Me.Soot_YieldTextBox = New System.Windows.Forms.TextBox
        Me.Max_HRRTextBox = New System.Windows.Forms.TextBox
        Me.Growth_RateComboBox = New System.Windows.Forms.ComboBox
        Me.FuelComboBox = New System.Windows.Forms.ComboBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.DescriptionListBox = New System.Windows.Forms.ListBox
        Me.cmdClose = New System.Windows.Forms.Button
        Energy_YieldLabel = New System.Windows.Forms.Label
        HCN_YieldLabel = New System.Windows.Forms.Label
        CO2_YieldLabel = New System.Windows.Forms.Label
        Fire_HeightLabel = New System.Windows.Forms.Label
        Object_TypeLabel = New System.Windows.Forms.Label
        DescriptionLabel = New System.Windows.Forms.Label
        lblTimeHeat_Release = New System.Windows.Forms.Label
        Soot_YieldLabel = New System.Windows.Forms.Label
        Max_HRRLabel = New System.Windows.Forms.Label
        Growth_RateLabel = New System.Windows.Forms.Label
        FuelLabel = New System.Windows.Forms.Label
        DescriptionLabel1 = New System.Windows.Forms.Label
        CType(Me.FireDataSet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Fire_DataBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Fire_DataBindingNavigator, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Fire_DataBindingNavigator.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Energy_YieldLabel
        '
        Energy_YieldLabel.AutoSize = True
        Energy_YieldLabel.Location = New System.Drawing.Point(24, 72)
        Energy_YieldLabel.Name = "Energy_YieldLabel"
        Energy_YieldLabel.Size = New System.Drawing.Size(134, 13)
        Energy_YieldLabel.TabIndex = 3
        Energy_YieldLabel.Text = "Heat of Combustion (kJ/g):"
        '
        'HCN_YieldLabel
        '
        HCN_YieldLabel.AutoSize = True
        HCN_YieldLabel.Location = New System.Drawing.Point(65, 106)
        HCN_YieldLabel.Name = "HCN_YieldLabel"
        HCN_YieldLabel.Size = New System.Drawing.Size(85, 13)
        HCN_YieldLabel.TabIndex = 7
        HCN_YieldLabel.Text = "HCN Yield (g/g):"
        '
        'CO2_YieldLabel
        '
        CO2_YieldLabel.AutoSize = True
        CO2_YieldLabel.Location = New System.Drawing.Point(65, 138)
        CO2_YieldLabel.Name = "CO2_YieldLabel"
        CO2_YieldLabel.Size = New System.Drawing.Size(83, 13)
        CO2_YieldLabel.TabIndex = 11
        CO2_YieldLabel.Text = "CO2 Yield (g/g):"
        '
        'Fire_HeightLabel
        '
        Fire_HeightLabel.AutoSize = True
        Fire_HeightLabel.Location = New System.Drawing.Point(17, 58)
        Fire_HeightLabel.Name = "Fire_HeightLabel"
        Fire_HeightLabel.Size = New System.Drawing.Size(78, 13)
        Fire_HeightLabel.TabIndex = 13
        Fire_HeightLabel.Text = "Fire Height (m):"
        '
        'Object_TypeLabel
        '
        Object_TypeLabel.AutoSize = True
        Object_TypeLabel.Location = New System.Drawing.Point(17, 89)
        Object_TypeLabel.Name = "Object_TypeLabel"
        Object_TypeLabel.Size = New System.Drawing.Size(68, 13)
        Object_TypeLabel.TabIndex = 15
        Object_TypeLabel.Text = "Object Type:"
        '
        'DescriptionLabel
        '
        DescriptionLabel.AutoSize = True
        DescriptionLabel.Location = New System.Drawing.Point(17, 25)
        DescriptionLabel.Name = "DescriptionLabel"
        DescriptionLabel.Size = New System.Drawing.Size(63, 13)
        DescriptionLabel.TabIndex = 17
        DescriptionLabel.Text = "Description:"
        '
        'lblTimeHeat_Release
        '
        lblTimeHeat_Release.AutoSize = True
        lblTimeHeat_Release.Location = New System.Drawing.Point(24, 289)
        lblTimeHeat_Release.Name = "lblTimeHeat_Release"
        lblTimeHeat_Release.Size = New System.Drawing.Size(107, 13)
        lblTimeHeat_Release.TabIndex = 19
        lblTimeHeat_Release.Text = "Time - Heat Release:"
        '
        'Soot_YieldLabel
        '
        Soot_YieldLabel.AutoSize = True
        Soot_YieldLabel.Location = New System.Drawing.Point(65, 168)
        Soot_YieldLabel.Name = "Soot_YieldLabel"
        Soot_YieldLabel.Size = New System.Drawing.Size(84, 13)
        Soot_YieldLabel.TabIndex = 21
        Soot_YieldLabel.Text = "Soot Yield (g/g):"
        '
        'Max_HRRLabel
        '
        Max_HRRLabel.AutoSize = True
        Max_HRRLabel.Location = New System.Drawing.Point(34, 36)
        Max_HRRLabel.Name = "Max_HRRLabel"
        Max_HRRLabel.Size = New System.Drawing.Size(83, 13)
        Max_HRRLabel.TabIndex = 23
        Max_HRRLabel.Text = "Max HRR (kW):"
        '
        'Growth_RateLabel
        '
        Growth_RateLabel.AutoSize = True
        Growth_RateLabel.Location = New System.Drawing.Point(34, 72)
        Growth_RateLabel.Name = "Growth_RateLabel"
        Growth_RateLabel.Size = New System.Drawing.Size(70, 13)
        Growth_RateLabel.TabIndex = 25
        Growth_RateLabel.Text = "Growth Rate:"
        '
        'FuelLabel
        '
        FuelLabel.AutoSize = True
        FuelLabel.Location = New System.Drawing.Point(24, 37)
        FuelLabel.Name = "FuelLabel"
        FuelLabel.Size = New System.Drawing.Size(30, 13)
        FuelLabel.TabIndex = 27
        FuelLabel.Text = "Fuel:"
        '
        'DescriptionLabel1
        '
        DescriptionLabel1.AutoSize = True
        DescriptionLabel1.Location = New System.Drawing.Point(24, 46)
        DescriptionLabel1.Name = "DescriptionLabel1"
        DescriptionLabel1.Size = New System.Drawing.Size(98, 13)
        DescriptionLabel1.TabIndex = 31
        DescriptionLabel1.Text = "Select a fire object:"
        '
        'FireDataSet1
        '
        Me.FireDataSet1.DataSetName = "fireDataSet1"
        Me.FireDataSet1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'Fire_DataBindingSource
        '
        Me.Fire_DataBindingSource.DataMember = "Fire Data"
        Me.Fire_DataBindingSource.DataSource = Me.FireDataSet1
        '
        'Fire_DataTableAdapter
        '
        Me.Fire_DataTableAdapter.ClearBeforeFill = True
        '
        'TableAdapterManager
        '
        Me.TableAdapterManager.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager.Fire_DataTableAdapter = Me.Fire_DataTableAdapter
        Me.TableAdapterManager.UpdateOrder = fireDataSetTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        '
        'Fire_DataBindingNavigator
        '
        Me.Fire_DataBindingNavigator.AddNewItem = Me.BindingNavigatorAddNewItem
        Me.Fire_DataBindingNavigator.BindingSource = Me.Fire_DataBindingSource
        Me.Fire_DataBindingNavigator.CountItem = Me.BindingNavigatorCountItem
        Me.Fire_DataBindingNavigator.DeleteItem = Me.BindingNavigatorDeleteItem
        Me.Fire_DataBindingNavigator.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BindingNavigatorMoveFirstItem, Me.BindingNavigatorMovePreviousItem, Me.BindingNavigatorSeparator, Me.BindingNavigatorPositionItem, Me.BindingNavigatorCountItem, Me.BindingNavigatorSeparator1, Me.BindingNavigatorMoveNextItem, Me.BindingNavigatorMoveLastItem, Me.BindingNavigatorSeparator2, Me.BindingNavigatorAddNewItem, Me.BindingNavigatorDeleteItem, Me.Fire_DataBindingNavigatorSaveItem, Me.ToolStripButton1, Me.ToolStripSeparator1, Me.ToolStripButton2, Me.ToolStripLabel1})
        Me.Fire_DataBindingNavigator.Location = New System.Drawing.Point(0, 0)
        Me.Fire_DataBindingNavigator.MoveFirstItem = Me.BindingNavigatorMoveFirstItem
        Me.Fire_DataBindingNavigator.MoveLastItem = Me.BindingNavigatorMoveLastItem
        Me.Fire_DataBindingNavigator.MoveNextItem = Me.BindingNavigatorMoveNextItem
        Me.Fire_DataBindingNavigator.MovePreviousItem = Me.BindingNavigatorMovePreviousItem
        Me.Fire_DataBindingNavigator.Name = "Fire_DataBindingNavigator"
        Me.Fire_DataBindingNavigator.PositionItem = Me.BindingNavigatorPositionItem
        Me.Fire_DataBindingNavigator.Size = New System.Drawing.Size(806, 25)
        Me.Fire_DataBindingNavigator.TabIndex = 0
        Me.Fire_DataBindingNavigator.Text = "BindingNavigator1"
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
        'Fire_DataBindingNavigatorSaveItem
        '
        Me.Fire_DataBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.Fire_DataBindingNavigatorSaveItem.Image = CType(resources.GetObject("Fire_DataBindingNavigatorSaveItem.Image"), System.Drawing.Image)
        Me.Fire_DataBindingNavigatorSaveItem.Name = "Fire_DataBindingNavigatorSaveItem"
        Me.Fire_DataBindingNavigatorSaveItem.Size = New System.Drawing.Size(23, 22)
        Me.Fire_DataBindingNavigatorSaveItem.Text = "Save Data"
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(47, 22)
        Me.ToolStripButton1.Text = "Cancel"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStripButton2
        '
        Me.ToolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton2.Image = CType(resources.GetObject("ToolStripButton2.Image"), System.Drawing.Image)
        Me.ToolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton2.Name = "ToolStripButton2"
        Me.ToolStripButton2.Size = New System.Drawing.Size(150, 22)
        Me.ToolStripButton2.Text = "Add to Current Simulation"
        '
        'ToolStripLabel1
        '
        Me.ToolStripLabel1.Name = "ToolStripLabel1"
        Me.ToolStripLabel1.Size = New System.Drawing.Size(95, 22)
        Me.ToolStripLabel1.Text = "Send to Item List"
        '
        'Energy_YieldTextBox
        '
        Me.Energy_YieldTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Fire_DataBindingSource, "Energy Yield", True))
        Me.Energy_YieldTextBox.Location = New System.Drawing.Point(164, 69)
        Me.Energy_YieldTextBox.Name = "Energy_YieldTextBox"
        Me.Energy_YieldTextBox.Size = New System.Drawing.Size(94, 20)
        Me.Energy_YieldTextBox.TabIndex = 4
        '
        'HCN_YieldTextBox
        '
        Me.HCN_YieldTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Fire_DataBindingSource, "HCN Yield", True))
        Me.HCN_YieldTextBox.Location = New System.Drawing.Point(164, 103)
        Me.HCN_YieldTextBox.Name = "HCN_YieldTextBox"
        Me.HCN_YieldTextBox.Size = New System.Drawing.Size(94, 20)
        Me.HCN_YieldTextBox.TabIndex = 8
        '
        'CO2_YieldTextBox
        '
        Me.CO2_YieldTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Fire_DataBindingSource, "CO2 Yield", True))
        Me.CO2_YieldTextBox.Location = New System.Drawing.Point(164, 135)
        Me.CO2_YieldTextBox.Name = "CO2_YieldTextBox"
        Me.CO2_YieldTextBox.Size = New System.Drawing.Size(94, 20)
        Me.CO2_YieldTextBox.TabIndex = 12
        '
        'Fire_HeightTextBox
        '
        Me.Fire_HeightTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Fire_DataBindingSource, "Fire Height", True))
        Me.Fire_HeightTextBox.Location = New System.Drawing.Point(130, 55)
        Me.Fire_HeightTextBox.Name = "Fire_HeightTextBox"
        Me.Fire_HeightTextBox.Size = New System.Drawing.Size(84, 20)
        Me.Fire_HeightTextBox.TabIndex = 14
        '
        'Object_TypeComboBox
        '
        Me.Object_TypeComboBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Fire_DataBindingSource, "Object Type", True))
        Me.Object_TypeComboBox.FormattingEnabled = True
        Me.Object_TypeComboBox.Items.AddRange(New Object() {"Furniture", "Gas Burner", "T Squared Fire", "Vehicle", "Other"})
        Me.Object_TypeComboBox.Location = New System.Drawing.Point(130, 86)
        Me.Object_TypeComboBox.Name = "Object_TypeComboBox"
        Me.Object_TypeComboBox.Size = New System.Drawing.Size(162, 21)
        Me.Object_TypeComboBox.TabIndex = 16
        '
        'DescriptionTextBox
        '
        Me.DescriptionTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Fire_DataBindingSource, "Description", True))
        Me.DescriptionTextBox.Location = New System.Drawing.Point(130, 22)
        Me.DescriptionTextBox.Name = "DescriptionTextBox"
        Me.DescriptionTextBox.Size = New System.Drawing.Size(162, 20)
        Me.DescriptionTextBox.TabIndex = 18
        '
        'Time___Heat_ReleaseTextBox
        '
        Me.Time___Heat_ReleaseTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Fire_DataBindingSource, "Time - Heat Release", True))
        Me.Time___Heat_ReleaseTextBox.Location = New System.Drawing.Point(27, 314)
        Me.Time___Heat_ReleaseTextBox.Multiline = True
        Me.Time___Heat_ReleaseTextBox.Name = "Time___Heat_ReleaseTextBox"
        Me.Time___Heat_ReleaseTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.Time___Heat_ReleaseTextBox.Size = New System.Drawing.Size(104, 184)
        Me.Time___Heat_ReleaseTextBox.TabIndex = 20
        '
        'Soot_YieldTextBox
        '
        Me.Soot_YieldTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Fire_DataBindingSource, "Soot Yield", True))
        Me.Soot_YieldTextBox.Location = New System.Drawing.Point(164, 165)
        Me.Soot_YieldTextBox.Name = "Soot_YieldTextBox"
        Me.Soot_YieldTextBox.Size = New System.Drawing.Size(94, 20)
        Me.Soot_YieldTextBox.TabIndex = 22
        '
        'Max_HRRTextBox
        '
        Me.Max_HRRTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Fire_DataBindingSource, "Max HRR", True))
        Me.Max_HRRTextBox.Location = New System.Drawing.Point(147, 33)
        Me.Max_HRRTextBox.Name = "Max_HRRTextBox"
        Me.Max_HRRTextBox.Size = New System.Drawing.Size(111, 20)
        Me.Max_HRRTextBox.TabIndex = 24
        '
        'Growth_RateComboBox
        '
        Me.Growth_RateComboBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Fire_DataBindingSource, "Growth Rate", True))
        Me.Growth_RateComboBox.FormattingEnabled = True
        Me.Growth_RateComboBox.Items.AddRange(New Object() {"Ultra Fast", "Fast", "Medium", "Slow"})
        Me.Growth_RateComboBox.Location = New System.Drawing.Point(147, 69)
        Me.Growth_RateComboBox.Name = "Growth_RateComboBox"
        Me.Growth_RateComboBox.Size = New System.Drawing.Size(111, 21)
        Me.Growth_RateComboBox.TabIndex = 26
        '
        'FuelComboBox
        '
        Me.FuelComboBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Fire_DataBindingSource, "Fuel", True))
        Me.FuelComboBox.FormattingEnabled = True
        Me.FuelComboBox.Items.AddRange(New Object() {"User Specified", "Ethanol", "Methane", "Hexane", "Heptane", "Propane", "Polyurethane Foam (GM23)", "Polymethylmethacrylate", "Wood"})
        Me.FuelComboBox.Location = New System.Drawing.Point(137, 34)
        Me.FuelComboBox.Name = "FuelComboBox"
        Me.FuelComboBox.Size = New System.Drawing.Size(121, 21)
        Me.FuelComboBox.TabIndex = 28
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(DescriptionLabel)
        Me.GroupBox1.Controls.Add(Me.DescriptionTextBox)
        Me.GroupBox1.Controls.Add(Fire_HeightLabel)
        Me.GroupBox1.Controls.Add(Me.Fire_HeightTextBox)
        Me.GroupBox1.Controls.Add(Object_TypeLabel)
        Me.GroupBox1.Controls.Add(Me.Object_TypeComboBox)
        Me.GroupBox1.Location = New System.Drawing.Point(409, 46)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(310, 123)
        Me.GroupBox1.TabIndex = 29
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "General"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(FuelLabel)
        Me.GroupBox2.Controls.Add(Me.FuelComboBox)
        Me.GroupBox2.Controls.Add(Energy_YieldLabel)
        Me.GroupBox2.Controls.Add(Me.Soot_YieldTextBox)
        Me.GroupBox2.Controls.Add(Me.Energy_YieldTextBox)
        Me.GroupBox2.Controls.Add(Soot_YieldLabel)
        Me.GroupBox2.Controls.Add(HCN_YieldLabel)
        Me.GroupBox2.Controls.Add(Me.CO2_YieldTextBox)
        Me.GroupBox2.Controls.Add(Me.HCN_YieldTextBox)
        Me.GroupBox2.Controls.Add(CO2_YieldLabel)
        Me.GroupBox2.Location = New System.Drawing.Point(409, 185)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(310, 200)
        Me.GroupBox2.TabIndex = 30
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Yields"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Max_HRRLabel)
        Me.GroupBox3.Controls.Add(Me.Growth_RateComboBox)
        Me.GroupBox3.Controls.Add(Growth_RateLabel)
        Me.GroupBox3.Controls.Add(Me.Max_HRRTextBox)
        Me.GroupBox3.Location = New System.Drawing.Point(409, 403)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(310, 107)
        Me.GroupBox3.TabIndex = 31
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "T Squared Fire"
        '
        'DescriptionListBox
        '
        Me.DescriptionListBox.DataSource = Me.Fire_DataBindingSource
        Me.DescriptionListBox.DisplayMember = "Description"
        Me.DescriptionListBox.FormattingEnabled = True
        Me.DescriptionListBox.Location = New System.Drawing.Point(27, 65)
        Me.DescriptionListBox.Name = "DescriptionListBox"
        Me.DescriptionListBox.Size = New System.Drawing.Size(362, 199)
        Me.DescriptionListBox.TabIndex = 32
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(725, 487)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 33
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'frmFireObjectDB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(806, 533)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(DescriptionLabel1)
        Me.Controls.Add(Me.DescriptionListBox)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(lblTimeHeat_Release)
        Me.Controls.Add(Me.Time___Heat_ReleaseTextBox)
        Me.Controls.Add(Me.Fire_DataBindingNavigator)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmFireObjectDB"
        Me.ShowInTaskbar = False
        Me.Text = "Fire Object Database"
        CType(Me.FireDataSet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Fire_DataBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Fire_DataBindingNavigator, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Fire_DataBindingNavigator.ResumeLayout(False)
        Me.Fire_DataBindingNavigator.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents FireDataSet1 As fireDataSet
    Friend WithEvents Fire_DataBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Fire_DataTableAdapter As fireDataSetTableAdapters.Fire_DataTableAdapter
    Friend WithEvents TableAdapterManager As fireDataSetTableAdapters.TableAdapterManager
    Friend WithEvents Fire_DataBindingNavigator As System.Windows.Forms.BindingNavigator
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
    Friend WithEvents Fire_DataBindingNavigatorSaveItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents Energy_YieldTextBox As System.Windows.Forms.TextBox
    Friend WithEvents HCN_YieldTextBox As System.Windows.Forms.TextBox
    Friend WithEvents CO2_YieldTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Fire_HeightTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Object_TypeComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents DescriptionTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Time___Heat_ReleaseTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Soot_YieldTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Max_HRRTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Growth_RateComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents FuelComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents DescriptionListBox As System.Windows.Forms.ListBox
    Friend WithEvents ToolStripButton1 As System.Windows.Forms.ToolStripButton
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    'Friend WithEvents AxMSChart1 As AxMSChart20Lib.AxMSChart
    Friend WithEvents ToolStripButton2 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripLabel1 As System.Windows.Forms.ToolStripLabel
End Class
