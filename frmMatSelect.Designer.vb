<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMatSelect
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
        Dim Soot_YieldLabel As System.Windows.Forms.Label
        Dim CO2_YieldLabel As System.Windows.Forms.Label
        Dim H2O_YieldLabel As System.Windows.Forms.Label
        Dim Comments2Label As System.Windows.Forms.Label
        Me.ThermalDataSet = New BRISK.thermalDataSet()
        Me.Table1BindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.Table1TableAdapter = New BRISK.thermalDataSetTableAdapters.Table1TableAdapter()
        Me.TableAdapterManager = New BRISK.thermalDataSetTableAdapters.TableAdapterManager()
        Me.Material_DescriptionListBox = New System.Windows.Forms.ListBox()
        Me.Thermal_ConductivityTextBox = New System.Windows.Forms.TextBox()
        Me.Specific_HeatTextBox = New System.Windows.Forms.TextBox()
        Me.DensityTextBox = New System.Windows.Forms.TextBox()
        Me.EmissivityTextBox = New System.Windows.Forms.TextBox()
        Me.Min_Temp_for_SpreadTextBox = New System.Windows.Forms.TextBox()
        Me.Flame_Spread_ParameterTextBox = New System.Windows.Forms.TextBox()
        Me.Cone_Data_FileTextBox = New System.Windows.Forms.TextBox()
        Me.Heat_of_CombustionTextBox = New System.Windows.Forms.TextBox()
        Me.Soot_YieldTextBox = New System.Windows.Forms.TextBox()
        Me.CO2_YieldTextBox = New System.Windows.Forms.TextBox()
        Me.H2O_YieldTextBox = New System.Windows.Forms.TextBox()
        Me.Comments2TextBox = New System.Windows.Forms.TextBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdSelectMat = New System.Windows.Forms.Button()
        Material_DescriptionLabel = New System.Windows.Forms.Label()
        Thermal_ConductivityLabel = New System.Windows.Forms.Label()
        Specific_HeatLabel = New System.Windows.Forms.Label()
        DensityLabel = New System.Windows.Forms.Label()
        EmissivityLabel = New System.Windows.Forms.Label()
        Min_Temp_for_SpreadLabel = New System.Windows.Forms.Label()
        Flame_Spread_ParameterLabel = New System.Windows.Forms.Label()
        Cone_Data_FileLabel = New System.Windows.Forms.Label()
        Heat_of_CombustionLabel = New System.Windows.Forms.Label()
        Soot_YieldLabel = New System.Windows.Forms.Label()
        CO2_YieldLabel = New System.Windows.Forms.Label()
        H2O_YieldLabel = New System.Windows.Forms.Label()
        Comments2Label = New System.Windows.Forms.Label()
        CType(Me.ThermalDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Table1BindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Material_DescriptionLabel
        '
        Material_DescriptionLabel.AutoSize = True
        Material_DescriptionLabel.Location = New System.Drawing.Point(12, 9)
        Material_DescriptionLabel.Name = "Material_DescriptionLabel"
        Material_DescriptionLabel.Size = New System.Drawing.Size(103, 13)
        Material_DescriptionLabel.TabIndex = 2
        Material_DescriptionLabel.Text = "Material Description:"
        '
        'Thermal_ConductivityLabel
        '
        Thermal_ConductivityLabel.AutoSize = True
        Thermal_ConductivityLabel.Location = New System.Drawing.Point(370, 25)
        Thermal_ConductivityLabel.Name = "Thermal_ConductivityLabel"
        Thermal_ConductivityLabel.Size = New System.Drawing.Size(109, 13)
        Thermal_ConductivityLabel.TabIndex = 7
        Thermal_ConductivityLabel.Text = "Thermal Conductivity:"
        '
        'Specific_HeatLabel
        '
        Specific_HeatLabel.AutoSize = True
        Specific_HeatLabel.Location = New System.Drawing.Point(370, 51)
        Specific_HeatLabel.Name = "Specific_HeatLabel"
        Specific_HeatLabel.Size = New System.Drawing.Size(74, 13)
        Specific_HeatLabel.TabIndex = 9
        Specific_HeatLabel.Text = "Specific Heat:"
        '
        'DensityLabel
        '
        DensityLabel.AutoSize = True
        DensityLabel.Location = New System.Drawing.Point(370, 77)
        DensityLabel.Name = "DensityLabel"
        DensityLabel.Size = New System.Drawing.Size(45, 13)
        DensityLabel.TabIndex = 11
        DensityLabel.Text = "Density:"
        '
        'EmissivityLabel
        '
        EmissivityLabel.AutoSize = True
        EmissivityLabel.Location = New System.Drawing.Point(370, 103)
        EmissivityLabel.Name = "EmissivityLabel"
        EmissivityLabel.Size = New System.Drawing.Size(55, 13)
        EmissivityLabel.TabIndex = 13
        EmissivityLabel.Text = "Emissivity:"
        '
        'Min_Temp_for_SpreadLabel
        '
        Min_Temp_for_SpreadLabel.AutoSize = True
        Min_Temp_for_SpreadLabel.Location = New System.Drawing.Point(370, 129)
        Min_Temp_for_SpreadLabel.Name = "Min_Temp_for_SpreadLabel"
        Min_Temp_for_SpreadLabel.Size = New System.Drawing.Size(109, 13)
        Min_Temp_for_SpreadLabel.TabIndex = 15
        Min_Temp_for_SpreadLabel.Text = "Min Temp for Spread:"
        '
        'Flame_Spread_ParameterLabel
        '
        Flame_Spread_ParameterLabel.AutoSize = True
        Flame_Spread_ParameterLabel.Location = New System.Drawing.Point(370, 155)
        Flame_Spread_ParameterLabel.Name = "Flame_Spread_ParameterLabel"
        Flame_Spread_ParameterLabel.Size = New System.Drawing.Size(126, 13)
        Flame_Spread_ParameterLabel.TabIndex = 17
        Flame_Spread_ParameterLabel.Text = "Flame Spread Parameter:"
        '
        'Cone_Data_FileLabel
        '
        Cone_Data_FileLabel.AutoSize = True
        Cone_Data_FileLabel.Location = New System.Drawing.Point(370, 181)
        Cone_Data_FileLabel.Name = "Cone_Data_FileLabel"
        Cone_Data_FileLabel.Size = New System.Drawing.Size(80, 13)
        Cone_Data_FileLabel.TabIndex = 19
        Cone_Data_FileLabel.Text = "Cone Data File:"
        '
        'Heat_of_CombustionLabel
        '
        Heat_of_CombustionLabel.AutoSize = True
        Heat_of_CombustionLabel.Location = New System.Drawing.Point(370, 207)
        Heat_of_CombustionLabel.Name = "Heat_of_CombustionLabel"
        Heat_of_CombustionLabel.Size = New System.Drawing.Size(103, 13)
        Heat_of_CombustionLabel.TabIndex = 21
        Heat_of_CombustionLabel.Text = "Heat of Combustion:"
        '
        'Soot_YieldLabel
        '
        Soot_YieldLabel.AutoSize = True
        Soot_YieldLabel.Location = New System.Drawing.Point(370, 233)
        Soot_YieldLabel.Name = "Soot_YieldLabel"
        Soot_YieldLabel.Size = New System.Drawing.Size(58, 13)
        Soot_YieldLabel.TabIndex = 25
        Soot_YieldLabel.Text = "Soot Yield:"
        '
        'CO2_YieldLabel
        '
        CO2_YieldLabel.AutoSize = True
        CO2_YieldLabel.Location = New System.Drawing.Point(370, 259)
        CO2_YieldLabel.Name = "CO2_YieldLabel"
        CO2_YieldLabel.Size = New System.Drawing.Size(57, 13)
        CO2_YieldLabel.TabIndex = 27
        CO2_YieldLabel.Text = "CO2 Yield:"
        '
        'H2O_YieldLabel
        '
        H2O_YieldLabel.AutoSize = True
        H2O_YieldLabel.Location = New System.Drawing.Point(370, 285)
        H2O_YieldLabel.Name = "H2O_YieldLabel"
        H2O_YieldLabel.Size = New System.Drawing.Size(58, 13)
        H2O_YieldLabel.TabIndex = 29
        H2O_YieldLabel.Text = "H2O Yield:"
        '
        'Comments2Label
        '
        Comments2Label.AutoSize = True
        Comments2Label.Location = New System.Drawing.Point(370, 337)
        Comments2Label.Name = "Comments2Label"
        Comments2Label.Size = New System.Drawing.Size(59, 13)
        Comments2Label.TabIndex = 33
        Comments2Label.Text = "Comments:"
        '
        'ThermalDataSet
        '
        Me.ThermalDataSet.DataSetName = "thermalDataSet"
        Me.ThermalDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'Table1BindingSource
        '
        Me.Table1BindingSource.DataMember = "Table1"
        Me.Table1BindingSource.DataSource = Me.ThermalDataSet
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
        'Material_DescriptionListBox
        '
        Me.Material_DescriptionListBox.DataSource = Me.Table1BindingSource
        Me.Material_DescriptionListBox.DisplayMember = "Material Description"
        Me.Material_DescriptionListBox.FormattingEnabled = True
        Me.Material_DescriptionListBox.Location = New System.Drawing.Point(15, 25)
        Me.Material_DescriptionListBox.Name = "Material_DescriptionListBox"
        Me.Material_DescriptionListBox.Size = New System.Drawing.Size(331, 290)
        Me.Material_DescriptionListBox.TabIndex = 3
        '
        'Thermal_ConductivityTextBox
        '
        Me.Thermal_ConductivityTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Thermal Conductivity", True))
        Me.Thermal_ConductivityTextBox.Location = New System.Drawing.Point(502, 22)
        Me.Thermal_ConductivityTextBox.Name = "Thermal_ConductivityTextBox"
        Me.Thermal_ConductivityTextBox.ReadOnly = True
        Me.Thermal_ConductivityTextBox.Size = New System.Drawing.Size(120, 20)
        Me.Thermal_ConductivityTextBox.TabIndex = 8
        '
        'Specific_HeatTextBox
        '
        Me.Specific_HeatTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Specific Heat", True))
        Me.Specific_HeatTextBox.Location = New System.Drawing.Point(502, 48)
        Me.Specific_HeatTextBox.Name = "Specific_HeatTextBox"
        Me.Specific_HeatTextBox.ReadOnly = True
        Me.Specific_HeatTextBox.Size = New System.Drawing.Size(120, 20)
        Me.Specific_HeatTextBox.TabIndex = 10
        '
        'DensityTextBox
        '
        Me.DensityTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Density", True))
        Me.DensityTextBox.Location = New System.Drawing.Point(502, 74)
        Me.DensityTextBox.Name = "DensityTextBox"
        Me.DensityTextBox.ReadOnly = True
        Me.DensityTextBox.Size = New System.Drawing.Size(120, 20)
        Me.DensityTextBox.TabIndex = 12
        '
        'EmissivityTextBox
        '
        Me.EmissivityTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Emissivity", True))
        Me.EmissivityTextBox.Location = New System.Drawing.Point(502, 100)
        Me.EmissivityTextBox.Name = "EmissivityTextBox"
        Me.EmissivityTextBox.ReadOnly = True
        Me.EmissivityTextBox.Size = New System.Drawing.Size(120, 20)
        Me.EmissivityTextBox.TabIndex = 14
        '
        'Min_Temp_for_SpreadTextBox
        '
        Me.Min_Temp_for_SpreadTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Min Temp for Spread", True))
        Me.Min_Temp_for_SpreadTextBox.Location = New System.Drawing.Point(502, 126)
        Me.Min_Temp_for_SpreadTextBox.Name = "Min_Temp_for_SpreadTextBox"
        Me.Min_Temp_for_SpreadTextBox.ReadOnly = True
        Me.Min_Temp_for_SpreadTextBox.Size = New System.Drawing.Size(120, 20)
        Me.Min_Temp_for_SpreadTextBox.TabIndex = 16
        '
        'Flame_Spread_ParameterTextBox
        '
        Me.Flame_Spread_ParameterTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Flame Spread Parameter", True))
        Me.Flame_Spread_ParameterTextBox.Location = New System.Drawing.Point(502, 152)
        Me.Flame_Spread_ParameterTextBox.Name = "Flame_Spread_ParameterTextBox"
        Me.Flame_Spread_ParameterTextBox.ReadOnly = True
        Me.Flame_Spread_ParameterTextBox.Size = New System.Drawing.Size(120, 20)
        Me.Flame_Spread_ParameterTextBox.TabIndex = 18
        '
        'Cone_Data_FileTextBox
        '
        Me.Cone_Data_FileTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Cone Data File", True))
        Me.Cone_Data_FileTextBox.Location = New System.Drawing.Point(502, 178)
        Me.Cone_Data_FileTextBox.Name = "Cone_Data_FileTextBox"
        Me.Cone_Data_FileTextBox.ReadOnly = True
        Me.Cone_Data_FileTextBox.Size = New System.Drawing.Size(120, 20)
        Me.Cone_Data_FileTextBox.TabIndex = 20
        '
        'Heat_of_CombustionTextBox
        '
        Me.Heat_of_CombustionTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Heat of Combustion", True))
        Me.Heat_of_CombustionTextBox.Location = New System.Drawing.Point(502, 204)
        Me.Heat_of_CombustionTextBox.Name = "Heat_of_CombustionTextBox"
        Me.Heat_of_CombustionTextBox.ReadOnly = True
        Me.Heat_of_CombustionTextBox.Size = New System.Drawing.Size(120, 20)
        Me.Heat_of_CombustionTextBox.TabIndex = 22
        '
        'Soot_YieldTextBox
        '
        Me.Soot_YieldTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Soot Yield", True))
        Me.Soot_YieldTextBox.Location = New System.Drawing.Point(502, 230)
        Me.Soot_YieldTextBox.Name = "Soot_YieldTextBox"
        Me.Soot_YieldTextBox.ReadOnly = True
        Me.Soot_YieldTextBox.Size = New System.Drawing.Size(120, 20)
        Me.Soot_YieldTextBox.TabIndex = 26
        '
        'CO2_YieldTextBox
        '
        Me.CO2_YieldTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "CO2 Yield", True))
        Me.CO2_YieldTextBox.Location = New System.Drawing.Point(502, 256)
        Me.CO2_YieldTextBox.Name = "CO2_YieldTextBox"
        Me.CO2_YieldTextBox.ReadOnly = True
        Me.CO2_YieldTextBox.Size = New System.Drawing.Size(120, 20)
        Me.CO2_YieldTextBox.TabIndex = 28
        '
        'H2O_YieldTextBox
        '
        Me.H2O_YieldTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "H2O Yield", True))
        Me.H2O_YieldTextBox.Location = New System.Drawing.Point(502, 282)
        Me.H2O_YieldTextBox.Name = "H2O_YieldTextBox"
        Me.H2O_YieldTextBox.ReadOnly = True
        Me.H2O_YieldTextBox.Size = New System.Drawing.Size(120, 20)
        Me.H2O_YieldTextBox.TabIndex = 30
        '
        'Comments2TextBox
        '
        Me.Comments2TextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.Table1BindingSource, "Comments2", True))
        Me.Comments2TextBox.Location = New System.Drawing.Point(502, 308)
        Me.Comments2TextBox.Multiline = True
        Me.Comments2TextBox.Name = "Comments2TextBox"
        Me.Comments2TextBox.ReadOnly = True
        Me.Comments2TextBox.Size = New System.Drawing.Size(120, 46)
        Me.Comments2TextBox.TabIndex = 34
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(190, 337)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 35
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdSelectMat
        '
        Me.cmdSelectMat.Location = New System.Drawing.Point(271, 337)
        Me.cmdSelectMat.Name = "cmdSelectMat"
        Me.cmdSelectMat.Size = New System.Drawing.Size(75, 23)
        Me.cmdSelectMat.TabIndex = 36
        Me.cmdSelectMat.Text = "Select"
        Me.cmdSelectMat.UseVisualStyleBackColor = True
        '
        'frmMatSelect
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(645, 368)
        Me.Controls.Add(Me.cmdSelectMat)
        Me.Controls.Add(Me.cmdCancel)
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
        Me.Controls.Add(Soot_YieldLabel)
        Me.Controls.Add(Me.Soot_YieldTextBox)
        Me.Controls.Add(CO2_YieldLabel)
        Me.Controls.Add(Me.CO2_YieldTextBox)
        Me.Controls.Add(H2O_YieldLabel)
        Me.Controls.Add(Me.H2O_YieldTextBox)
        Me.Controls.Add(Comments2Label)
        Me.Controls.Add(Me.Comments2TextBox)
        Me.Controls.Add(Material_DescriptionLabel)
        Me.Controls.Add(Me.Material_DescriptionListBox)
        Me.Name = "frmMatSelect"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Pick Material"
        Me.TopMost = True
        CType(Me.ThermalDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Table1BindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ThermalDataSet As thermalDataSet
    Friend WithEvents Table1BindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Table1TableAdapter As thermalDataSetTableAdapters.Table1TableAdapter
    Friend WithEvents TableAdapterManager As thermalDataSetTableAdapters.TableAdapterManager
    Friend WithEvents Material_DescriptionListBox As System.Windows.Forms.ListBox
    Friend WithEvents Thermal_ConductivityTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Specific_HeatTextBox As System.Windows.Forms.TextBox
    Friend WithEvents DensityTextBox As System.Windows.Forms.TextBox
    Friend WithEvents EmissivityTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Min_Temp_for_SpreadTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Flame_Spread_ParameterTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Cone_Data_FileTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Heat_of_CombustionTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Soot_YieldTextBox As System.Windows.Forms.TextBox
    Friend WithEvents CO2_YieldTextBox As System.Windows.Forms.TextBox
    Friend WithEvents H2O_YieldTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Comments2TextBox As System.Windows.Forms.TextBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdSelectMat As Button
End Class
