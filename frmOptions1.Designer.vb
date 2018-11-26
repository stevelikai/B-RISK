<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmOptions1
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmdOK As System.Windows.Forms.Button
    Public WithEvents txtDescription As System.Windows.Forms.TextBox
	Public WithEvents txtJobNumber As System.Windows.Forms.TextBox
    Public WithEvents lblDescription As System.Windows.Forms.Label
	Public WithEvents lblJobNumber As System.Windows.Forms.Label
    Public WithEvents _Frame1_0 As System.Windows.Forms.GroupBox
	Public WithEvents _SSTab2_TabPage0 As System.Windows.Forms.TabPage
	Public WithEvents optCCFM As System.Windows.Forms.RadioButton
	Public WithEvents optCFAST As System.Windows.Forms.RadioButton
	Public WithEvents Frame28 As System.Windows.Forms.GroupBox
	Public WithEvents chkWallFlowDisable As System.Windows.Forms.CheckBox
	Public WithEvents Frame23 As System.Windows.Forms.GroupBox
	Public WithEvents optJET As System.Windows.Forms.RadioButton
	Public WithEvents optAlpert As System.Windows.Forms.RadioButton
	Public WithEvents Frame22 As System.Windows.Forms.GroupBox
	Public WithEvents optOneLayer As System.Windows.Forms.RadioButton
	Public WithEvents optTwoLayer As System.Windows.Forms.RadioButton
	Public WithEvents lstRoomZone As System.Windows.Forms.ComboBox
	Public WithEvents Label11 As System.Windows.Forms.Label
	Public WithEvents Frame21 As System.Windows.Forms.GroupBox
	Public WithEvents optStrongPlume As System.Windows.Forms.RadioButton
	Public WithEvents optMcCaffrey As System.Windows.Forms.RadioButton
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents _SSTab2_TabPage1 As System.Windows.Forms.TabPage
	Public WithEvents txtnC As System.Windows.Forms.TextBox
	Public WithEvents txtnH As System.Windows.Forms.TextBox
	Public WithEvents txtnO As System.Windows.Forms.TextBox
	Public WithEvents txtnN As System.Windows.Forms.TextBox
	Public WithEvents chkHCNcalc As System.Windows.Forms.CheckBox
	Public WithEvents txtSootAlpha As System.Windows.Forms.TextBox
	Public WithEvents txtSootEps As System.Windows.Forms.TextBox
	Public WithEvents cboABSCoeff As System.Windows.Forms.ComboBox
    Public WithEvents txtEmissionCoefficient As System.Windows.Forms.TextBox
    Public WithEvents Label23 As System.Windows.Forms.Label
	Public WithEvents Label24 As System.Windows.Forms.Label
	Public WithEvents Label25 As System.Windows.Forms.Label
	Public WithEvents Label26 As System.Windows.Forms.Label
	Public WithEvents Label27 As System.Windows.Forms.Label
	Public WithEvents Label22 As System.Windows.Forms.Label
	Public WithEvents Label21 As System.Windows.Forms.Label
	Public WithEvents Label18 As System.Windows.Forms.Label
    Public WithEvents lblEmissionCoefficient As System.Windows.Forms.Label
    Public WithEvents Fr15combparam As System.Windows.Forms.GroupBox
	Public WithEvents _SSTab2_TabPage2 As System.Windows.Forms.TabPage
    Public WithEvents txtRelativeHumidity As System.Windows.Forms.TextBox
	Public WithEvents txtExteriorTemp As System.Windows.Forms.TextBox
	Public WithEvents txtInteriorTemp As System.Windows.Forms.TextBox
	Public WithEvents lblRelativeHumidity As System.Windows.Forms.Label
	Public WithEvents lblExteriorTemp As System.Windows.Forms.Label
	Public WithEvents lblInteriorTemp As System.Windows.Forms.Label
	Public WithEvents _Frame5_1 As System.Windows.Forms.GroupBox
	Public WithEvents _SSTab2_TabPage4 As System.Windows.Forms.TabPage
	Public WithEvents optReflectiveSign As System.Windows.Forms.RadioButton
	Public WithEvents optIlluminatedSign As System.Windows.Forms.RadioButton
	Public WithEvents Frame15 As System.Windows.Forms.GroupBox
	Public WithEvents txtMonitorHeight As System.Windows.Forms.TextBox
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Frame7 As System.Windows.Forms.GroupBox
	Public WithEvents optHeavyWork As System.Windows.Forms.RadioButton
	Public WithEvents optLightWork As System.Windows.Forms.RadioButton
	Public WithEvents optAtRest As System.Windows.Forms.RadioButton
    Public WithEvents GB_activity As System.Windows.Forms.GroupBox
    Public WithEvents txtTemp As System.Windows.Forms.TextBox
    Public WithEvents txtConvect As System.Windows.Forms.TextBox
    Public WithEvents txtTarget As System.Windows.Forms.TextBox
    Public WithEvents txtVisibility As System.Windows.Forms.TextBox
    Public WithEvents txtFED As System.Windows.Forms.TextBox
    Public WithEvents lblConvect As System.Windows.Forms.Label
    Public WithEvents lblTarget As System.Windows.Forms.Label
    Public WithEvents lblVisibility As System.Windows.Forms.Label
    Public WithEvents lblFED As System.Windows.Forms.Label
    Public WithEvents lblTemp As System.Windows.Forms.Label
    Public WithEvents Frame6 As System.Windows.Forms.GroupBox
    Public WithEvents _SSTab2_TabPage5 As System.Windows.Forms.TabPage
    Public WithEvents cmdQuintiere As System.Windows.Forms.Button
    Public WithEvents txtCeilingHeatFlux As System.Windows.Forms.TextBox
    Public WithEvents txtWallHeatFlux As System.Windows.Forms.TextBox
    Public WithEvents lblCeilingHeatFLux As System.Windows.Forms.Label
    Public WithEvents lblHeatFluxWall As System.Windows.Forms.Label
    Public WithEvents _Frame14_0 As System.Windows.Forms.GroupBox
    Public WithEvents txtFlameLengthPower As System.Windows.Forms.TextBox
    Public WithEvents txtFlameAreaConstant As System.Windows.Forms.TextBox
    Public WithEvents txtBurnerWidth As System.Windows.Forms.TextBox
    Public WithEvents lblFlameLengthPower As System.Windows.Forms.Label
    Public WithEvents lblFlameAreaConstant As System.Windows.Forms.Label
    Public WithEvents lblBurnerWidth As System.Windows.Forms.Label
    Public WithEvents Frame13 As System.Windows.Forms.GroupBox
    Public WithEvents optQuintiere As System.Windows.Forms.RadioButton
    Public WithEvents optKarlsson As System.Windows.Forms.RadioButton
    Public WithEvents optRCNone As System.Windows.Forms.RadioButton
    Public WithEvents Frame12 As System.Windows.Forms.GroupBox
    Public WithEvents _SSTab2_TabPage6 As System.Windows.Forms.TabPage
    Public WithEvents lstTimeStep As System.Windows.Forms.ComboBox
    Public WithEvents txtErrorControl As System.Windows.Forms.TextBox
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Label13 As System.Windows.Forms.Label
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents optLUdecom As System.Windows.Forms.RadioButton
    Public WithEvents optGaussJor As System.Windows.Forms.RadioButton
    Public WithEvents _txtNodes_0 As System.Windows.Forms.TextBox
    Public WithEvents _txtNodes_1 As System.Windows.Forms.TextBox
    Public WithEvents _txtNodes_2 As System.Windows.Forms.TextBox
    Public WithEvents Label14 As System.Windows.Forms.Label
    Public WithEvents Label15 As System.Windows.Forms.Label
    Public WithEvents Label16 As System.Windows.Forms.Label
    Public WithEvents Label17 As System.Windows.Forms.Label
    Public WithEvents Frame19 As System.Windows.Forms.GroupBox
    Public WithEvents _SSTab2_TabPage7 As System.Windows.Forms.TabPage
    Public WithEvents ChkFRR As System.Windows.Forms.CheckBox
    Public WithEvents OptPreFlashover As System.Windows.Forms.RadioButton
    Public WithEvents optPostFlashover As System.Windows.Forms.RadioButton
    Public WithEvents Frame17 As System.Windows.Forms.GroupBox
    Public WithEvents txtStickSpacing As System.Windows.Forms.TextBox
    Public WithEvents chkModGER As System.Windows.Forms.CheckBox
    Public WithEvents txtFuelThickness As System.Windows.Forms.TextBox
    Public WithEvents txtHOCFuel As System.Windows.Forms.TextBox
    Public WithEvents Label20 As System.Windows.Forms.Label
    Public WithEvents Label12 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Frame18 As System.Windows.Forms.GroupBox
    Public WithEvents Frame16 As System.Windows.Forms.GroupBox
    Public WithEvents _SSTab2_TabPage8 As System.Windows.Forms.TabPage
    Public WithEvents optFuelLimited As System.Windows.Forms.RadioButton
    Public WithEvents optFuelNoLimit As System.Windows.Forms.RadioButton
    Public WithEvents Frame25 As System.Windows.Forms.GroupBox
    Public WithEvents txtFuelSurfaceArea As System.Windows.Forms.TextBox
    Public WithEvents txtLHoG As System.Windows.Forms.TextBox
    Public WithEvents optEnhanceOff As System.Windows.Forms.RadioButton
    Public WithEvents optEnhanceOn As System.Windows.Forms.RadioButton
    Public WithEvents lblFuelSurfaceArea As System.Windows.Forms.Label
    Public WithEvents lblLHoG As System.Windows.Forms.Label
    Public WithEvents Frame20 As System.Windows.Forms.GroupBox
    Public WithEvents _SSTab2_TabPage9 As System.Windows.Forms.TabPage
    Public WithEvents txtpreCO As System.Windows.Forms.TextBox
    Public WithEvents txtpostCO As System.Windows.Forms.TextBox
    Public WithEvents optCOauto As System.Windows.Forms.RadioButton
    Public WithEvents optCOman As System.Windows.Forms.RadioButton
    Public WithEvents Label19 As System.Windows.Forms.Label
    Public WithEvents Label28 As System.Windows.Forms.Label
    Public WithEvents Frame26 As System.Windows.Forms.GroupBox
    Public WithEvents txtpreSoot As System.Windows.Forms.TextBox
    Public WithEvents txtPostSoot As System.Windows.Forms.TextBox
    Public WithEvents optSootman As System.Windows.Forms.RadioButton
    Public WithEvents optSootauto As System.Windows.Forms.RadioButton
    Public WithEvents Label30 As System.Windows.Forms.Label
    Public WithEvents Label29 As System.Windows.Forms.Label
    Public WithEvents Frame27 As System.Windows.Forms.GroupBox
    Public WithEvents _SSTab2_TabPage10 As System.Windows.Forms.TabPage
    Public WithEvents SSTab2 As System.Windows.Forms.TabControl
    Public WithEvents Frame1 As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents Frame14 As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents Frame5 As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents txtNodes As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.SSTab2 = New System.Windows.Forms.TabControl()
        Me._SSTab2_TabPage0 = New System.Windows.Forms.TabPage()
        Me._Frame1_0 = New System.Windows.Forms.GroupBox()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.txtJobNumber = New System.Windows.Forms.TextBox()
        Me.lblDescription = New System.Windows.Forms.Label()
        Me.lblJobNumber = New System.Windows.Forms.Label()
        Me._SSTab2_TabPage1 = New System.Windows.Forms.TabPage()
        Me.Frame28 = New System.Windows.Forms.GroupBox()
        Me.optCCFM = New System.Windows.Forms.RadioButton()
        Me.optCFAST = New System.Windows.Forms.RadioButton()
        Me.Frame23 = New System.Windows.Forms.GroupBox()
        Me.chkWallFlowDisable = New System.Windows.Forms.CheckBox()
        Me.Frame22 = New System.Windows.Forms.GroupBox()
        Me.optJET = New System.Windows.Forms.RadioButton()
        Me.optAlpert = New System.Windows.Forms.RadioButton()
        Me.Frame21 = New System.Windows.Forms.GroupBox()
        Me.optOneLayer = New System.Windows.Forms.RadioButton()
        Me.optTwoLayer = New System.Windows.Forms.RadioButton()
        Me.lstRoomZone = New System.Windows.Forms.ComboBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.optStrongPlume = New System.Windows.Forms.RadioButton()
        Me.optMcCaffrey = New System.Windows.Forms.RadioButton()
        Me._SSTab2_TabPage2 = New System.Windows.Forms.TabPage()
        Me.Fr15combparam = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtStoich = New System.Windows.Forms.TextBox()
        Me.txtnC = New System.Windows.Forms.TextBox()
        Me.txtnH = New System.Windows.Forms.TextBox()
        Me.txtnO = New System.Windows.Forms.TextBox()
        Me.txtnN = New System.Windows.Forms.TextBox()
        Me.chkHCNcalc = New System.Windows.Forms.CheckBox()
        Me.txtSootAlpha = New System.Windows.Forms.TextBox()
        Me.txtSootEps = New System.Windows.Forms.TextBox()
        Me.cboABSCoeff = New System.Windows.Forms.ComboBox()
        Me.txtEmissionCoefficient = New System.Windows.Forms.TextBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.lblEmissionCoefficient = New System.Windows.Forms.Label()
        Me._SSTab2_TabPage4 = New System.Windows.Forms.TabPage()
        Me._Frame5_1 = New System.Windows.Forms.GroupBox()
        Me.cmdDist_RH = New System.Windows.Forms.Button()
        Me.cmdDist_exteriortemp = New System.Windows.Forms.Button()
        Me.cmdDist_interiortemp = New System.Windows.Forms.Button()
        Me.txtRelativeHumidity = New System.Windows.Forms.TextBox()
        Me.txtExteriorTemp = New System.Windows.Forms.TextBox()
        Me.txtInteriorTemp = New System.Windows.Forms.TextBox()
        Me.lblRelativeHumidity = New System.Windows.Forms.Label()
        Me.lblExteriorTemp = New System.Windows.Forms.Label()
        Me.lblInteriorTemp = New System.Windows.Forms.Label()
        Me._SSTab2_TabPage5 = New System.Windows.Forms.TabPage()
        Me.GB_gasmodel = New System.Windows.Forms.GroupBox()
        Me.optFEDGeneral = New System.Windows.Forms.RadioButton()
        Me.optFEDCO = New System.Windows.Forms.RadioButton()
        Me.Frame15 = New System.Windows.Forms.GroupBox()
        Me.optReflectiveSign = New System.Windows.Forms.RadioButton()
        Me.optIlluminatedSign = New System.Windows.Forms.RadioButton()
        Me.Frame7 = New System.Windows.Forms.GroupBox()
        Me.txtMonitorHeight = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.GB_activity = New System.Windows.Forms.GroupBox()
        Me.optHeavyWork = New System.Windows.Forms.RadioButton()
        Me.optLightWork = New System.Windows.Forms.RadioButton()
        Me.optAtRest = New System.Windows.Forms.RadioButton()
        Me.Frame6 = New System.Windows.Forms.GroupBox()
        Me.txtTemp = New System.Windows.Forms.TextBox()
        Me.txtConvect = New System.Windows.Forms.TextBox()
        Me.txtTarget = New System.Windows.Forms.TextBox()
        Me.txtVisibility = New System.Windows.Forms.TextBox()
        Me.txtFED = New System.Windows.Forms.TextBox()
        Me.lblConvect = New System.Windows.Forms.Label()
        Me.lblTarget = New System.Windows.Forms.Label()
        Me.lblVisibility = New System.Windows.Forms.Label()
        Me.lblFED = New System.Windows.Forms.Label()
        Me.lblTemp = New System.Windows.Forms.Label()
        Me._SSTab2_TabPage6 = New System.Windows.Forms.TabPage()
        Me.cmdQuintiere = New System.Windows.Forms.Button()
        Me._Frame14_0 = New System.Windows.Forms.GroupBox()
        Me.txtCeilingHeatFlux = New System.Windows.Forms.TextBox()
        Me.txtWallHeatFlux = New System.Windows.Forms.TextBox()
        Me.lblCeilingHeatFLux = New System.Windows.Forms.Label()
        Me.lblHeatFluxWall = New System.Windows.Forms.Label()
        Me.Frame13 = New System.Windows.Forms.GroupBox()
        Me.txtFlameLengthPower = New System.Windows.Forms.TextBox()
        Me.txtFlameAreaConstant = New System.Windows.Forms.TextBox()
        Me.txtBurnerWidth = New System.Windows.Forms.TextBox()
        Me.lblFlameLengthPower = New System.Windows.Forms.Label()
        Me.lblFlameAreaConstant = New System.Windows.Forms.Label()
        Me.lblBurnerWidth = New System.Windows.Forms.Label()
        Me.Frame12 = New System.Windows.Forms.GroupBox()
        Me.optQuintiere = New System.Windows.Forms.RadioButton()
        Me.optKarlsson = New System.Windows.Forms.RadioButton()
        Me.optRCNone = New System.Windows.Forms.RadioButton()
        Me._SSTab2_TabPage7 = New System.Windows.Forms.TabPage()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.txtErrorControlVentFlow = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lstTimeStep = New System.Windows.Forms.ComboBox()
        Me.txtErrorControl = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Frame19 = New System.Windows.Forms.GroupBox()
        Me.optLUdecom = New System.Windows.Forms.RadioButton()
        Me.optGaussJor = New System.Windows.Forms.RadioButton()
        Me._txtNodes_0 = New System.Windows.Forms.TextBox()
        Me._txtNodes_1 = New System.Windows.Forms.TextBox()
        Me._txtNodes_2 = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me._SSTab2_TabPage8 = New System.Windows.Forms.TabPage()
        Me.Frame16 = New System.Windows.Forms.GroupBox()
        Me.GB_flashovercriteria = New System.Windows.Forms.GroupBox()
        Me.optFOtemp = New System.Windows.Forms.RadioButton()
        Me.optFOflux = New System.Windows.Forms.RadioButton()
        Me.ChkFRR = New System.Windows.Forms.CheckBox()
        Me.Frame17 = New System.Windows.Forms.GroupBox()
        Me.cmdWoodOption = New System.Windows.Forms.Button()
        Me.OptPreFlashover = New System.Windows.Forms.RadioButton()
        Me.optPostFlashover = New System.Windows.Forms.RadioButton()
        Me.Frame18 = New System.Windows.Forms.GroupBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtCribheight = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtExcessFuelFactor = New System.Windows.Forms.TextBox()
        Me.cmdDist_HOC = New System.Windows.Forms.Button()
        Me.txtStickSpacing = New System.Windows.Forms.TextBox()
        Me.chkModGER = New System.Windows.Forms.CheckBox()
        Me.txtFuelThickness = New System.Windows.Forms.TextBox()
        Me.txtHOCFuel = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me._SSTab2_TabPage10 = New System.Windows.Forms.TabPage()
        Me.Frame26 = New System.Windows.Forms.GroupBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.txtpreCO = New System.Windows.Forms.TextBox()
        Me.txtpostCO = New System.Windows.Forms.TextBox()
        Me.optCOauto = New System.Windows.Forms.RadioButton()
        Me.optCOman = New System.Windows.Forms.RadioButton()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.Frame27 = New System.Windows.Forms.GroupBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.txtpreSoot = New System.Windows.Forms.TextBox()
        Me.txtPostSoot = New System.Windows.Forms.TextBox()
        Me.optSootman = New System.Windows.Forms.RadioButton()
        Me.optSootauto = New System.Windows.Forms.RadioButton()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.Label29 = New System.Windows.Forms.Label()
        Me._SSTab2_TabPage9 = New System.Windows.Forms.TabPage()
        Me.Frame25 = New System.Windows.Forms.GroupBox()
        Me.optFuelLimited = New System.Windows.Forms.RadioButton()
        Me.optFuelNoLimit = New System.Windows.Forms.RadioButton()
        Me.Frame20 = New System.Windows.Forms.GroupBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.txtFuelSurfaceArea = New System.Windows.Forms.TextBox()
        Me.txtLHoG = New System.Windows.Forms.TextBox()
        Me.optEnhanceOff = New System.Windows.Forms.RadioButton()
        Me.optEnhanceOn = New System.Windows.Forms.RadioButton()
        Me.lblFuelSurfaceArea = New System.Windows.Forms.Label()
        Me.lblLHoG = New System.Windows.Forms.Label()
        Me.Frame1 = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.Frame14 = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.Frame5 = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.txtNodes = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.SSTab2.SuspendLayout()
        Me._SSTab2_TabPage0.SuspendLayout()
        Me._Frame1_0.SuspendLayout()
        Me._SSTab2_TabPage1.SuspendLayout()
        Me.Frame28.SuspendLayout()
        Me.Frame23.SuspendLayout()
        Me.Frame22.SuspendLayout()
        Me.Frame21.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me._SSTab2_TabPage2.SuspendLayout()
        Me.Fr15combparam.SuspendLayout()
        Me._SSTab2_TabPage4.SuspendLayout()
        Me._Frame5_1.SuspendLayout()
        Me._SSTab2_TabPage5.SuspendLayout()
        Me.GB_gasmodel.SuspendLayout()
        Me.Frame15.SuspendLayout()
        Me.Frame7.SuspendLayout()
        Me.GB_activity.SuspendLayout()
        Me.Frame6.SuspendLayout()
        Me._SSTab2_TabPage6.SuspendLayout()
        Me._Frame14_0.SuspendLayout()
        Me.Frame13.SuspendLayout()
        Me.Frame12.SuspendLayout()
        Me._SSTab2_TabPage7.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame19.SuspendLayout()
        Me._SSTab2_TabPage8.SuspendLayout()
        Me.Frame16.SuspendLayout()
        Me.GB_flashovercriteria.SuspendLayout()
        Me.Frame17.SuspendLayout()
        Me.Frame18.SuspendLayout()
        Me._SSTab2_TabPage10.SuspendLayout()
        Me.Frame26.SuspendLayout()
        Me.Frame27.SuspendLayout()
        Me._SSTab2_TabPage9.SuspendLayout()
        Me.Frame25.SuspendLayout()
        Me.Frame20.SuspendLayout()
        CType(Me.Frame1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Frame14, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Frame5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtNodes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Location = New System.Drawing.Point(444, 399)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(65, 25)
        Me.cmdOK.TabIndex = 0
        Me.cmdOK.Text = "OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'SSTab2
        '
        Me.SSTab2.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.SSTab2.Controls.Add(Me._SSTab2_TabPage0)
        Me.SSTab2.Controls.Add(Me._SSTab2_TabPage1)
        Me.SSTab2.Controls.Add(Me._SSTab2_TabPage2)
        Me.SSTab2.Controls.Add(Me._SSTab2_TabPage4)
        Me.SSTab2.Controls.Add(Me._SSTab2_TabPage5)
        Me.SSTab2.Controls.Add(Me._SSTab2_TabPage6)
        Me.SSTab2.Controls.Add(Me._SSTab2_TabPage7)
        Me.SSTab2.Controls.Add(Me._SSTab2_TabPage8)
        Me.SSTab2.Controls.Add(Me._SSTab2_TabPage10)
        Me.SSTab2.Controls.Add(Me._SSTab2_TabPage9)
        Me.SSTab2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SSTab2.ItemSize = New System.Drawing.Size(42, 18)
        Me.SSTab2.Location = New System.Drawing.Point(8, 12)
        Me.SSTab2.Multiline = True
        Me.SSTab2.Name = "SSTab2"
        Me.SSTab2.SelectedIndex = 0
        Me.SSTab2.Size = New System.Drawing.Size(505, 381)
        Me.SSTab2.TabIndex = 5
        '
        '_SSTab2_TabPage0
        '
        Me._SSTab2_TabPage0.Controls.Add(Me._Frame1_0)
        Me._SSTab2_TabPage0.Location = New System.Drawing.Point(4, 43)
        Me._SSTab2_TabPage0.Name = "_SSTab2_TabPage0"
        Me._SSTab2_TabPage0.Size = New System.Drawing.Size(497, 334)
        Me._SSTab2_TabPage0.TabIndex = 0
        Me._SSTab2_TabPage0.Text = "General"
        Me._SSTab2_TabPage0.UseVisualStyleBackColor = True
        '
        '_Frame1_0
        '
        Me._Frame1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Frame1_0.Controls.Add(Me.txtDescription)
        Me._Frame1_0.Controls.Add(Me.txtJobNumber)
        Me._Frame1_0.Controls.Add(Me.lblDescription)
        Me._Frame1_0.Controls.Add(Me.lblJobNumber)
        Me._Frame1_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Frame1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.SetIndex(Me._Frame1_0, CType(0, Short))
        Me._Frame1_0.Location = New System.Drawing.Point(19, 21)
        Me._Frame1_0.Name = "_Frame1_0"
        Me._Frame1_0.Padding = New System.Windows.Forms.Padding(0)
        Me._Frame1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame1_0.Size = New System.Drawing.Size(361, 157)
        Me._Frame1_0.TabIndex = 28
        Me._Frame1_0.TabStop = False
        '
        'txtDescription
        '
        Me.txtDescription.AcceptsReturn = True
        Me.txtDescription.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescription.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescription.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDescription.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescription.Location = New System.Drawing.Point(26, 51)
        Me.txtDescription.MaxLength = 0
        Me.txtDescription.Multiline = True
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescription.Size = New System.Drawing.Size(265, 49)
        Me.txtDescription.TabIndex = 5
        Me.txtDescription.Text = "Describe your project here" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'txtJobNumber
        '
        Me.txtJobNumber.AcceptsReturn = True
        Me.txtJobNumber.BackColor = System.Drawing.SystemColors.Window
        Me.txtJobNumber.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtJobNumber.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtJobNumber.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtJobNumber.Location = New System.Drawing.Point(170, 107)
        Me.txtJobNumber.MaxLength = 0
        Me.txtJobNumber.Name = "txtJobNumber"
        Me.txtJobNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtJobNumber.Size = New System.Drawing.Size(121, 20)
        Me.txtJobNumber.TabIndex = 6
        '
        'lblDescription
        '
        Me.lblDescription.BackColor = System.Drawing.SystemColors.Control
        Me.lblDescription.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDescription.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDescription.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDescription.Location = New System.Drawing.Point(26, 27)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDescription.Size = New System.Drawing.Size(153, 17)
        Me.lblDescription.TabIndex = 36
        Me.lblDescription.Text = "Description of Current Project"
        '
        'lblJobNumber
        '
        Me.lblJobNumber.BackColor = System.Drawing.SystemColors.Control
        Me.lblJobNumber.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblJobNumber.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblJobNumber.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblJobNumber.Location = New System.Drawing.Point(58, 107)
        Me.lblJobNumber.Name = "lblJobNumber"
        Me.lblJobNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblJobNumber.Size = New System.Drawing.Size(105, 17)
        Me.lblJobNumber.TabIndex = 35
        Me.lblJobNumber.Text = "Job Number"
        Me.lblJobNumber.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_SSTab2_TabPage1
        '
        Me._SSTab2_TabPage1.Controls.Add(Me.Frame28)
        Me._SSTab2_TabPage1.Controls.Add(Me.Frame23)
        Me._SSTab2_TabPage1.Controls.Add(Me.Frame22)
        Me._SSTab2_TabPage1.Controls.Add(Me.Frame21)
        Me._SSTab2_TabPage1.Controls.Add(Me.Frame2)
        Me._SSTab2_TabPage1.Location = New System.Drawing.Point(4, 43)
        Me._SSTab2_TabPage1.Name = "_SSTab2_TabPage1"
        Me._SSTab2_TabPage1.Size = New System.Drawing.Size(497, 334)
        Me._SSTab2_TabPage1.TabIndex = 1
        Me._SSTab2_TabPage1.Text = "Model Physics"
        Me._SSTab2_TabPage1.UseVisualStyleBackColor = True
        '
        'Frame28
        '
        Me.Frame28.BackColor = System.Drawing.SystemColors.Control
        Me.Frame28.Controls.Add(Me.optCCFM)
        Me.Frame28.Controls.Add(Me.optCFAST)
        Me.Frame28.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame28.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame28.Location = New System.Drawing.Point(233, 99)
        Me.Frame28.Name = "Frame28"
        Me.Frame28.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame28.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame28.Size = New System.Drawing.Size(237, 65)
        Me.Frame28.TabIndex = 176
        Me.Frame28.TabStop = False
        Me.Frame28.Text = "Vent Flow Deposition Rules"
        Me.Frame28.Visible = False
        '
        'optCCFM
        '
        Me.optCCFM.BackColor = System.Drawing.SystemColors.Control
        Me.optCCFM.Cursor = System.Windows.Forms.Cursors.Default
        Me.optCCFM.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optCCFM.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optCCFM.Location = New System.Drawing.Point(104, 32)
        Me.optCCFM.Name = "optCCFM"
        Me.optCCFM.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optCCFM.Size = New System.Drawing.Size(57, 17)
        Me.optCCFM.TabIndex = 178
        Me.optCCFM.TabStop = True
        Me.optCCFM.Text = "ccfm"
        Me.optCCFM.UseVisualStyleBackColor = False
        '
        'optCFAST
        '
        Me.optCFAST.BackColor = System.Drawing.SystemColors.Control
        Me.optCFAST.Checked = True
        Me.optCFAST.Cursor = System.Windows.Forms.Cursors.Default
        Me.optCFAST.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optCFAST.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optCFAST.Location = New System.Drawing.Point(24, 32)
        Me.optCFAST.Name = "optCFAST"
        Me.optCFAST.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optCFAST.Size = New System.Drawing.Size(65, 17)
        Me.optCFAST.TabIndex = 177
        Me.optCFAST.TabStop = True
        Me.optCFAST.Text = "cfast"
        Me.optCFAST.UseVisualStyleBackColor = False
        '
        'Frame23
        '
        Me.Frame23.BackColor = System.Drawing.SystemColors.Control
        Me.Frame23.Controls.Add(Me.chkWallFlowDisable)
        Me.Frame23.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame23.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame23.Location = New System.Drawing.Point(233, 20)
        Me.Frame23.Name = "Frame23"
        Me.Frame23.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame23.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame23.Size = New System.Drawing.Size(237, 73)
        Me.Frame23.TabIndex = 154
        Me.Frame23.TabStop = False
        Me.Frame23.Text = "Other Physics"
        '
        'chkWallFlowDisable
        '
        Me.chkWallFlowDisable.BackColor = System.Drawing.SystemColors.Control
        Me.chkWallFlowDisable.Checked = True
        Me.chkWallFlowDisable.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkWallFlowDisable.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkWallFlowDisable.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkWallFlowDisable.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkWallFlowDisable.Location = New System.Drawing.Point(16, 16)
        Me.chkWallFlowDisable.Name = "chkWallFlowDisable"
        Me.chkWallFlowDisable.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkWallFlowDisable.Size = New System.Drawing.Size(210, 33)
        Me.chkWallFlowDisable.TabIndex = 155
        Me.chkWallFlowDisable.Text = "Disable wall convection flow"
        Me.chkWallFlowDisable.UseVisualStyleBackColor = False
        '
        'Frame22
        '
        Me.Frame22.BackColor = System.Drawing.SystemColors.Control
        Me.Frame22.Controls.Add(Me.optJET)
        Me.Frame22.Controls.Add(Me.optAlpert)
        Me.Frame22.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame22.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame22.Location = New System.Drawing.Point(17, 179)
        Me.Frame22.Name = "Frame22"
        Me.Frame22.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame22.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame22.Size = New System.Drawing.Size(210, 113)
        Me.Frame22.TabIndex = 121
        Me.Frame22.TabStop = False
        Me.Frame22.Text = "Ceiling Jet Model"
        '
        'optJET
        '
        Me.optJET.BackColor = System.Drawing.SystemColors.Control
        Me.optJET.Checked = True
        Me.optJET.Cursor = System.Windows.Forms.Cursors.Default
        Me.optJET.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optJET.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optJET.Location = New System.Drawing.Point(16, 64)
        Me.optJET.Name = "optJET"
        Me.optJET.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optJET.Size = New System.Drawing.Size(182, 33)
        Me.optJET.TabIndex = 123
        Me.optJET.TabStop = True
        Me.optJET.Text = "NIST JET model"
        Me.optJET.UseVisualStyleBackColor = False
        '
        'optAlpert
        '
        Me.optAlpert.BackColor = System.Drawing.SystemColors.Control
        Me.optAlpert.Cursor = System.Windows.Forms.Cursors.Default
        Me.optAlpert.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optAlpert.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optAlpert.Location = New System.Drawing.Point(16, 32)
        Me.optAlpert.Name = "optAlpert"
        Me.optAlpert.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optAlpert.Size = New System.Drawing.Size(182, 25)
        Me.optAlpert.TabIndex = 122
        Me.optAlpert.TabStop = True
        Me.optAlpert.Text = "Alpert's Unconfined Ceiling"
        Me.optAlpert.UseVisualStyleBackColor = False
        '
        'Frame21
        '
        Me.Frame21.BackColor = System.Drawing.SystemColors.Control
        Me.Frame21.Controls.Add(Me.optOneLayer)
        Me.Frame21.Controls.Add(Me.optTwoLayer)
        Me.Frame21.Controls.Add(Me.lstRoomZone)
        Me.Frame21.Controls.Add(Me.Label11)
        Me.Frame21.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame21.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame21.Location = New System.Drawing.Point(17, 20)
        Me.Frame21.Name = "Frame21"
        Me.Frame21.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame21.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame21.Size = New System.Drawing.Size(210, 153)
        Me.Frame21.TabIndex = 116
        Me.Frame21.TabStop = False
        Me.Frame21.Text = "Zone/Layer Assumptions"
        '
        'optOneLayer
        '
        Me.optOneLayer.BackColor = System.Drawing.SystemColors.Control
        Me.optOneLayer.Cursor = System.Windows.Forms.Cursors.Default
        Me.optOneLayer.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optOneLayer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optOneLayer.Location = New System.Drawing.Point(16, 104)
        Me.optOneLayer.Name = "optOneLayer"
        Me.optOneLayer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optOneLayer.Size = New System.Drawing.Size(191, 24)
        Me.optOneLayer.TabIndex = 119
        Me.optOneLayer.TabStop = True
        Me.optOneLayer.Text = "Model Room as Single Zone"
        Me.optOneLayer.UseVisualStyleBackColor = False
        '
        'optTwoLayer
        '
        Me.optTwoLayer.BackColor = System.Drawing.SystemColors.Control
        Me.optTwoLayer.Checked = True
        Me.optTwoLayer.Cursor = System.Windows.Forms.Cursors.Default
        Me.optTwoLayer.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optTwoLayer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optTwoLayer.Location = New System.Drawing.Point(16, 72)
        Me.optTwoLayer.Name = "optTwoLayer"
        Me.optTwoLayer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optTwoLayer.Size = New System.Drawing.Size(191, 26)
        Me.optTwoLayer.TabIndex = 118
        Me.optTwoLayer.TabStop = True
        Me.optTwoLayer.Text = "Two Zone/Layers (default)"
        Me.optTwoLayer.UseVisualStyleBackColor = False
        '
        'lstRoomZone
        '
        Me.lstRoomZone.BackColor = System.Drawing.SystemColors.Window
        Me.lstRoomZone.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstRoomZone.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstRoomZone.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstRoomZone.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstRoomZone.Location = New System.Drawing.Point(96, 32)
        Me.lstRoomZone.Name = "lstRoomZone"
        Me.lstRoomZone.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstRoomZone.Size = New System.Drawing.Size(73, 22)
        Me.lstRoomZone.TabIndex = 117
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Location = New System.Drawing.Point(24, 32)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(49, 25)
        Me.Label11.TabIndex = 120
        Me.Label11.Text = "Room #"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.optStrongPlume)
        Me.Frame2.Controls.Add(Me.optMcCaffrey)
        Me.Frame2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(233, 179)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(237, 113)
        Me.Frame2.TabIndex = 59
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Select Plume Model"
        Me.Frame2.Visible = False
        '
        'optStrongPlume
        '
        Me.optStrongPlume.BackColor = System.Drawing.SystemColors.Control
        Me.optStrongPlume.Cursor = System.Windows.Forms.Cursors.Default
        Me.optStrongPlume.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optStrongPlume.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optStrongPlume.Location = New System.Drawing.Point(16, 64)
        Me.optStrongPlume.Name = "optStrongPlume"
        Me.optStrongPlume.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optStrongPlume.Size = New System.Drawing.Size(153, 33)
        Me.optStrongPlume.TabIndex = 61
        Me.optStrongPlume.TabStop = True
        Me.optStrongPlume.Text = "Delichatsios"
        Me.optStrongPlume.UseVisualStyleBackColor = False
        '
        'optMcCaffrey
        '
        Me.optMcCaffrey.BackColor = System.Drawing.SystemColors.Control
        Me.optMcCaffrey.Checked = True
        Me.optMcCaffrey.Cursor = System.Windows.Forms.Cursors.Default
        Me.optMcCaffrey.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optMcCaffrey.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMcCaffrey.Location = New System.Drawing.Point(16, 32)
        Me.optMcCaffrey.Name = "optMcCaffrey"
        Me.optMcCaffrey.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optMcCaffrey.Size = New System.Drawing.Size(164, 26)
        Me.optMcCaffrey.TabIndex = 60
        Me.optMcCaffrey.TabStop = True
        Me.optMcCaffrey.Text = "McCaffrey (default)"
        Me.optMcCaffrey.UseVisualStyleBackColor = False
        '
        '_SSTab2_TabPage2
        '
        Me._SSTab2_TabPage2.Controls.Add(Me.Fr15combparam)
        Me._SSTab2_TabPage2.Location = New System.Drawing.Point(4, 43)
        Me._SSTab2_TabPage2.Name = "_SSTab2_TabPage2"
        Me._SSTab2_TabPage2.Size = New System.Drawing.Size(497, 334)
        Me._SSTab2_TabPage2.TabIndex = 2
        Me._SSTab2_TabPage2.Text = "Combustion Parameters"
        Me._SSTab2_TabPage2.UseVisualStyleBackColor = True
        '
        'Fr15combparam
        '
        Me.Fr15combparam.BackColor = System.Drawing.SystemColors.Control
        Me.Fr15combparam.Controls.Add(Me.Label4)
        Me.Fr15combparam.Controls.Add(Me.txtStoich)
        Me.Fr15combparam.Controls.Add(Me.txtnC)
        Me.Fr15combparam.Controls.Add(Me.txtnH)
        Me.Fr15combparam.Controls.Add(Me.txtnO)
        Me.Fr15combparam.Controls.Add(Me.txtnN)
        Me.Fr15combparam.Controls.Add(Me.chkHCNcalc)
        Me.Fr15combparam.Controls.Add(Me.txtSootAlpha)
        Me.Fr15combparam.Controls.Add(Me.txtSootEps)
        Me.Fr15combparam.Controls.Add(Me.cboABSCoeff)
        Me.Fr15combparam.Controls.Add(Me.txtEmissionCoefficient)
        Me.Fr15combparam.Controls.Add(Me.Label23)
        Me.Fr15combparam.Controls.Add(Me.Label24)
        Me.Fr15combparam.Controls.Add(Me.Label25)
        Me.Fr15combparam.Controls.Add(Me.Label26)
        Me.Fr15combparam.Controls.Add(Me.Label27)
        Me.Fr15combparam.Controls.Add(Me.Label22)
        Me.Fr15combparam.Controls.Add(Me.Label21)
        Me.Fr15combparam.Controls.Add(Me.Label18)
        Me.Fr15combparam.Controls.Add(Me.lblEmissionCoefficient)
        Me.Fr15combparam.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Fr15combparam.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Fr15combparam.Location = New System.Drawing.Point(19, 18)
        Me.Fr15combparam.Name = "Fr15combparam"
        Me.Fr15combparam.Padding = New System.Windows.Forms.Padding(0)
        Me.Fr15combparam.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Fr15combparam.Size = New System.Drawing.Size(409, 293)
        Me.Fr15combparam.TabIndex = 52
        Me.Fr15combparam.TabStop = False
        '
        'Label4
        '
        Me.Label4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(175, 77)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(146, 14)
        Me.Label4.TabIndex = 155
        Me.Label4.Text = "Stoichiometric air to fuel ratio"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtStoich
        '
        Me.txtStoich.Location = New System.Drawing.Point(333, 77)
        Me.txtStoich.Name = "txtStoich"
        Me.txtStoich.Size = New System.Drawing.Size(60, 20)
        Me.txtStoich.TabIndex = 154
        Me.txtStoich.Text = "1.4"
        '
        'txtnC
        '
        Me.txtnC.AcceptsReturn = True
        Me.txtnC.BackColor = System.Drawing.SystemColors.Window
        Me.txtnC.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtnC.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtnC.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtnC.Location = New System.Drawing.Point(48, 232)
        Me.txtnC.MaxLength = 0
        Me.txtnC.Name = "txtnC"
        Me.txtnC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtnC.Size = New System.Drawing.Size(33, 20)
        Me.txtnC.TabIndex = 148
        Me.txtnC.Text = "3"
        '
        'txtnH
        '
        Me.txtnH.AcceptsReturn = True
        Me.txtnH.BackColor = System.Drawing.SystemColors.Window
        Me.txtnH.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtnH.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtnH.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtnH.Location = New System.Drawing.Point(88, 232)
        Me.txtnH.MaxLength = 0
        Me.txtnH.Name = "txtnH"
        Me.txtnH.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtnH.Size = New System.Drawing.Size(33, 20)
        Me.txtnH.TabIndex = 147
        Me.txtnH.Text = "8"
        '
        'txtnO
        '
        Me.txtnO.AcceptsReturn = True
        Me.txtnO.BackColor = System.Drawing.SystemColors.Window
        Me.txtnO.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtnO.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtnO.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtnO.Location = New System.Drawing.Point(128, 232)
        Me.txtnO.MaxLength = 0
        Me.txtnO.Name = "txtnO"
        Me.txtnO.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtnO.Size = New System.Drawing.Size(33, 20)
        Me.txtnO.TabIndex = 146
        Me.txtnO.Text = "0"
        '
        'txtnN
        '
        Me.txtnN.AcceptsReturn = True
        Me.txtnN.BackColor = System.Drawing.SystemColors.Window
        Me.txtnN.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtnN.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtnN.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtnN.Location = New System.Drawing.Point(168, 232)
        Me.txtnN.MaxLength = 0
        Me.txtnN.Name = "txtnN"
        Me.txtnN.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtnN.Size = New System.Drawing.Size(33, 20)
        Me.txtnN.TabIndex = 145
        Me.txtnN.Text = "0"
        '
        'chkHCNcalc
        '
        Me.chkHCNcalc.BackColor = System.Drawing.SystemColors.Control
        Me.chkHCNcalc.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkHCNcalc.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkHCNcalc.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkHCNcalc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkHCNcalc.Location = New System.Drawing.Point(48, 264)
        Me.chkHCNcalc.Name = "chkHCNcalc"
        Me.chkHCNcalc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkHCNcalc.Size = New System.Drawing.Size(260, 25)
        Me.chkHCNcalc.TabIndex = 144
        Me.chkHCNcalc.Text = "Calculate HCN based on combustion chemistry"
        Me.chkHCNcalc.UseVisualStyleBackColor = False
        '
        'txtSootAlpha
        '
        Me.txtSootAlpha.AcceptsReturn = True
        Me.txtSootAlpha.BackColor = System.Drawing.SystemColors.Window
        Me.txtSootAlpha.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSootAlpha.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSootAlpha.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSootAlpha.Location = New System.Drawing.Point(336, 144)
        Me.txtSootAlpha.MaxLength = 0
        Me.txtSootAlpha.Name = "txtSootAlpha"
        Me.txtSootAlpha.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSootAlpha.Size = New System.Drawing.Size(57, 20)
        Me.txtSootAlpha.TabIndex = 141
        Me.txtSootAlpha.Text = "2.5"
        '
        'txtSootEps
        '
        Me.txtSootEps.AcceptsReturn = True
        Me.txtSootEps.BackColor = System.Drawing.SystemColors.Window
        Me.txtSootEps.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSootEps.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSootEps.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSootEps.Location = New System.Drawing.Point(336, 176)
        Me.txtSootEps.MaxLength = 0
        Me.txtSootEps.Name = "txtSootEps"
        Me.txtSootEps.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSootEps.Size = New System.Drawing.Size(57, 20)
        Me.txtSootEps.TabIndex = 140
        Me.txtSootEps.Text = "1.2"
        '
        'cboABSCoeff
        '
        Me.cboABSCoeff.BackColor = System.Drawing.SystemColors.Window
        Me.cboABSCoeff.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboABSCoeff.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboABSCoeff.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboABSCoeff.Items.AddRange(New Object() {"wood", "plexiglas", "polystyrene", "methane", "heptane", "propane", "polyurethane foam flexible", "ethanol", "VM2", "user defined"})
        Me.cboABSCoeff.Location = New System.Drawing.Point(240, 16)
        Me.cboABSCoeff.Name = "cboABSCoeff"
        Me.cboABSCoeff.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboABSCoeff.Size = New System.Drawing.Size(153, 22)
        Me.cboABSCoeff.TabIndex = 124
        Me.cboABSCoeff.Text = "wood"
        '
        'txtEmissionCoefficient
        '
        Me.txtEmissionCoefficient.AcceptsReturn = True
        Me.txtEmissionCoefficient.BackColor = System.Drawing.SystemColors.Window
        Me.txtEmissionCoefficient.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEmissionCoefficient.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEmissionCoefficient.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtEmissionCoefficient.Location = New System.Drawing.Point(336, 112)
        Me.txtEmissionCoefficient.MaxLength = 0
        Me.txtEmissionCoefficient.Name = "txtEmissionCoefficient"
        Me.txtEmissionCoefficient.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEmissionCoefficient.Size = New System.Drawing.Size(57, 20)
        Me.txtEmissionCoefficient.TabIndex = 54
        Me.txtEmissionCoefficient.Text = "0.8"
        '
        'Label23
        '
        Me.Label23.BackColor = System.Drawing.SystemColors.Control
        Me.Label23.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label23.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label23.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label23.Location = New System.Drawing.Point(48, 208)
        Me.Label23.Name = "Label23"
        Me.Label23.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label23.Size = New System.Drawing.Size(33, 17)
        Me.Label23.TabIndex = 153
        Me.Label23.Text = "C"
        '
        'Label24
        '
        Me.Label24.BackColor = System.Drawing.SystemColors.Control
        Me.Label24.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label24.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label24.Location = New System.Drawing.Point(88, 208)
        Me.Label24.Name = "Label24"
        Me.Label24.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label24.Size = New System.Drawing.Size(33, 17)
        Me.Label24.TabIndex = 152
        Me.Label24.Text = "H"
        '
        'Label25
        '
        Me.Label25.BackColor = System.Drawing.SystemColors.Control
        Me.Label25.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label25.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label25.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label25.Location = New System.Drawing.Point(128, 208)
        Me.Label25.Name = "Label25"
        Me.Label25.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label25.Size = New System.Drawing.Size(33, 17)
        Me.Label25.TabIndex = 151
        Me.Label25.Text = "O"
        '
        'Label26
        '
        Me.Label26.BackColor = System.Drawing.SystemColors.Control
        Me.Label26.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label26.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label26.Location = New System.Drawing.Point(168, 208)
        Me.Label26.Name = "Label26"
        Me.Label26.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label26.Size = New System.Drawing.Size(33, 17)
        Me.Label26.TabIndex = 150
        Me.Label26.Text = "N"
        '
        'Label27
        '
        Me.Label27.BackColor = System.Drawing.SystemColors.Control
        Me.Label27.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label27.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label27.Location = New System.Drawing.Point(216, 240)
        Me.Label27.Name = "Label27"
        Me.Label27.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label27.Size = New System.Drawing.Size(121, 17)
        Me.Label27.TabIndex = 149
        Me.Label27.Text = "Atoms in 1 mole of fuel"
        '
        'Label22
        '
        Me.Label22.BackColor = System.Drawing.SystemColors.Control
        Me.Label22.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label22.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label22.Location = New System.Drawing.Point(143, 144)
        Me.Label22.Name = "Label22"
        Me.Label22.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label22.Size = New System.Drawing.Size(178, 20)
        Me.Label22.TabIndex = 143
        Me.Label22.Text = "Soot Yield Alpha Constant "
        Me.Label22.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label21
        '
        Me.Label21.BackColor = System.Drawing.SystemColors.Control
        Me.Label21.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label21.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label21.Location = New System.Drawing.Point(128, 176)
        Me.Label21.Name = "Label21"
        Me.Label21.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label21.Size = New System.Drawing.Size(193, 20)
        Me.Label21.TabIndex = 142
        Me.Label21.Text = "Soot Yield Epsilon Constant"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label18
        '
        Me.Label18.BackColor = System.Drawing.SystemColors.Control
        Me.Label18.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label18.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label18.Location = New System.Drawing.Point(80, 16)
        Me.Label18.Name = "Label18"
        Me.Label18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label18.Size = New System.Drawing.Size(113, 17)
        Me.Label18.TabIndex = 125
        Me.Label18.Text = "Fuel Type"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblEmissionCoefficient
        '
        Me.lblEmissionCoefficient.BackColor = System.Drawing.SystemColors.Control
        Me.lblEmissionCoefficient.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEmissionCoefficient.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEmissionCoefficient.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblEmissionCoefficient.Location = New System.Drawing.Point(128, 112)
        Me.lblEmissionCoefficient.Name = "lblEmissionCoefficient"
        Me.lblEmissionCoefficient.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEmissionCoefficient.Size = New System.Drawing.Size(193, 20)
        Me.lblEmissionCoefficient.TabIndex = 57
        Me.lblEmissionCoefficient.Text = "Flame emission coefficient (1/m)"
        Me.lblEmissionCoefficient.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_SSTab2_TabPage4
        '
        Me._SSTab2_TabPage4.Controls.Add(Me._Frame5_1)
        Me._SSTab2_TabPage4.Location = New System.Drawing.Point(4, 43)
        Me._SSTab2_TabPage4.Name = "_SSTab2_TabPage4"
        Me._SSTab2_TabPage4.Size = New System.Drawing.Size(497, 334)
        Me._SSTab2_TabPage4.TabIndex = 4
        Me._SSTab2_TabPage4.Text = "Environment"
        Me._SSTab2_TabPage4.UseVisualStyleBackColor = True
        '
        '_Frame5_1
        '
        Me._Frame5_1.BackColor = System.Drawing.SystemColors.Control
        Me._Frame5_1.Controls.Add(Me.cmdDist_RH)
        Me._Frame5_1.Controls.Add(Me.cmdDist_exteriortemp)
        Me._Frame5_1.Controls.Add(Me.cmdDist_interiortemp)
        Me._Frame5_1.Controls.Add(Me.txtRelativeHumidity)
        Me._Frame5_1.Controls.Add(Me.txtExteriorTemp)
        Me._Frame5_1.Controls.Add(Me.txtInteriorTemp)
        Me._Frame5_1.Controls.Add(Me.lblRelativeHumidity)
        Me._Frame5_1.Controls.Add(Me.lblExteriorTemp)
        Me._Frame5_1.Controls.Add(Me.lblInteriorTemp)
        Me._Frame5_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Frame5_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame5.SetIndex(Me._Frame5_1, CType(1, Short))
        Me._Frame5_1.Location = New System.Drawing.Point(27, 28)
        Me._Frame5_1.Name = "_Frame5_1"
        Me._Frame5_1.Padding = New System.Windows.Forms.Padding(0)
        Me._Frame5_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame5_1.Size = New System.Drawing.Size(383, 185)
        Me._Frame5_1.TabIndex = 109
        Me._Frame5_1.TabStop = False
        '
        'cmdDist_RH
        '
        Me.cmdDist_RH.Location = New System.Drawing.Point(302, 120)
        Me.cmdDist_RH.Name = "cmdDist_RH"
        Me.cmdDist_RH.Size = New System.Drawing.Size(68, 19)
        Me.cmdDist_RH.TabIndex = 153
        Me.cmdDist_RH.Text = "distribution"
        Me.cmdDist_RH.UseVisualStyleBackColor = True
        '
        'cmdDist_exteriortemp
        '
        Me.cmdDist_exteriortemp.Location = New System.Drawing.Point(302, 80)
        Me.cmdDist_exteriortemp.Name = "cmdDist_exteriortemp"
        Me.cmdDist_exteriortemp.Size = New System.Drawing.Size(68, 19)
        Me.cmdDist_exteriortemp.TabIndex = 152
        Me.cmdDist_exteriortemp.Text = "distribution"
        Me.cmdDist_exteriortemp.UseVisualStyleBackColor = True
        '
        'cmdDist_interiortemp
        '
        Me.cmdDist_interiortemp.Location = New System.Drawing.Point(302, 40)
        Me.cmdDist_interiortemp.Name = "cmdDist_interiortemp"
        Me.cmdDist_interiortemp.Size = New System.Drawing.Size(68, 19)
        Me.cmdDist_interiortemp.TabIndex = 151
        Me.cmdDist_interiortemp.Text = "distribution"
        Me.cmdDist_interiortemp.UseVisualStyleBackColor = True
        '
        'txtRelativeHumidity
        '
        Me.txtRelativeHumidity.AcceptsReturn = True
        Me.txtRelativeHumidity.BackColor = System.Drawing.SystemColors.Window
        Me.txtRelativeHumidity.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRelativeHumidity.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRelativeHumidity.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRelativeHumidity.Location = New System.Drawing.Point(216, 120)
        Me.txtRelativeHumidity.MaxLength = 0
        Me.txtRelativeHumidity.Name = "txtRelativeHumidity"
        Me.txtRelativeHumidity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRelativeHumidity.Size = New System.Drawing.Size(57, 20)
        Me.txtRelativeHumidity.TabIndex = 112
        Me.txtRelativeHumidity.Text = "65"
        '
        'txtExteriorTemp
        '
        Me.txtExteriorTemp.AcceptsReturn = True
        Me.txtExteriorTemp.BackColor = System.Drawing.SystemColors.Window
        Me.txtExteriorTemp.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtExteriorTemp.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtExteriorTemp.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtExteriorTemp.Location = New System.Drawing.Point(216, 80)
        Me.txtExteriorTemp.MaxLength = 0
        Me.txtExteriorTemp.Name = "txtExteriorTemp"
        Me.txtExteriorTemp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtExteriorTemp.Size = New System.Drawing.Size(57, 20)
        Me.txtExteriorTemp.TabIndex = 111
        Me.txtExteriorTemp.Text = "15"
        '
        'txtInteriorTemp
        '
        Me.txtInteriorTemp.AcceptsReturn = True
        Me.txtInteriorTemp.BackColor = System.Drawing.SystemColors.Window
        Me.txtInteriorTemp.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtInteriorTemp.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtInteriorTemp.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInteriorTemp.Location = New System.Drawing.Point(216, 40)
        Me.txtInteriorTemp.MaxLength = 0
        Me.txtInteriorTemp.Name = "txtInteriorTemp"
        Me.txtInteriorTemp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtInteriorTemp.Size = New System.Drawing.Size(57, 20)
        Me.txtInteriorTemp.TabIndex = 110
        Me.txtInteriorTemp.Text = "20"
        '
        'lblRelativeHumidity
        '
        Me.lblRelativeHumidity.BackColor = System.Drawing.SystemColors.Control
        Me.lblRelativeHumidity.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblRelativeHumidity.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRelativeHumidity.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRelativeHumidity.Location = New System.Drawing.Point(59, 120)
        Me.lblRelativeHumidity.Name = "lblRelativeHumidity"
        Me.lblRelativeHumidity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblRelativeHumidity.Size = New System.Drawing.Size(137, 25)
        Me.lblRelativeHumidity.TabIndex = 115
        Me.lblRelativeHumidity.Text = "Relative Humidity (%)"
        Me.lblRelativeHumidity.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblExteriorTemp
        '
        Me.lblExteriorTemp.BackColor = System.Drawing.SystemColors.Control
        Me.lblExteriorTemp.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblExteriorTemp.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblExteriorTemp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblExteriorTemp.Location = New System.Drawing.Point(59, 80)
        Me.lblExteriorTemp.Name = "lblExteriorTemp"
        Me.lblExteriorTemp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblExteriorTemp.Size = New System.Drawing.Size(137, 25)
        Me.lblExteriorTemp.TabIndex = 114
        Me.lblExteriorTemp.Text = "ExteriorTemperature (C)"
        Me.lblExteriorTemp.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblInteriorTemp
        '
        Me.lblInteriorTemp.BackColor = System.Drawing.SystemColors.Control
        Me.lblInteriorTemp.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblInteriorTemp.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInteriorTemp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInteriorTemp.Location = New System.Drawing.Point(67, 40)
        Me.lblInteriorTemp.Name = "lblInteriorTemp"
        Me.lblInteriorTemp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblInteriorTemp.Size = New System.Drawing.Size(129, 25)
        Me.lblInteriorTemp.TabIndex = 113
        Me.lblInteriorTemp.Text = "Interior Temperature (C)"
        Me.lblInteriorTemp.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_SSTab2_TabPage5
        '
        Me._SSTab2_TabPage5.Controls.Add(Me.GB_gasmodel)
        Me._SSTab2_TabPage5.Controls.Add(Me.Frame15)
        Me._SSTab2_TabPage5.Controls.Add(Me.Frame7)
        Me._SSTab2_TabPage5.Controls.Add(Me.GB_activity)
        Me._SSTab2_TabPage5.Controls.Add(Me.Frame6)
        Me._SSTab2_TabPage5.Location = New System.Drawing.Point(4, 43)
        Me._SSTab2_TabPage5.Name = "_SSTab2_TabPage5"
        Me._SSTab2_TabPage5.Size = New System.Drawing.Size(497, 334)
        Me._SSTab2_TabPage5.TabIndex = 5
        Me._SSTab2_TabPage5.Text = "Tenability"
        Me._SSTab2_TabPage5.UseVisualStyleBackColor = True
        '
        'GB_gasmodel
        '
        Me.GB_gasmodel.Controls.Add(Me.optFEDGeneral)
        Me.GB_gasmodel.Controls.Add(Me.optFEDCO)
        Me.GB_gasmodel.Location = New System.Drawing.Point(15, 7)
        Me.GB_gasmodel.Name = "GB_gasmodel"
        Me.GB_gasmodel.Size = New System.Drawing.Size(469, 50)
        Me.GB_gasmodel.TabIndex = 86
        Me.GB_gasmodel.TabStop = False
        Me.GB_gasmodel.Text = "Asphyxiant Gas Model"
        '
        'optFEDGeneral
        '
        Me.optFEDGeneral.AutoSize = True
        Me.optFEDGeneral.Location = New System.Drawing.Point(180, 19)
        Me.optFEDGeneral.Name = "optFEDGeneral"
        Me.optFEDGeneral.Size = New System.Drawing.Size(115, 18)
        Me.optFEDGeneral.TabIndex = 1
        Me.optFEDGeneral.Text = "FED(CO/CO2/HCN)"
        Me.optFEDGeneral.UseVisualStyleBackColor = True
        '
        'optFEDCO
        '
        Me.optFEDCO.AutoSize = True
        Me.optFEDCO.Checked = True
        Me.optFEDCO.Location = New System.Drawing.Point(320, 19)
        Me.optFEDCO.Name = "optFEDCO"
        Me.optFEDCO.Size = New System.Drawing.Size(109, 18)
        Me.optFEDCO.TabIndex = 0
        Me.optFEDCO.TabStop = True
        Me.optFEDCO.Text = "FED(CO) - C/VM2"
        Me.optFEDCO.UseVisualStyleBackColor = True
        '
        'Frame15
        '
        Me.Frame15.BackColor = System.Drawing.SystemColors.Control
        Me.Frame15.Controls.Add(Me.optReflectiveSign)
        Me.Frame15.Controls.Add(Me.optIlluminatedSign)
        Me.Frame15.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame15.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame15.Location = New System.Drawing.Point(15, 169)
        Me.Frame15.Name = "Frame15"
        Me.Frame15.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame15.Size = New System.Drawing.Size(138, 68)
        Me.Frame15.TabIndex = 85
        Me.Frame15.TabStop = False
        Me.Frame15.Text = "Signage"
        '
        'optReflectiveSign
        '
        Me.optReflectiveSign.BackColor = System.Drawing.SystemColors.Control
        Me.optReflectiveSign.Checked = True
        Me.optReflectiveSign.Cursor = System.Windows.Forms.Cursors.Default
        Me.optReflectiveSign.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optReflectiveSign.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optReflectiveSign.Location = New System.Drawing.Point(8, 40)
        Me.optReflectiveSign.Name = "optReflectiveSign"
        Me.optReflectiveSign.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optReflectiveSign.Size = New System.Drawing.Size(122, 22)
        Me.optReflectiveSign.TabIndex = 87
        Me.optReflectiveSign.TabStop = True
        Me.optReflectiveSign.Text = "Reflective or other"
        Me.optReflectiveSign.UseVisualStyleBackColor = False
        '
        'optIlluminatedSign
        '
        Me.optIlluminatedSign.BackColor = System.Drawing.SystemColors.Control
        Me.optIlluminatedSign.Cursor = System.Windows.Forms.Cursors.Default
        Me.optIlluminatedSign.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optIlluminatedSign.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optIlluminatedSign.Location = New System.Drawing.Point(8, 17)
        Me.optIlluminatedSign.Name = "optIlluminatedSign"
        Me.optIlluminatedSign.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optIlluminatedSign.Size = New System.Drawing.Size(105, 17)
        Me.optIlluminatedSign.TabIndex = 86
        Me.optIlluminatedSign.TabStop = True
        Me.optIlluminatedSign.Text = "Illuminated"
        Me.optIlluminatedSign.UseVisualStyleBackColor = False
        '
        'Frame7
        '
        Me.Frame7.BackColor = System.Drawing.SystemColors.Control
        Me.Frame7.Controls.Add(Me.txtMonitorHeight)
        Me.Frame7.Controls.Add(Me.Label6)
        Me.Frame7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame7.Location = New System.Drawing.Point(15, 243)
        Me.Frame7.Name = "Frame7"
        Me.Frame7.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame7.Size = New System.Drawing.Size(384, 57)
        Me.Frame7.TabIndex = 77
        Me.Frame7.TabStop = False
        '
        'txtMonitorHeight
        '
        Me.txtMonitorHeight.AcceptsReturn = True
        Me.txtMonitorHeight.BackColor = System.Drawing.SystemColors.Window
        Me.txtMonitorHeight.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMonitorHeight.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMonitorHeight.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMonitorHeight.Location = New System.Drawing.Point(224, 16)
        Me.txtMonitorHeight.MaxLength = 0
        Me.txtMonitorHeight.Name = "txtMonitorHeight"
        Me.txtMonitorHeight.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMonitorHeight.Size = New System.Drawing.Size(41, 20)
        Me.txtMonitorHeight.TabIndex = 80
        Me.txtMonitorHeight.Text = "2"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(33, 19)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(185, 26)
        Me.Label6.TabIndex = 84
        Me.Label6.Text = "Monitoring Height Above Floor (m)"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'GB_activity
        '
        Me.GB_activity.BackColor = System.Drawing.SystemColors.Control
        Me.GB_activity.Controls.Add(Me.optHeavyWork)
        Me.GB_activity.Controls.Add(Me.optLightWork)
        Me.GB_activity.Controls.Add(Me.optAtRest)
        Me.GB_activity.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GB_activity.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.SetIndex(Me.GB_activity, CType(1, Short))
        Me.GB_activity.Location = New System.Drawing.Point(15, 63)
        Me.GB_activity.Name = "GB_activity"
        Me.GB_activity.Padding = New System.Windows.Forms.Padding(0)
        Me.GB_activity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GB_activity.Size = New System.Drawing.Size(138, 100)
        Me.GB_activity.TabIndex = 73
        Me.GB_activity.TabStop = False
        Me.GB_activity.Text = "Activity Level"
        '
        'optHeavyWork
        '
        Me.optHeavyWork.BackColor = System.Drawing.SystemColors.Control
        Me.optHeavyWork.Cursor = System.Windows.Forms.Cursors.Default
        Me.optHeavyWork.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optHeavyWork.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optHeavyWork.Location = New System.Drawing.Point(16, 66)
        Me.optHeavyWork.Name = "optHeavyWork"
        Me.optHeavyWork.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optHeavyWork.Size = New System.Drawing.Size(89, 25)
        Me.optHeavyWork.TabIndex = 76
        Me.optHeavyWork.TabStop = True
        Me.optHeavyWork.Text = "Heavy Work"
        Me.optHeavyWork.UseVisualStyleBackColor = False
        '
        'optLightWork
        '
        Me.optLightWork.BackColor = System.Drawing.SystemColors.Control
        Me.optLightWork.Checked = True
        Me.optLightWork.Cursor = System.Windows.Forms.Cursors.Default
        Me.optLightWork.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optLightWork.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optLightWork.Location = New System.Drawing.Point(16, 36)
        Me.optLightWork.Name = "optLightWork"
        Me.optLightWork.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optLightWork.Size = New System.Drawing.Size(114, 33)
        Me.optLightWork.TabIndex = 75
        Me.optLightWork.TabStop = True
        Me.optLightWork.Text = "Light Work"
        Me.optLightWork.UseVisualStyleBackColor = False
        '
        'optAtRest
        '
        Me.optAtRest.BackColor = System.Drawing.SystemColors.Control
        Me.optAtRest.Cursor = System.Windows.Forms.Cursors.Default
        Me.optAtRest.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optAtRest.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optAtRest.Location = New System.Drawing.Point(16, 16)
        Me.optAtRest.Name = "optAtRest"
        Me.optAtRest.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optAtRest.Size = New System.Drawing.Size(65, 25)
        Me.optAtRest.TabIndex = 74
        Me.optAtRest.TabStop = True
        Me.optAtRest.Text = "At Rest"
        Me.optAtRest.UseVisualStyleBackColor = False
        '
        'Frame6
        '
        Me.Frame6.BackColor = System.Drawing.SystemColors.Control
        Me.Frame6.Controls.Add(Me.txtTemp)
        Me.Frame6.Controls.Add(Me.txtConvect)
        Me.Frame6.Controls.Add(Me.txtTarget)
        Me.Frame6.Controls.Add(Me.txtVisibility)
        Me.Frame6.Controls.Add(Me.txtFED)
        Me.Frame6.Controls.Add(Me.lblConvect)
        Me.Frame6.Controls.Add(Me.lblTarget)
        Me.Frame6.Controls.Add(Me.lblVisibility)
        Me.Frame6.Controls.Add(Me.lblFED)
        Me.Frame6.Controls.Add(Me.lblTemp)
        Me.Frame6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame6.Location = New System.Drawing.Point(159, 63)
        Me.Frame6.Name = "Frame6"
        Me.Frame6.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame6.Size = New System.Drawing.Size(240, 174)
        Me.Frame6.TabIndex = 62
        Me.Frame6.TabStop = False
        Me.Frame6.Text = "Tenability and Other Criteria"
        '
        'txtTemp
        '
        Me.txtTemp.AcceptsReturn = True
        Me.txtTemp.BackColor = System.Drawing.SystemColors.Window
        Me.txtTemp.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTemp.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTemp.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTemp.Location = New System.Drawing.Point(176, 99)
        Me.txtTemp.MaxLength = 0
        Me.txtTemp.Name = "txtTemp"
        Me.txtTemp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTemp.Size = New System.Drawing.Size(41, 20)
        Me.txtTemp.TabIndex = 67
        Me.txtTemp.Text = "500"
        '
        'txtConvect
        '
        Me.txtConvect.AcceptsReturn = True
        Me.txtConvect.BackColor = System.Drawing.SystemColors.Window
        Me.txtConvect.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtConvect.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtConvect.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtConvect.Location = New System.Drawing.Point(177, 136)
        Me.txtConvect.MaxLength = 0
        Me.txtConvect.Name = "txtConvect"
        Me.txtConvect.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtConvect.Size = New System.Drawing.Size(41, 20)
        Me.txtConvect.TabIndex = 66
        Me.txtConvect.Text = "80"
        Me.txtConvect.Visible = False
        '
        'txtTarget
        '
        Me.txtTarget.AcceptsReturn = True
        Me.txtTarget.BackColor = System.Drawing.SystemColors.Window
        Me.txtTarget.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTarget.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTarget.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTarget.Location = New System.Drawing.Point(177, 47)
        Me.txtTarget.MaxLength = 0
        Me.txtTarget.Name = "txtTarget"
        Me.txtTarget.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTarget.Size = New System.Drawing.Size(41, 20)
        Me.txtTarget.TabIndex = 65
        Me.txtTarget.Text = "0.3"
        '
        'txtVisibility
        '
        Me.txtVisibility.AcceptsReturn = True
        Me.txtVisibility.BackColor = System.Drawing.SystemColors.Window
        Me.txtVisibility.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtVisibility.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVisibility.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVisibility.Location = New System.Drawing.Point(176, 73)
        Me.txtVisibility.MaxLength = 0
        Me.txtVisibility.Name = "txtVisibility"
        Me.txtVisibility.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtVisibility.Size = New System.Drawing.Size(41, 20)
        Me.txtVisibility.TabIndex = 64
        Me.txtVisibility.Text = "10"
        '
        'txtFED
        '
        Me.txtFED.AcceptsReturn = True
        Me.txtFED.BackColor = System.Drawing.SystemColors.Window
        Me.txtFED.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFED.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFED.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFED.Location = New System.Drawing.Point(177, 21)
        Me.txtFED.MaxLength = 0
        Me.txtFED.Name = "txtFED"
        Me.txtFED.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFED.Size = New System.Drawing.Size(41, 20)
        Me.txtFED.TabIndex = 63
        Me.txtFED.Text = "0.3"
        '
        'lblConvect
        '
        Me.lblConvect.BackColor = System.Drawing.SystemColors.Control
        Me.lblConvect.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblConvect.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblConvect.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblConvect.Location = New System.Drawing.Point(41, 139)
        Me.lblConvect.Name = "lblConvect"
        Me.lblConvect.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblConvect.Size = New System.Drawing.Size(129, 17)
        Me.lblConvect.TabIndex = 72
        Me.lblConvect.Text = "Convective Heat Limit (C)"
        Me.lblConvect.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.lblConvect.Visible = False
        '
        'lblTarget
        '
        Me.lblTarget.BackColor = System.Drawing.SystemColors.Control
        Me.lblTarget.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTarget.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTarget.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTarget.Location = New System.Drawing.Point(33, 50)
        Me.lblTarget.Name = "lblTarget"
        Me.lblTarget.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTarget.Size = New System.Drawing.Size(137, 25)
        Me.lblTarget.TabIndex = 71
        Me.lblTarget.Text = "FED thermal (incap)"
        Me.lblTarget.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblVisibility
        '
        Me.lblVisibility.BackColor = System.Drawing.SystemColors.Control
        Me.lblVisibility.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVisibility.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVisibility.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVisibility.Location = New System.Drawing.Point(65, 76)
        Me.lblVisibility.Name = "lblVisibility"
        Me.lblVisibility.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVisibility.Size = New System.Drawing.Size(105, 17)
        Me.lblVisibility.TabIndex = 70
        Me.lblVisibility.Text = "Visibility (m)"
        Me.lblVisibility.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblFED
        '
        Me.lblFED.BackColor = System.Drawing.SystemColors.Control
        Me.lblFED.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFED.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFED.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFED.Location = New System.Drawing.Point(33, 24)
        Me.lblFED.Name = "lblFED"
        Me.lblFED.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFED.Size = New System.Drawing.Size(137, 17)
        Me.lblFED.TabIndex = 69
        Me.lblFED.Text = "FED gases (incap)"
        Me.lblFED.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTemp
        '
        Me.lblTemp.BackColor = System.Drawing.SystemColors.Control
        Me.lblTemp.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTemp.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTemp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTemp.Location = New System.Drawing.Point(14, 102)
        Me.lblTemp.Name = "lblTemp"
        Me.lblTemp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTemp.Size = New System.Drawing.Size(156, 14)
        Me.lblTemp.TabIndex = 68
        Me.lblTemp.Text = "Upper Layer Temperature (C)"
        Me.lblTemp.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_SSTab2_TabPage6
        '
        Me._SSTab2_TabPage6.Controls.Add(Me.cmdQuintiere)
        Me._SSTab2_TabPage6.Controls.Add(Me._Frame14_0)
        Me._SSTab2_TabPage6.Controls.Add(Me.Frame13)
        Me._SSTab2_TabPage6.Controls.Add(Me.Frame12)
        Me._SSTab2_TabPage6.Location = New System.Drawing.Point(4, 43)
        Me._SSTab2_TabPage6.Name = "_SSTab2_TabPage6"
        Me._SSTab2_TabPage6.Size = New System.Drawing.Size(497, 334)
        Me._SSTab2_TabPage6.TabIndex = 6
        Me._SSTab2_TabPage6.Text = "Fire Growth "
        Me._SSTab2_TabPage6.UseVisualStyleBackColor = True
        '
        'cmdQuintiere
        '
        Me.cmdQuintiere.BackColor = System.Drawing.SystemColors.Control
        Me.cmdQuintiere.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdQuintiere.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdQuintiere.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdQuintiere.Location = New System.Drawing.Point(214, 234)
        Me.cmdQuintiere.Name = "cmdQuintiere"
        Me.cmdQuintiere.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdQuintiere.Size = New System.Drawing.Size(153, 25)
        Me.cmdQuintiere.TabIndex = 3
        Me.cmdQuintiere.Text = "More Options"
        Me.cmdQuintiere.UseVisualStyleBackColor = False
        '
        '_Frame14_0
        '
        Me._Frame14_0.BackColor = System.Drawing.SystemColors.Control
        Me._Frame14_0.Controls.Add(Me.txtCeilingHeatFlux)
        Me._Frame14_0.Controls.Add(Me.txtWallHeatFlux)
        Me._Frame14_0.Controls.Add(Me.lblCeilingHeatFLux)
        Me._Frame14_0.Controls.Add(Me.lblHeatFluxWall)
        Me._Frame14_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Frame14_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame14.SetIndex(Me._Frame14_0, CType(0, Short))
        Me._Frame14_0.Location = New System.Drawing.Point(214, 78)
        Me._Frame14_0.Name = "_Frame14_0"
        Me._Frame14_0.Padding = New System.Windows.Forms.Padding(0)
        Me._Frame14_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame14_0.Size = New System.Drawing.Size(153, 153)
        Me._Frame14_0.TabIndex = 99
        Me._Frame14_0.TabStop = False
        Me._Frame14_0.Text = "Data for Karlsson Model"
        Me._Frame14_0.Visible = False
        '
        'txtCeilingHeatFlux
        '
        Me.txtCeilingHeatFlux.AcceptsReturn = True
        Me.txtCeilingHeatFlux.BackColor = System.Drawing.SystemColors.Window
        Me.txtCeilingHeatFlux.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCeilingHeatFlux.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCeilingHeatFlux.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCeilingHeatFlux.Location = New System.Drawing.Point(96, 96)
        Me.txtCeilingHeatFlux.MaxLength = 0
        Me.txtCeilingHeatFlux.Name = "txtCeilingHeatFlux"
        Me.txtCeilingHeatFlux.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCeilingHeatFlux.Size = New System.Drawing.Size(49, 20)
        Me.txtCeilingHeatFlux.TabIndex = 101
        Me.txtCeilingHeatFlux.Text = "35"
        '
        'txtWallHeatFlux
        '
        Me.txtWallHeatFlux.AcceptsReturn = True
        Me.txtWallHeatFlux.BackColor = System.Drawing.SystemColors.Window
        Me.txtWallHeatFlux.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWallHeatFlux.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWallHeatFlux.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWallHeatFlux.Location = New System.Drawing.Point(96, 40)
        Me.txtWallHeatFlux.MaxLength = 0
        Me.txtWallHeatFlux.Name = "txtWallHeatFlux"
        Me.txtWallHeatFlux.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWallHeatFlux.Size = New System.Drawing.Size(49, 20)
        Me.txtWallHeatFlux.TabIndex = 100
        Me.txtWallHeatFlux.Text = "45"
        '
        'lblCeilingHeatFLux
        '
        Me.lblCeilingHeatFLux.BackColor = System.Drawing.SystemColors.Control
        Me.lblCeilingHeatFLux.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCeilingHeatFLux.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCeilingHeatFLux.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCeilingHeatFLux.Location = New System.Drawing.Point(8, 96)
        Me.lblCeilingHeatFLux.Name = "lblCeilingHeatFLux"
        Me.lblCeilingHeatFLux.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCeilingHeatFLux.Size = New System.Drawing.Size(81, 33)
        Me.lblCeilingHeatFLux.TabIndex = 103
        Me.lblCeilingHeatFLux.Text = "Heat Flux to Ceiling (kW/m2)"
        '
        'lblHeatFluxWall
        '
        Me.lblHeatFluxWall.BackColor = System.Drawing.SystemColors.Control
        Me.lblHeatFluxWall.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblHeatFluxWall.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHeatFluxWall.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblHeatFluxWall.Location = New System.Drawing.Point(8, 40)
        Me.lblHeatFluxWall.Name = "lblHeatFluxWall"
        Me.lblHeatFluxWall.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblHeatFluxWall.Size = New System.Drawing.Size(81, 25)
        Me.lblHeatFluxWall.TabIndex = 102
        Me.lblHeatFluxWall.Text = "Heat Flux to Wall (kW/m2)"
        '
        'Frame13
        '
        Me.Frame13.BackColor = System.Drawing.SystemColors.Control
        Me.Frame13.Controls.Add(Me.txtFlameLengthPower)
        Me.Frame13.Controls.Add(Me.txtFlameAreaConstant)
        Me.Frame13.Controls.Add(Me.txtBurnerWidth)
        Me.Frame13.Controls.Add(Me.lblFlameLengthPower)
        Me.Frame13.Controls.Add(Me.lblFlameAreaConstant)
        Me.Frame13.Controls.Add(Me.lblBurnerWidth)
        Me.Frame13.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame13.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame13.Location = New System.Drawing.Point(22, 78)
        Me.Frame13.Name = "Frame13"
        Me.Frame13.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame13.Size = New System.Drawing.Size(185, 153)
        Me.Frame13.TabIndex = 92
        Me.Frame13.TabStop = False
        Me.Frame13.Text = "General Parameters"
        '
        'txtFlameLengthPower
        '
        Me.txtFlameLengthPower.AcceptsReturn = True
        Me.txtFlameLengthPower.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlameLengthPower.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlameLengthPower.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFlameLengthPower.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlameLengthPower.Location = New System.Drawing.Point(128, 112)
        Me.txtFlameLengthPower.MaxLength = 0
        Me.txtFlameLengthPower.Name = "txtFlameLengthPower"
        Me.txtFlameLengthPower.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlameLengthPower.Size = New System.Drawing.Size(49, 20)
        Me.txtFlameLengthPower.TabIndex = 95
        Me.txtFlameLengthPower.Text = "1"
        '
        'txtFlameAreaConstant
        '
        Me.txtFlameAreaConstant.AcceptsReturn = True
        Me.txtFlameAreaConstant.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlameAreaConstant.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlameAreaConstant.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFlameAreaConstant.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlameAreaConstant.Location = New System.Drawing.Point(128, 72)
        Me.txtFlameAreaConstant.MaxLength = 0
        Me.txtFlameAreaConstant.Name = "txtFlameAreaConstant"
        Me.txtFlameAreaConstant.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlameAreaConstant.Size = New System.Drawing.Size(49, 20)
        Me.txtFlameAreaConstant.TabIndex = 94
        Me.txtFlameAreaConstant.Text = "0.0065"
        '
        'txtBurnerWidth
        '
        Me.txtBurnerWidth.AcceptsReturn = True
        Me.txtBurnerWidth.BackColor = System.Drawing.SystemColors.Window
        Me.txtBurnerWidth.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBurnerWidth.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBurnerWidth.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtBurnerWidth.Location = New System.Drawing.Point(128, 32)
        Me.txtBurnerWidth.MaxLength = 0
        Me.txtBurnerWidth.Name = "txtBurnerWidth"
        Me.txtBurnerWidth.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBurnerWidth.Size = New System.Drawing.Size(49, 20)
        Me.txtBurnerWidth.TabIndex = 93
        Me.txtBurnerWidth.Text = "0.17"
        Me.txtBurnerWidth.Visible = False
        '
        'lblFlameLengthPower
        '
        Me.lblFlameLengthPower.BackColor = System.Drawing.SystemColors.Control
        Me.lblFlameLengthPower.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFlameLengthPower.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFlameLengthPower.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFlameLengthPower.Location = New System.Drawing.Point(3, 115)
        Me.lblFlameLengthPower.Name = "lblFlameLengthPower"
        Me.lblFlameLengthPower.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFlameLengthPower.Size = New System.Drawing.Size(119, 25)
        Me.lblFlameLengthPower.TabIndex = 98
        Me.lblFlameLengthPower.Text = "Flame Length Power"
        Me.lblFlameLengthPower.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblFlameAreaConstant
        '
        Me.lblFlameAreaConstant.BackColor = System.Drawing.SystemColors.Control
        Me.lblFlameAreaConstant.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFlameAreaConstant.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFlameAreaConstant.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFlameAreaConstant.Location = New System.Drawing.Point(13, 75)
        Me.lblFlameAreaConstant.Name = "lblFlameAreaConstant"
        Me.lblFlameAreaConstant.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFlameAreaConstant.Size = New System.Drawing.Size(114, 25)
        Me.lblFlameAreaConstant.TabIndex = 97
        Me.lblFlameAreaConstant.Text = "Flame Area Constant (m2/kW)"
        '
        'lblBurnerWidth
        '
        Me.lblBurnerWidth.BackColor = System.Drawing.SystemColors.Control
        Me.lblBurnerWidth.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblBurnerWidth.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBurnerWidth.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblBurnerWidth.Location = New System.Drawing.Point(16, 40)
        Me.lblBurnerWidth.Name = "lblBurnerWidth"
        Me.lblBurnerWidth.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblBurnerWidth.Size = New System.Drawing.Size(89, 25)
        Me.lblBurnerWidth.TabIndex = 96
        Me.lblBurnerWidth.Text = "Burner Width (m)"
        Me.lblBurnerWidth.Visible = False
        '
        'Frame12
        '
        Me.Frame12.BackColor = System.Drawing.SystemColors.Control
        Me.Frame12.Controls.Add(Me.optQuintiere)
        Me.Frame12.Controls.Add(Me.optKarlsson)
        Me.Frame12.Controls.Add(Me.optRCNone)
        Me.Frame12.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame12.Location = New System.Drawing.Point(22, 22)
        Me.Frame12.Name = "Frame12"
        Me.Frame12.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame12.Size = New System.Drawing.Size(345, 49)
        Me.Frame12.TabIndex = 88
        Me.Frame12.TabStop = False
        Me.Frame12.Text = "Fire Growth Simulation Model for Room Lining Materials"
        '
        'optQuintiere
        '
        Me.optQuintiere.BackColor = System.Drawing.SystemColors.Control
        Me.optQuintiere.Cursor = System.Windows.Forms.Cursors.Default
        Me.optQuintiere.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optQuintiere.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optQuintiere.Location = New System.Drawing.Point(176, 24)
        Me.optQuintiere.Name = "optQuintiere"
        Me.optQuintiere.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optQuintiere.Size = New System.Drawing.Size(161, 17)
        Me.optQuintiere.TabIndex = 91
        Me.optQuintiere.TabStop = True
        Me.optQuintiere.Text = "Use Flame Spread Model"
        Me.optQuintiere.UseVisualStyleBackColor = False
        '
        'optKarlsson
        '
        Me.optKarlsson.BackColor = System.Drawing.SystemColors.Control
        Me.optKarlsson.Cursor = System.Windows.Forms.Cursors.Default
        Me.optKarlsson.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optKarlsson.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optKarlsson.Location = New System.Drawing.Point(72, 24)
        Me.optKarlsson.Name = "optKarlsson"
        Me.optKarlsson.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optKarlsson.Size = New System.Drawing.Size(105, 17)
        Me.optKarlsson.TabIndex = 90
        Me.optKarlsson.TabStop = True
        Me.optKarlsson.Text = "Karlsson model"
        Me.optKarlsson.UseVisualStyleBackColor = False
        Me.optKarlsson.Visible = False
        '
        'optRCNone
        '
        Me.optRCNone.BackColor = System.Drawing.SystemColors.Control
        Me.optRCNone.Checked = True
        Me.optRCNone.Cursor = System.Windows.Forms.Cursors.Default
        Me.optRCNone.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optRCNone.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optRCNone.Location = New System.Drawing.Point(16, 24)
        Me.optRCNone.Name = "optRCNone"
        Me.optRCNone.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optRCNone.Size = New System.Drawing.Size(65, 17)
        Me.optRCNone.TabIndex = 89
        Me.optRCNone.TabStop = True
        Me.optRCNone.Text = "None"
        Me.optRCNone.UseVisualStyleBackColor = False
        '
        '_SSTab2_TabPage7
        '
        Me._SSTab2_TabPage7.Controls.Add(Me.Frame3)
        Me._SSTab2_TabPage7.Controls.Add(Me.Frame19)
        Me._SSTab2_TabPage7.Location = New System.Drawing.Point(4, 43)
        Me._SSTab2_TabPage7.Name = "_SSTab2_TabPage7"
        Me._SSTab2_TabPage7.Size = New System.Drawing.Size(497, 334)
        Me._SSTab2_TabPage7.TabIndex = 7
        Me._SSTab2_TabPage7.Text = "Solvers"
        Me._SSTab2_TabPage7.UseVisualStyleBackColor = True
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.txtErrorControlVentFlow)
        Me.Frame3.Controls.Add(Me.Label1)
        Me.Frame3.Controls.Add(Me.lstTimeStep)
        Me.Frame3.Controls.Add(Me.txtErrorControl)
        Me.Frame3.Controls.Add(Me.Label7)
        Me.Frame3.Controls.Add(Me.Label13)
        Me.Frame3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame3.Location = New System.Drawing.Point(34, 32)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(329, 120)
        Me.Frame3.TabIndex = 104
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Solver Settings"
        '
        'txtErrorControlVentFlow
        '
        Me.txtErrorControlVentFlow.Location = New System.Drawing.Point(208, 83)
        Me.txtErrorControlVentFlow.Name = "txtErrorControlVentFlow"
        Me.txtErrorControlVentFlow.Size = New System.Drawing.Size(65, 20)
        Me.txtErrorControlVentFlow.TabIndex = 110
        Me.txtErrorControlVentFlow.Text = "0.001"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(62, 86)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(131, 14)
        Me.Label1.TabIndex = 109
        Me.Label1.Text = "Error Control (vent flows)"
        '
        'lstTimeStep
        '
        Me.lstTimeStep.BackColor = System.Drawing.SystemColors.Window
        Me.lstTimeStep.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstTimeStep.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstTimeStep.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstTimeStep.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstTimeStep.Items.AddRange(New Object() {"4", "2", "1", "0.5", "0.25", "0.125"})
        Me.lstTimeStep.Location = New System.Drawing.Point(208, 24)
        Me.lstTimeStep.Name = "lstTimeStep"
        Me.lstTimeStep.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstTimeStep.Size = New System.Drawing.Size(65, 22)
        Me.lstTimeStep.TabIndex = 106
        '
        'txtErrorControl
        '
        Me.txtErrorControl.AcceptsReturn = True
        Me.txtErrorControl.BackColor = System.Drawing.SystemColors.Window
        Me.txtErrorControl.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtErrorControl.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtErrorControl.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtErrorControl.Location = New System.Drawing.Point(208, 56)
        Me.txtErrorControl.MaxLength = 0
        Me.txtErrorControl.Name = "txtErrorControl"
        Me.txtErrorControl.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtErrorControl.Size = New System.Drawing.Size(65, 20)
        Me.txtErrorControl.TabIndex = 105
        Me.txtErrorControl.Text = "0.1"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(32, 24)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(161, 17)
        Me.Label7.TabIndex = 108
        Me.Label7.Text = "Suggested Timestep (sec)"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.SystemColors.Control
        Me.Label13.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label13.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label13.Location = New System.Drawing.Point(46, 56)
        Me.Label13.Name = "Label13"
        Me.Label13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label13.Size = New System.Drawing.Size(147, 20)
        Me.Label13.TabIndex = 107
        Me.Label13.Text = "Error Control (ODE solver)"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Frame19
        '
        Me.Frame19.BackColor = System.Drawing.SystemColors.Control
        Me.Frame19.Controls.Add(Me.optLUdecom)
        Me.Frame19.Controls.Add(Me.optGaussJor)
        Me.Frame19.Controls.Add(Me._txtNodes_0)
        Me.Frame19.Controls.Add(Me._txtNodes_1)
        Me.Frame19.Controls.Add(Me._txtNodes_2)
        Me.Frame19.Controls.Add(Me.Label14)
        Me.Frame19.Controls.Add(Me.Label15)
        Me.Frame19.Controls.Add(Me.Label16)
        Me.Frame19.Controls.Add(Me.Label17)
        Me.Frame19.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame19.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame19.Location = New System.Drawing.Point(34, 158)
        Me.Frame19.Name = "Frame19"
        Me.Frame19.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame19.Size = New System.Drawing.Size(329, 137)
        Me.Frame19.TabIndex = 7
        Me.Frame19.TabStop = False
        Me.Frame19.Text = "Wall, Ceiling, Floor Heat Transfer "
        '
        'optLUdecom
        '
        Me.optLUdecom.BackColor = System.Drawing.SystemColors.Control
        Me.optLUdecom.Checked = True
        Me.optLUdecom.Cursor = System.Windows.Forms.Cursors.Default
        Me.optLUdecom.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optLUdecom.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optLUdecom.Location = New System.Drawing.Point(16, 32)
        Me.optLUdecom.Name = "optLUdecom"
        Me.optLUdecom.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optLUdecom.Size = New System.Drawing.Size(137, 17)
        Me.optLUdecom.TabIndex = 12
        Me.optLUdecom.TabStop = True
        Me.optLUdecom.Text = "LU Decomposition"
        Me.optLUdecom.UseVisualStyleBackColor = False
        '
        'optGaussJor
        '
        Me.optGaussJor.BackColor = System.Drawing.SystemColors.Control
        Me.optGaussJor.Cursor = System.Windows.Forms.Cursors.Default
        Me.optGaussJor.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optGaussJor.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optGaussJor.Location = New System.Drawing.Point(160, 32)
        Me.optGaussJor.Name = "optGaussJor"
        Me.optGaussJor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optGaussJor.Size = New System.Drawing.Size(161, 17)
        Me.optGaussJor.TabIndex = 11
        Me.optGaussJor.TabStop = True
        Me.optGaussJor.Text = "Gauss-Jordan Elimination"
        Me.optGaussJor.UseVisualStyleBackColor = False
        '
        '_txtNodes_0
        '
        Me._txtNodes_0.AcceptsReturn = True
        Me._txtNodes_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtNodes_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtNodes_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtNodes_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNodes.SetIndex(Me._txtNodes_0, CType(0, Short))
        Me._txtNodes_0.Location = New System.Drawing.Point(24, 80)
        Me._txtNodes_0.MaxLength = 0
        Me._txtNodes_0.Name = "_txtNodes_0"
        Me._txtNodes_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtNodes_0.Size = New System.Drawing.Size(57, 20)
        Me._txtNodes_0.TabIndex = 10
        Me._txtNodes_0.Text = "20"
        '
        '_txtNodes_1
        '
        Me._txtNodes_1.AcceptsReturn = True
        Me._txtNodes_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtNodes_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtNodes_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtNodes_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNodes.SetIndex(Me._txtNodes_1, CType(1, Short))
        Me._txtNodes_1.Location = New System.Drawing.Point(120, 80)
        Me._txtNodes_1.MaxLength = 0
        Me._txtNodes_1.Name = "_txtNodes_1"
        Me._txtNodes_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtNodes_1.Size = New System.Drawing.Size(57, 20)
        Me._txtNodes_1.TabIndex = 9
        Me._txtNodes_1.Text = "20"
        '
        '_txtNodes_2
        '
        Me._txtNodes_2.AcceptsReturn = True
        Me._txtNodes_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtNodes_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtNodes_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtNodes_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNodes.SetIndex(Me._txtNodes_2, CType(2, Short))
        Me._txtNodes_2.Location = New System.Drawing.Point(216, 80)
        Me._txtNodes_2.MaxLength = 0
        Me._txtNodes_2.Name = "_txtNodes_2"
        Me._txtNodes_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtNodes_2.Size = New System.Drawing.Size(57, 20)
        Me._txtNodes_2.TabIndex = 8
        Me._txtNodes_2.Text = "10"
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.SystemColors.Control
        Me.Label14.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label14.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label14.Location = New System.Drawing.Point(24, 56)
        Me.Label14.Name = "Label14"
        Me.Label14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label14.Size = New System.Drawing.Size(257, 17)
        Me.Label14.TabIndex = 16
        Me.Label14.Text = "Number of Nodes per Layer (not < 10)"
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.SystemColors.Control
        Me.Label15.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label15.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label15.Location = New System.Drawing.Point(24, 104)
        Me.Label15.Name = "Label15"
        Me.Label15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label15.Size = New System.Drawing.Size(73, 17)
        Me.Label15.TabIndex = 15
        Me.Label15.Text = "Ceiling"
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.SystemColors.Control
        Me.Label16.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label16.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label16.Location = New System.Drawing.Point(120, 104)
        Me.Label16.Name = "Label16"
        Me.Label16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label16.Size = New System.Drawing.Size(73, 17)
        Me.Label16.TabIndex = 14
        Me.Label16.Text = "Walls"
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.SystemColors.Control
        Me.Label17.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label17.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label17.Location = New System.Drawing.Point(216, 104)
        Me.Label17.Name = "Label17"
        Me.Label17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label17.Size = New System.Drawing.Size(73, 17)
        Me.Label17.TabIndex = 13
        Me.Label17.Text = "Floor"
        '
        '_SSTab2_TabPage8
        '
        Me._SSTab2_TabPage8.Controls.Add(Me.Frame16)
        Me._SSTab2_TabPage8.Location = New System.Drawing.Point(4, 43)
        Me._SSTab2_TabPage8.Name = "_SSTab2_TabPage8"
        Me._SSTab2_TabPage8.Size = New System.Drawing.Size(497, 334)
        Me._SSTab2_TabPage8.TabIndex = 8
        Me._SSTab2_TabPage8.Text = "Post Flashover"
        Me._SSTab2_TabPage8.UseVisualStyleBackColor = True
        '
        'Frame16
        '
        Me.Frame16.BackColor = System.Drawing.SystemColors.Control
        Me.Frame16.Controls.Add(Me.GB_flashovercriteria)
        Me.Frame16.Controls.Add(Me.ChkFRR)
        Me.Frame16.Controls.Add(Me.Frame17)
        Me.Frame16.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame16.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame16.Location = New System.Drawing.Point(18, 3)
        Me.Frame16.Name = "Frame16"
        Me.Frame16.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame16.Size = New System.Drawing.Size(475, 328)
        Me.Frame16.TabIndex = 17
        Me.Frame16.TabStop = False
        '
        'GB_flashovercriteria
        '
        Me.GB_flashovercriteria.Controls.Add(Me.optFOtemp)
        Me.GB_flashovercriteria.Controls.Add(Me.optFOflux)
        Me.GB_flashovercriteria.Location = New System.Drawing.Point(8, 16)
        Me.GB_flashovercriteria.Name = "GB_flashovercriteria"
        Me.GB_flashovercriteria.Size = New System.Drawing.Size(453, 49)
        Me.GB_flashovercriteria.TabIndex = 157
        Me.GB_flashovercriteria.TabStop = False
        Me.GB_flashovercriteria.Text = "Flashover Criteria"
        '
        'optFOtemp
        '
        Me.optFOtemp.AutoSize = True
        Me.optFOtemp.Checked = True
        Me.optFOtemp.Location = New System.Drawing.Point(223, 19)
        Me.optFOtemp.Name = "optFOtemp"
        Me.optFOtemp.Size = New System.Drawing.Size(153, 18)
        Me.optFOtemp.TabIndex = 1
        Me.optFOtemp.TabStop = True
        Me.optFOtemp.Text = "Upper Layer Temp > 500 C"
        Me.optFOtemp.UseVisualStyleBackColor = True
        '
        'optFOflux
        '
        Me.optFOflux.AutoSize = True
        Me.optFOflux.Location = New System.Drawing.Point(16, 19)
        Me.optFOflux.Name = "optFOflux"
        Me.optFOflux.Size = New System.Drawing.Size(183, 18)
        Me.optFOflux.TabIndex = 0
        Me.optFOflux.Text = "Radiant Flux on floor > 20 kW/m2"
        Me.optFOflux.UseVisualStyleBackColor = True
        '
        'ChkFRR
        '
        Me.ChkFRR.BackColor = System.Drawing.SystemColors.Control
        Me.ChkFRR.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ChkFRR.Cursor = System.Windows.Forms.Cursors.Default
        Me.ChkFRR.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkFRR.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ChkFRR.Location = New System.Drawing.Point(24, 296)
        Me.ChkFRR.Name = "ChkFRR"
        Me.ChkFRR.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ChkFRR.Size = New System.Drawing.Size(244, 18)
        Me.ChkFRR.TabIndex = 136
        Me.ChkFRR.Text = "Calculate equivalent fire resistance ratings?"
        Me.ChkFRR.UseVisualStyleBackColor = False
        '
        'Frame17
        '
        Me.Frame17.BackColor = System.Drawing.SystemColors.Control
        Me.Frame17.Controls.Add(Me.cmdWoodOption)
        Me.Frame17.Controls.Add(Me.OptPreFlashover)
        Me.Frame17.Controls.Add(Me.optPostFlashover)
        Me.Frame17.Controls.Add(Me.Frame18)
        Me.Frame17.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame17.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame17.Location = New System.Drawing.Point(8, 71)
        Me.Frame17.Name = "Frame17"
        Me.Frame17.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame17.Size = New System.Drawing.Size(453, 219)
        Me.Frame17.TabIndex = 25
        Me.Frame17.TabStop = False
        Me.Frame17.Text = "Use Wood Crib Post-Flashover Model?"
        '
        'cmdWoodOption
        '
        Me.cmdWoodOption.Location = New System.Drawing.Point(320, 20)
        Me.cmdWoodOption.Name = "cmdWoodOption"
        Me.cmdWoodOption.Size = New System.Drawing.Size(121, 30)
        Me.cmdWoodOption.TabIndex = 28
        Me.cmdWoodOption.Text = "CLT model options"
        Me.cmdWoodOption.UseVisualStyleBackColor = True
        '
        'OptPreFlashover
        '
        Me.OptPreFlashover.BackColor = System.Drawing.SystemColors.Control
        Me.OptPreFlashover.Checked = True
        Me.OptPreFlashover.Cursor = System.Windows.Forms.Cursors.Default
        Me.OptPreFlashover.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OptPreFlashover.ForeColor = System.Drawing.SystemColors.ControlText
        Me.OptPreFlashover.Location = New System.Drawing.Point(16, 24)
        Me.OptPreFlashover.Name = "OptPreFlashover"
        Me.OptPreFlashover.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.OptPreFlashover.Size = New System.Drawing.Size(49, 17)
        Me.OptPreFlashover.TabIndex = 27
        Me.OptPreFlashover.TabStop = True
        Me.OptPreFlashover.Text = "No"
        Me.OptPreFlashover.UseVisualStyleBackColor = False
        '
        'optPostFlashover
        '
        Me.optPostFlashover.BackColor = System.Drawing.SystemColors.Control
        Me.optPostFlashover.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPostFlashover.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optPostFlashover.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPostFlashover.Location = New System.Drawing.Point(80, 24)
        Me.optPostFlashover.Name = "optPostFlashover"
        Me.optPostFlashover.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPostFlashover.Size = New System.Drawing.Size(57, 17)
        Me.optPostFlashover.TabIndex = 26
        Me.optPostFlashover.TabStop = True
        Me.optPostFlashover.Text = "Yes"
        Me.optPostFlashover.UseVisualStyleBackColor = False
        '
        'Frame18
        '
        Me.Frame18.BackColor = System.Drawing.SystemColors.Control
        Me.Frame18.Controls.Add(Me.Label3)
        Me.Frame18.Controls.Add(Me.txtCribheight)
        Me.Frame18.Controls.Add(Me.Label2)
        Me.Frame18.Controls.Add(Me.txtExcessFuelFactor)
        Me.Frame18.Controls.Add(Me.cmdDist_HOC)
        Me.Frame18.Controls.Add(Me.txtStickSpacing)
        Me.Frame18.Controls.Add(Me.chkModGER)
        Me.Frame18.Controls.Add(Me.txtFuelThickness)
        Me.Frame18.Controls.Add(Me.txtHOCFuel)
        Me.Frame18.Controls.Add(Me.Label20)
        Me.Frame18.Controls.Add(Me.Label12)
        Me.Frame18.Controls.Add(Me.Label5)
        Me.Frame18.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Frame18.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame18.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame18.Location = New System.Drawing.Point(16, 56)
        Me.Frame18.Name = "Frame18"
        Me.Frame18.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame18.Size = New System.Drawing.Size(425, 151)
        Me.Frame18.TabIndex = 18
        Me.Frame18.TabStop = False
        Me.Frame18.Text = "Fuel (assumes wood crib behaviour)"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(112, 120)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(77, 14)
        Me.Label3.TabIndex = 159
        Me.Label3.Text = "Crib height (m)"
        '
        'txtCribheight
        '
        Me.txtCribheight.Location = New System.Drawing.Point(195, 117)
        Me.txtCribheight.Name = "txtCribheight"
        Me.txtCribheight.Size = New System.Drawing.Size(49, 20)
        Me.txtCribheight.TabIndex = 158
        Me.txtCribheight.Text = "0.8"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(257, 120)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 14)
        Me.Label2.TabIndex = 157
        Me.Label2.Text = "Excess Fuel Factor"
        '
        'txtExcessFuelFactor
        '
        Me.txtExcessFuelFactor.Location = New System.Drawing.Point(363, 117)
        Me.txtExcessFuelFactor.Name = "txtExcessFuelFactor"
        Me.txtExcessFuelFactor.Size = New System.Drawing.Size(49, 20)
        Me.txtExcessFuelFactor.TabIndex = 156
        Me.txtExcessFuelFactor.Text = "1.0"
        '
        'cmdDist_HOC
        '
        Me.cmdDist_HOC.Location = New System.Drawing.Point(263, 26)
        Me.cmdDist_HOC.Name = "cmdDist_HOC"
        Me.cmdDist_HOC.Size = New System.Drawing.Size(68, 19)
        Me.cmdDist_HOC.TabIndex = 155
        Me.cmdDist_HOC.Text = "distribution"
        Me.cmdDist_HOC.UseVisualStyleBackColor = True
        '
        'txtStickSpacing
        '
        Me.txtStickSpacing.AcceptsReturn = True
        Me.txtStickSpacing.BackColor = System.Drawing.SystemColors.Window
        Me.txtStickSpacing.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtStickSpacing.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStickSpacing.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtStickSpacing.Location = New System.Drawing.Point(195, 87)
        Me.txtStickSpacing.MaxLength = 0
        Me.txtStickSpacing.Name = "txtStickSpacing"
        Me.txtStickSpacing.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtStickSpacing.Size = New System.Drawing.Size(49, 20)
        Me.txtStickSpacing.TabIndex = 138
        Me.txtStickSpacing.Text = "0.05"
        '
        'chkModGER
        '
        Me.chkModGER.BackColor = System.Drawing.SystemColors.Control
        Me.chkModGER.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkModGER.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkModGER.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkModGER.Location = New System.Drawing.Point(263, 53)
        Me.chkModGER.Name = "chkModGER"
        Me.chkModGER.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkModGER.Size = New System.Drawing.Size(139, 33)
        Me.chkModGER.TabIndex = 137
        Me.chkModGER.Text = "HoC varies with GER"
        Me.chkModGER.UseVisualStyleBackColor = False
        '
        'txtFuelThickness
        '
        Me.txtFuelThickness.AcceptsReturn = True
        Me.txtFuelThickness.BackColor = System.Drawing.SystemColors.Window
        Me.txtFuelThickness.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFuelThickness.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFuelThickness.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFuelThickness.Location = New System.Drawing.Point(195, 58)
        Me.txtFuelThickness.MaxLength = 0
        Me.txtFuelThickness.Name = "txtFuelThickness"
        Me.txtFuelThickness.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFuelThickness.Size = New System.Drawing.Size(49, 20)
        Me.txtFuelThickness.TabIndex = 134
        Me.txtFuelThickness.Text = "0.05"
        '
        'txtHOCFuel
        '
        Me.txtHOCFuel.AcceptsReturn = True
        Me.txtHOCFuel.BackColor = System.Drawing.SystemColors.Window
        Me.txtHOCFuel.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtHOCFuel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHOCFuel.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtHOCFuel.Location = New System.Drawing.Point(195, 26)
        Me.txtHOCFuel.MaxLength = 0
        Me.txtHOCFuel.Name = "txtHOCFuel"
        Me.txtHOCFuel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtHOCFuel.Size = New System.Drawing.Size(49, 20)
        Me.txtHOCFuel.TabIndex = 19
        Me.txtHOCFuel.Text = "12.4"
        '
        'Label20
        '
        Me.Label20.BackColor = System.Drawing.SystemColors.Control
        Me.Label20.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label20.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label20.Location = New System.Drawing.Point(28, 90)
        Me.Label20.Name = "Label20"
        Me.Label20.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label20.Size = New System.Drawing.Size(161, 17)
        Me.Label20.TabIndex = 139
        Me.Label20.Text = "Stick Spacing (m)"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.SystemColors.Control
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label12.Location = New System.Drawing.Point(10, 58)
        Me.Label12.Name = "Label12"
        Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label12.Size = New System.Drawing.Size(179, 20)
        Me.Label12.TabIndex = 135
        Me.Label12.Text = "Characteristic Stick Thickness  (m)"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(44, 28)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(145, 25)
        Me.Label5.TabIndex = 22
        Me.Label5.Text = "Heat of Combustion (MJ/kg)"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_SSTab2_TabPage10
        '
        Me._SSTab2_TabPage10.Controls.Add(Me.Frame26)
        Me._SSTab2_TabPage10.Controls.Add(Me.Frame27)
        Me._SSTab2_TabPage10.Location = New System.Drawing.Point(4, 43)
        Me._SSTab2_TabPage10.Name = "_SSTab2_TabPage10"
        Me._SSTab2_TabPage10.Size = New System.Drawing.Size(497, 334)
        Me._SSTab2_TabPage10.TabIndex = 10
        Me._SSTab2_TabPage10.Text = "CO / Soot"
        Me._SSTab2_TabPage10.UseVisualStyleBackColor = True
        '
        'Frame26
        '
        Me.Frame26.BackColor = System.Drawing.SystemColors.Control
        Me.Frame26.Controls.Add(Me.Button2)
        Me.Frame26.Controls.Add(Me.txtpreCO)
        Me.Frame26.Controls.Add(Me.txtpostCO)
        Me.Frame26.Controls.Add(Me.optCOauto)
        Me.Frame26.Controls.Add(Me.optCOman)
        Me.Frame26.Controls.Add(Me.Label19)
        Me.Frame26.Controls.Add(Me.Label28)
        Me.Frame26.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame26.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame26.Location = New System.Drawing.Point(16, 24)
        Me.Frame26.Name = "Frame26"
        Me.Frame26.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame26.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame26.Size = New System.Drawing.Size(469, 156)
        Me.Frame26.TabIndex = 162
        Me.Frame26.TabStop = False
        Me.Frame26.Text = "CO Production"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(286, 61)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(68, 19)
        Me.Button2.TabIndex = 177
        Me.Button2.Text = "distribution"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'txtpreCO
        '
        Me.txtpreCO.AcceptsReturn = True
        Me.txtpreCO.BackColor = System.Drawing.SystemColors.Window
        Me.txtpreCO.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtpreCO.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtpreCO.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtpreCO.Location = New System.Drawing.Point(213, 61)
        Me.txtpreCO.MaxLength = 0
        Me.txtpreCO.Name = "txtpreCO"
        Me.txtpreCO.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtpreCO.Size = New System.Drawing.Size(57, 20)
        Me.txtpreCO.TabIndex = 170
        Me.txtpreCO.Text = "0.04"
        '
        'txtpostCO
        '
        Me.txtpostCO.AcceptsReturn = True
        Me.txtpostCO.BackColor = System.Drawing.SystemColors.Window
        Me.txtpostCO.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtpostCO.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtpostCO.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtpostCO.Location = New System.Drawing.Point(213, 93)
        Me.txtpostCO.MaxLength = 0
        Me.txtpostCO.Name = "txtpostCO"
        Me.txtpostCO.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtpostCO.Size = New System.Drawing.Size(57, 20)
        Me.txtpostCO.TabIndex = 168
        Me.txtpostCO.Text = "0.4"
        '
        'optCOauto
        '
        Me.optCOauto.BackColor = System.Drawing.SystemColors.Control
        Me.optCOauto.Checked = True
        Me.optCOauto.Cursor = System.Windows.Forms.Cursors.Default
        Me.optCOauto.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optCOauto.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optCOauto.Location = New System.Drawing.Point(21, 128)
        Me.optCOauto.Name = "optCOauto"
        Me.optCOauto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optCOauto.Size = New System.Drawing.Size(177, 25)
        Me.optCOauto.TabIndex = 165
        Me.optCOauto.TabStop = True
        Me.optCOauto.Text = "auto (based on GER)"
        Me.optCOauto.UseVisualStyleBackColor = False
        '
        'optCOman
        '
        Me.optCOman.BackColor = System.Drawing.SystemColors.Control
        Me.optCOman.Cursor = System.Windows.Forms.Cursors.Default
        Me.optCOman.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optCOman.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optCOman.Location = New System.Drawing.Point(16, 24)
        Me.optCOman.Name = "optCOman"
        Me.optCOman.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optCOman.Size = New System.Drawing.Size(273, 25)
        Me.optCOman.TabIndex = 164
        Me.optCOman.TabStop = True
        Me.optCOman.Text = "Manual input of pre/post flashover CO yields"
        Me.optCOman.UseVisualStyleBackColor = False
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.SystemColors.Control
        Me.Label19.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label19.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label19.Location = New System.Drawing.Point(45, 61)
        Me.Label19.Name = "Label19"
        Me.Label19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label19.Size = New System.Drawing.Size(153, 17)
        Me.Label19.TabIndex = 171
        Me.Label19.Text = "CO pre-flashover yield (g/g)"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label28
        '
        Me.Label28.BackColor = System.Drawing.SystemColors.Control
        Me.Label28.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label28.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label28.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label28.Location = New System.Drawing.Point(45, 93)
        Me.Label28.Name = "Label28"
        Me.Label28.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label28.Size = New System.Drawing.Size(153, 17)
        Me.Label28.TabIndex = 169
        Me.Label28.Text = "CO postflashover yield (g/g)"
        Me.Label28.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Frame27
        '
        Me.Frame27.BackColor = System.Drawing.SystemColors.Control
        Me.Frame27.Controls.Add(Me.Button1)
        Me.Frame27.Controls.Add(Me.txtpreSoot)
        Me.Frame27.Controls.Add(Me.txtPostSoot)
        Me.Frame27.Controls.Add(Me.optSootman)
        Me.Frame27.Controls.Add(Me.optSootauto)
        Me.Frame27.Controls.Add(Me.Label30)
        Me.Frame27.Controls.Add(Me.Label29)
        Me.Frame27.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame27.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame27.Location = New System.Drawing.Point(16, 186)
        Me.Frame27.Name = "Frame27"
        Me.Frame27.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame27.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame27.Size = New System.Drawing.Size(469, 145)
        Me.Frame27.TabIndex = 163
        Me.Frame27.TabStop = False
        Me.Frame27.Text = "Soot Production"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(321, 72)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(68, 19)
        Me.Button1.TabIndex = 176
        Me.Button1.Text = "distribution"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'txtpreSoot
        '
        Me.txtpreSoot.AcceptsReturn = True
        Me.txtpreSoot.BackColor = System.Drawing.SystemColors.Window
        Me.txtpreSoot.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtpreSoot.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtpreSoot.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtpreSoot.Location = New System.Drawing.Point(248, 72)
        Me.txtpreSoot.MaxLength = 0
        Me.txtpreSoot.Name = "txtpreSoot"
        Me.txtpreSoot.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtpreSoot.Size = New System.Drawing.Size(57, 20)
        Me.txtpreSoot.TabIndex = 174
        Me.txtpreSoot.Text = "0.07"
        '
        'txtPostSoot
        '
        Me.txtPostSoot.AcceptsReturn = True
        Me.txtPostSoot.BackColor = System.Drawing.SystemColors.Window
        Me.txtPostSoot.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPostSoot.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPostSoot.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPostSoot.Location = New System.Drawing.Point(248, 104)
        Me.txtPostSoot.MaxLength = 0
        Me.txtPostSoot.Name = "txtPostSoot"
        Me.txtPostSoot.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPostSoot.Size = New System.Drawing.Size(57, 20)
        Me.txtPostSoot.TabIndex = 172
        Me.txtPostSoot.Text = "0.14"
        '
        'optSootman
        '
        Me.optSootman.BackColor = System.Drawing.SystemColors.Control
        Me.optSootman.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSootman.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optSootman.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optSootman.Location = New System.Drawing.Point(16, 24)
        Me.optSootman.Name = "optSootman"
        Me.optSootman.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSootman.Size = New System.Drawing.Size(249, 25)
        Me.optSootman.TabIndex = 167
        Me.optSootman.TabStop = True
        Me.optSootman.Text = "Manual input of pre/post flashover Soot yields"
        Me.optSootman.UseVisualStyleBackColor = False
        '
        'optSootauto
        '
        Me.optSootauto.BackColor = System.Drawing.SystemColors.Control
        Me.optSootauto.Checked = True
        Me.optSootauto.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSootauto.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optSootauto.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optSootauto.Location = New System.Drawing.Point(271, 24)
        Me.optSootauto.Name = "optSootauto"
        Me.optSootauto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSootauto.Size = New System.Drawing.Size(195, 25)
        Me.optSootauto.TabIndex = 166
        Me.optSootauto.TabStop = True
        Me.optSootauto.Text = "auto (based on GER and object)"
        Me.optSootauto.UseVisualStyleBackColor = False
        '
        'Label30
        '
        Me.Label30.BackColor = System.Drawing.SystemColors.Control
        Me.Label30.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label30.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label30.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label30.Location = New System.Drawing.Point(56, 72)
        Me.Label30.Name = "Label30"
        Me.Label30.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label30.Size = New System.Drawing.Size(177, 17)
        Me.Label30.TabIndex = 175
        Me.Label30.Text = "Soot preflashover yield (g/g)"
        Me.Label30.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label29
        '
        Me.Label29.BackColor = System.Drawing.SystemColors.Control
        Me.Label29.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label29.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label29.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label29.Location = New System.Drawing.Point(56, 104)
        Me.Label29.Name = "Label29"
        Me.Label29.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label29.Size = New System.Drawing.Size(177, 17)
        Me.Label29.TabIndex = 173
        Me.Label29.Text = "Soot postflashover yield (g/g)"
        Me.Label29.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_SSTab2_TabPage9
        '
        Me._SSTab2_TabPage9.Controls.Add(Me.Frame25)
        Me._SSTab2_TabPage9.Controls.Add(Me.Frame20)
        Me._SSTab2_TabPage9.Location = New System.Drawing.Point(4, 43)
        Me._SSTab2_TabPage9.Name = "_SSTab2_TabPage9"
        Me._SSTab2_TabPage9.Size = New System.Drawing.Size(497, 334)
        Me._SSTab2_TabPage9.TabIndex = 9
        Me._SSTab2_TabPage9.Text = "Other"
        Me._SSTab2_TabPage9.UseVisualStyleBackColor = True
        '
        'Frame25
        '
        Me.Frame25.BackColor = System.Drawing.SystemColors.Control
        Me.Frame25.Controls.Add(Me.optFuelLimited)
        Me.Frame25.Controls.Add(Me.optFuelNoLimit)
        Me.Frame25.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame25.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame25.Location = New System.Drawing.Point(21, 160)
        Me.Frame25.Name = "Frame25"
        Me.Frame25.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame25.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame25.Size = New System.Drawing.Size(281, 73)
        Me.Frame25.TabIndex = 159
        Me.Frame25.TabStop = False
        Me.Frame25.Text = "Fuel Load"
        Me.Frame25.Visible = False
        '
        'optFuelLimited
        '
        Me.optFuelLimited.BackColor = System.Drawing.SystemColors.Control
        Me.optFuelLimited.Cursor = System.Windows.Forms.Cursors.Default
        Me.optFuelLimited.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optFuelLimited.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optFuelLimited.Location = New System.Drawing.Point(96, 32)
        Me.optFuelLimited.Name = "optFuelLimited"
        Me.optFuelLimited.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optFuelLimited.Size = New System.Drawing.Size(177, 25)
        Me.optFuelLimited.TabIndex = 161
        Me.optFuelLimited.TabStop = True
        Me.optFuelLimited.Text = "Limited - use postflashover FLED"
        Me.optFuelLimited.UseVisualStyleBackColor = False
        '
        'optFuelNoLimit
        '
        Me.optFuelNoLimit.BackColor = System.Drawing.SystemColors.Control
        Me.optFuelNoLimit.Checked = True
        Me.optFuelNoLimit.Cursor = System.Windows.Forms.Cursors.Default
        Me.optFuelNoLimit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optFuelNoLimit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optFuelNoLimit.Location = New System.Drawing.Point(16, 32)
        Me.optFuelNoLimit.Name = "optFuelNoLimit"
        Me.optFuelNoLimit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optFuelNoLimit.Size = New System.Drawing.Size(97, 25)
        Me.optFuelNoLimit.TabIndex = 160
        Me.optFuelNoLimit.TabStop = True
        Me.optFuelNoLimit.Text = "Unlimited"
        Me.optFuelNoLimit.UseVisualStyleBackColor = False
        '
        'Frame20
        '
        Me.Frame20.BackColor = System.Drawing.SystemColors.Control
        Me.Frame20.Controls.Add(Me.Button3)
        Me.Frame20.Controls.Add(Me.txtFuelSurfaceArea)
        Me.Frame20.Controls.Add(Me.txtLHoG)
        Me.Frame20.Controls.Add(Me.optEnhanceOff)
        Me.Frame20.Controls.Add(Me.optEnhanceOn)
        Me.Frame20.Controls.Add(Me.lblFuelSurfaceArea)
        Me.Frame20.Controls.Add(Me.lblLHoG)
        Me.Frame20.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame20.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame20.Location = New System.Drawing.Point(21, 24)
        Me.Frame20.Name = "Frame20"
        Me.Frame20.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame20.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame20.Size = New System.Drawing.Size(368, 121)
        Me.Frame20.TabIndex = 127
        Me.Frame20.TabStop = False
        Me.Frame20.Text = "Enhance Burning Rate due to Hot Layer Effects"
        Me.Frame20.Visible = False
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(272, 57)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(68, 19)
        Me.Button3.TabIndex = 155
        Me.Button3.Text = "distribution"
        Me.Button3.UseVisualStyleBackColor = True
        Me.Button3.Visible = False
        '
        'txtFuelSurfaceArea
        '
        Me.txtFuelSurfaceArea.AcceptsReturn = True
        Me.txtFuelSurfaceArea.BackColor = System.Drawing.SystemColors.Window
        Me.txtFuelSurfaceArea.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFuelSurfaceArea.Enabled = False
        Me.txtFuelSurfaceArea.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFuelSurfaceArea.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFuelSurfaceArea.Location = New System.Drawing.Point(192, 88)
        Me.txtFuelSurfaceArea.MaxLength = 0
        Me.txtFuelSurfaceArea.Name = "txtFuelSurfaceArea"
        Me.txtFuelSurfaceArea.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFuelSurfaceArea.Size = New System.Drawing.Size(57, 20)
        Me.txtFuelSurfaceArea.TabIndex = 131
        Me.txtFuelSurfaceArea.Text = "0"
        Me.txtFuelSurfaceArea.Visible = False
        '
        'txtLHoG
        '
        Me.txtLHoG.AcceptsReturn = True
        Me.txtLHoG.BackColor = System.Drawing.SystemColors.Window
        Me.txtLHoG.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLHoG.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLHoG.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLHoG.Location = New System.Drawing.Point(192, 56)
        Me.txtLHoG.MaxLength = 0
        Me.txtLHoG.Name = "txtLHoG"
        Me.txtLHoG.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLHoG.Size = New System.Drawing.Size(57, 20)
        Me.txtLHoG.TabIndex = 130
        Me.txtLHoG.Text = "3"
        Me.txtLHoG.Visible = False
        '
        'optEnhanceOff
        '
        Me.optEnhanceOff.BackColor = System.Drawing.SystemColors.Control
        Me.optEnhanceOff.Checked = True
        Me.optEnhanceOff.Cursor = System.Windows.Forms.Cursors.Default
        Me.optEnhanceOff.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optEnhanceOff.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optEnhanceOff.Location = New System.Drawing.Point(40, 24)
        Me.optEnhanceOff.Name = "optEnhanceOff"
        Me.optEnhanceOff.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optEnhanceOff.Size = New System.Drawing.Size(57, 17)
        Me.optEnhanceOff.TabIndex = 129
        Me.optEnhanceOff.TabStop = True
        Me.optEnhanceOff.Text = "Off"
        Me.optEnhanceOff.UseVisualStyleBackColor = False
        '
        'optEnhanceOn
        '
        Me.optEnhanceOn.BackColor = System.Drawing.SystemColors.Control
        Me.optEnhanceOn.Cursor = System.Windows.Forms.Cursors.Default
        Me.optEnhanceOn.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optEnhanceOn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optEnhanceOn.Location = New System.Drawing.Point(160, 24)
        Me.optEnhanceOn.Name = "optEnhanceOn"
        Me.optEnhanceOn.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optEnhanceOn.Size = New System.Drawing.Size(73, 17)
        Me.optEnhanceOn.TabIndex = 128
        Me.optEnhanceOn.TabStop = True
        Me.optEnhanceOn.Text = "On"
        Me.optEnhanceOn.UseVisualStyleBackColor = False
        '
        'lblFuelSurfaceArea
        '
        Me.lblFuelSurfaceArea.BackColor = System.Drawing.SystemColors.Control
        Me.lblFuelSurfaceArea.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFuelSurfaceArea.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFuelSurfaceArea.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFuelSurfaceArea.Location = New System.Drawing.Point(48, 88)
        Me.lblFuelSurfaceArea.Name = "lblFuelSurfaceArea"
        Me.lblFuelSurfaceArea.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFuelSurfaceArea.Size = New System.Drawing.Size(129, 17)
        Me.lblFuelSurfaceArea.TabIndex = 133
        Me.lblFuelSurfaceArea.Text = "Fuel Surface Area (m2)"
        Me.lblFuelSurfaceArea.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.lblFuelSurfaceArea.Visible = False
        '
        'lblLHoG
        '
        Me.lblLHoG.BackColor = System.Drawing.SystemColors.Control
        Me.lblLHoG.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblLHoG.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLHoG.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLHoG.Location = New System.Drawing.Point(16, 56)
        Me.lblLHoG.Name = "lblLHoG"
        Me.lblLHoG.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblLHoG.Size = New System.Drawing.Size(161, 17)
        Me.lblLHoG.TabIndex = 132
        Me.lblLHoG.Text = "Fuel Heat of Gasification (kJ/g)"
        Me.lblLHoG.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.lblLHoG.Visible = False
        '
        'txtNodes
        '
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'frmOptions1
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(523, 437)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.SSTab2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(3, 18)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmOptions1"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Settings"
        Me.SSTab2.ResumeLayout(False)
        Me._SSTab2_TabPage0.ResumeLayout(False)
        Me._Frame1_0.ResumeLayout(False)
        Me._Frame1_0.PerformLayout()
        Me._SSTab2_TabPage1.ResumeLayout(False)
        Me.Frame28.ResumeLayout(False)
        Me.Frame23.ResumeLayout(False)
        Me.Frame22.ResumeLayout(False)
        Me.Frame21.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me._SSTab2_TabPage2.ResumeLayout(False)
        Me.Fr15combparam.ResumeLayout(False)
        Me.Fr15combparam.PerformLayout()
        Me._SSTab2_TabPage4.ResumeLayout(False)
        Me._Frame5_1.ResumeLayout(False)
        Me._Frame5_1.PerformLayout()
        Me._SSTab2_TabPage5.ResumeLayout(False)
        Me.GB_gasmodel.ResumeLayout(False)
        Me.GB_gasmodel.PerformLayout()
        Me.Frame15.ResumeLayout(False)
        Me.Frame7.ResumeLayout(False)
        Me.Frame7.PerformLayout()
        Me.GB_activity.ResumeLayout(False)
        Me.Frame6.ResumeLayout(False)
        Me.Frame6.PerformLayout()
        Me._SSTab2_TabPage6.ResumeLayout(False)
        Me._Frame14_0.ResumeLayout(False)
        Me._Frame14_0.PerformLayout()
        Me.Frame13.ResumeLayout(False)
        Me.Frame13.PerformLayout()
        Me.Frame12.ResumeLayout(False)
        Me._SSTab2_TabPage7.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me.Frame3.PerformLayout()
        Me.Frame19.ResumeLayout(False)
        Me.Frame19.PerformLayout()
        Me._SSTab2_TabPage8.ResumeLayout(False)
        Me.Frame16.ResumeLayout(False)
        Me.GB_flashovercriteria.ResumeLayout(False)
        Me.GB_flashovercriteria.PerformLayout()
        Me.Frame17.ResumeLayout(False)
        Me.Frame18.ResumeLayout(False)
        Me.Frame18.PerformLayout()
        Me._SSTab2_TabPage10.ResumeLayout(False)
        Me.Frame26.ResumeLayout(False)
        Me.Frame26.PerformLayout()
        Me.Frame27.ResumeLayout(False)
        Me.Frame27.PerformLayout()
        Me._SSTab2_TabPage9.ResumeLayout(False)
        Me.Frame25.ResumeLayout(False)
        Me.Frame20.ResumeLayout(False)
        Me.Frame20.PerformLayout()
        CType(Me.Frame1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Frame14, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Frame5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtNodes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GB_flashovercriteria As System.Windows.Forms.GroupBox
    Friend WithEvents optFOtemp As System.Windows.Forms.RadioButton
    Friend WithEvents optFOflux As System.Windows.Forms.RadioButton
    Friend WithEvents cmdDist_interiortemp As System.Windows.Forms.Button
    Friend WithEvents cmdDist_exteriortemp As System.Windows.Forms.Button
    Friend WithEvents cmdDist_RH As System.Windows.Forms.Button
    Friend WithEvents cmdDist_HOC As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents GB_gasmodel As System.Windows.Forms.GroupBox
    Friend WithEvents optFEDGeneral As System.Windows.Forms.RadioButton
    Friend WithEvents optFEDCO As System.Windows.Forms.RadioButton
    Friend WithEvents txtErrorControlVentFlow As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtExcessFuelFactor As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtCribheight As System.Windows.Forms.TextBox
    Friend WithEvents cmdWoodOption As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtStoich As System.Windows.Forms.TextBox
#End Region
End Class