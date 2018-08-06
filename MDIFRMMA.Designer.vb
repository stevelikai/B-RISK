<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class MDIFrmMain
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
    Public WithEvents mnuNewFile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuISO9705 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuFileOpen As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCSV As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuExcel As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuGlassExcel As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuBREAK1 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuExport As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSaveModelData As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuseparator_1 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuSelectFolder As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuRunBatch As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuUpdateVersionInputFiles As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuBatchFiles As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPageSetup As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPrinter As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPrintGraph As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuTextFile As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuPrint As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuseparator1_2 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuExit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuRoom As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFireSpec As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCreateConeFile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuseparator4 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuPrintout As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuGraphsOn_HRR As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuGraphsOff As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuRunTimeGraphs As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuGraphsVisible As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuseparator5_1 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuSimulation As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuDmpFile As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuOptions As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuViewInput As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuViewRTB As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuViewEndPoints As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuView_CVent As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuWVent As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuView As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuMultiGraphs As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLayerHeight As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLayerTemp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuGraphLowerLayerTemp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCeilingJetTemp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuMaxCJettemp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuGasTemperatures As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHeat As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuMassLossGraph As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuMassLeft As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuPlumeGraph As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuwallupper As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuwalllower As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuWallFlow As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFlowtoLower As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFlowtoUpper As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFlowOut As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuVentFlowGraph As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuO2Upper As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuO2Lower As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuOxygenGraph As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCO2Upper As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCO2Lower As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCO2Graph As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCOUpper As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCOLower As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCOGraph As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHCNupper As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHCNlower As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHCNgraph As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuH2OUpper As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuH2OLower As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuH2OGraph As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuODupper As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuODlower As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuODcjet As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuOD As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuEXTu As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuExtl As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuExtcjet As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuExtinction As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuTUHCUpper As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuTUHVLower As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuUnburnedFuel As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuSpecies As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFEDgases As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFEDradiation As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFEDgraph As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuVisibilityGraph As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuVentFires As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuRadiationTarget As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuRadiationFloor As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuTargetRadiationGraph As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuDetectorGraph As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuWallTempGraph As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuGraphLowerWallTemp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCeilingTempGraph As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuGraphFloorTemp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuPressure As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFloor_YP As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFloor_YB As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFloor_X As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnufloorspread As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuYPyrolysis As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuY_Burnout As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuXPyrolysis As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuZPyrolysis As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuYvelocity As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuXvelocity As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuQuintiereGraph As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuConeHRRwall As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuConeHRRceiling As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuConeHRRfloor As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuWallIgn As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCeilingIgn As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFloorIgn As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuWallHoG As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCeilingHoG As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFloorHoG As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuConeHRR As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuSPR As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuGER As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuGlassTempGraph As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuGraphs As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHelpContents As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuTechRef As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuVerification As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuEULA As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuVersion As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuAnotherSep As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuAbout As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuTerminate As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuContinue As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCancelBatch As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuStop1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MDIFrmMain))
        Dim ChartArea5 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Legend5 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim Series5 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Dim ChartArea6 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Legend6 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim Series6 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Dim ChartArea7 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Legend7 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim Series7 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Dim ChartArea8 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Legend8 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim Series8 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.UserModeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NZBCVM2ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RiskSimulatorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TalkToEVACNZToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNewFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuISO9705 = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenBaseModelToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveBaseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveBaseModelAsAsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LoadIterationToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuSetRiskdataFolder = New System.Windows.Forms.ToolStripMenuItem()
        Me.UtilitiesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenBRANZFIREModFileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ChangeFireDatabaseFileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ChangeMaterialsDatabaseFilethermalmdbToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DevelopmentKeyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EvacuatioNZSettingsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.BatchFilesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuExport = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCSV = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuExcel = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuGlassExcel = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuBREAK1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.TempToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.BoundaryNodeTemperaturesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UpperWallToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CeilingToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuseparator_1 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuPageSetup = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuPrint = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuPrinter = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuPrintGraph = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTextFile = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuseparator1_2 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileOpen = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuSaveModelData = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuBatchFiles = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuSelectFolder = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuRunBatch = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuUpdateVersionInputFiles = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuInputs = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuRoom = New System.Windows.Forms.ToolStripMenuItem()
        Me.RoomsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SurfaceMaterialsDatabaseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.VentsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.WallVentsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CeilingVentsNewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MechanicalFansToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.CeilingVentsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDetection = New System.Windows.Forms.ToolStripMenuItem()
        Me.SprinklersHeatDetectorsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SmokeDetectorsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SmokeDetectorsnewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFireSpec = New System.Windows.Forms.ToolStripMenuItem()
        Me.SelectFireToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AdditionalCombustionParametersToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.COSootToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FireObjectDatabaseToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTenability = New System.Windows.Forms.ToolStripMenuItem()
        Me.TenabilityParametersToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FEDEgressPathToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFlameSpread = New System.Windows.Forms.ToolStripMenuItem()
        Me.SetingsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MaterialConeFileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOptions = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCreateConeFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuseparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuPrintout = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuRunTimeGraphs = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuGraphsOff = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuGraphsOn_HRR = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuGraphsVisible = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuseparator5_1 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuSimulation = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDmpFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.SolversToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AmbientConditionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CompartmentEffectsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PostflashoverBehaviourToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ModelPhysicsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuView = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuViewInput = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuViewRTB = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuView_CVent = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuWVent = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuViewEndPoints = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuGraphs = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuMultiGraphs = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLayerHeight = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuGasTemperatures = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLayerTemp = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuGraphLowerLayerTemp = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCeilingJetTemp = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuMaxCJettemp = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHeat = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuMassLossGraph = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuMassLeft = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuPlumeGraph = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuWallFlow = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuwallupper = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuwalllower = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuVentFlowGraph = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFlowtoLower = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFlowtoUpper = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFlowOut = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuSpecies = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOxygenGraph = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuO2Upper = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuO2Lower = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCO2Graph = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCO2Upper = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCO2Lower = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCOGraph = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCOUpper = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCOLower = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHCNgraph = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHCNupper = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHCNlower = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuH2OGraph = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuH2OUpper = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuH2OLower = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOD = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuODupper = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuODlower = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuODcjet = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuExtinction = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEXTu = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuExtl = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuExtcjet = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuUnburnedFuel = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTUHCUpper = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTUHVLower = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFEDgraph = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFEDgases = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFEDradiation = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuVisibilityGraph = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuVentFires = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTargetRadiationGraph = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuRadiationTarget = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuRadiationFloor = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDetectorGraph = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuSurfaceTemperatures = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCeilingTempGraph = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuWallTempGraph = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuGraphLowerWallTemp = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuGraphFloorTemp = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuSurfaceFluxes = New System.Windows.Forms.ToolStripMenuItem()
        Me.IncidentHeatFluxesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem10 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem11 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem12 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem13 = New System.Windows.Forms.ToolStripMenuItem()
        Me.NetHeatFluxesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem6 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem7 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem8 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem9 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuConvectHTC = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem24 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem25 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem26 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem27 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuAST = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem14 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem15 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem16 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem17 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem18 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem19 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem20 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem21 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem22 = New System.Windows.Forms.ToolStripMenuItem()
        Me.uwallchardepth = New System.Windows.Forms.ToolStripMenuItem()
        Me.ceilingchardepth = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuPressure = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCeilingJet = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem3 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem4 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem5 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnufloorspread = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFloor_YP = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFloor_YB = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFloor_X = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuQuintiereGraph = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuYPyrolysis = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuY_Burnout = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuXPyrolysis = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuZPyrolysis = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuYvelocity = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuXvelocity = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuConeHRR = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuConeHRRwall = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuConeHRRceiling = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuConeHRRfloor = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuWallIgn = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCeilingIgn = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFloorIgn = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuWallHoG = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCeilingHoG = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFloorHoG = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuSPR = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuGER = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuUnconstrainedHeat = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuGlassTempGraph = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNormalisedHeatLoad = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem23 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem28 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem29 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuPlotWoodBurningRate = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem31 = New System.Windows.Forms.ToolStripMenuItem()
        Me.burningrate = New System.Windows.Forms.ToolStripMenuItem()
        Me.areashrinkage = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuSmokeView = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CreateSmokeViewDataFileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelpContents = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTechRef = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuVerification = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEULA = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuVersion = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuAnotherSep = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuAbout = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuStop1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTerminate = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuContinue = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCancelBatch = New System.Windows.Forms.ToolStripMenuItem()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStripSeparator = New System.Windows.Forms.ToolStripSeparator()
        Me.SaveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveAsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.PrintToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PrintPreviewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UndoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RedoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.CutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CopyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PasteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.SelectAllToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CustomizeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContentsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.IndexToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SearchToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStripSeparator6 = New System.Windows.Forms.ToolStripSeparator()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CeilingToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.AverageToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.NewToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.OpenToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.SaveToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.PrintToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton5 = New System.Windows.Forms.ToolStripButton()
        Me.toolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuRun = New System.Windows.Forms.ToolStripButton()
        Me.mnuStop = New System.Windows.Forms.ToolStripButton()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel3 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel4 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Process1 = New System.Diagnostics.Process()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.ChartRuntime4 = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.ChartRuntime2 = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.ChartRuntime3 = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.ChartRuntime1 = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.MainMenu1.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.ChartRuntime4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ChartRuntime2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ChartRuntime3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ChartRuntime1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuInputs, Me.mnuRoom, Me.VentsToolStripMenuItem, Me.mnuDetection, Me.mnuFireSpec, Me.mnuTenability, Me.mnuFlameSpread, Me.mnuOptions, Me.mnuView, Me.mnuGraphs, Me.mnuSmokeView, Me.mnuHelp, Me.mnuStop1, Me.FileToolStripMenuItem, Me.EditToolStripMenuItem, Me.ToolsToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(1249, 24)
        Me.MainMenu1.TabIndex = 5
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UserModeToolStripMenuItem, Me.mnuNewFile, Me.mnuISO9705, Me.OpenBaseModelToolStripMenuItem, Me.SaveBaseToolStripMenuItem, Me.SaveBaseModelAsAsToolStripMenuItem, Me.LoadIterationToolStripMenuItem, Me.mnuSetRiskdataFolder, Me.UtilitiesToolStripMenuItem, Me.mnuExport, Me._mnuseparator_1, Me.mnuPageSetup, Me.mnuPrint, Me._mnuseparator1_2, Me.mnuExit, Me.ToolStripMenuItem1, Me.mnuFileOpen, Me.mnuSaveModelData, Me.mnuBatchFiles})
        Me.mnuFile.Image = Global.BRISK.My.Resources.Resources.Company
        Me.mnuFile.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(53, 20)
        Me.mnuFile.Text = "File"
        Me.mnuFile.ToolTipText = "File menu"
        '
        'UserModeToolStripMenuItem
        '
        Me.UserModeToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NZBCVM2ToolStripMenuItem, Me.RiskSimulatorToolStripMenuItem, Me.TalkToEVACNZToolStripMenuItem})
        Me.UserModeToolStripMenuItem.Name = "UserModeToolStripMenuItem"
        Me.UserModeToolStripMenuItem.Size = New System.Drawing.Size(251, 22)
        Me.UserModeToolStripMenuItem.Text = "User Mode"
        '
        'NZBCVM2ToolStripMenuItem
        '
        Me.NZBCVM2ToolStripMenuItem.Name = "NZBCVM2ToolStripMenuItem"
        Me.NZBCVM2ToolStripMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.NZBCVM2ToolStripMenuItem.Text = "NZBC - VM2"
        '
        'RiskSimulatorToolStripMenuItem
        '
        Me.RiskSimulatorToolStripMenuItem.Checked = True
        Me.RiskSimulatorToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked
        Me.RiskSimulatorToolStripMenuItem.Name = "RiskSimulatorToolStripMenuItem"
        Me.RiskSimulatorToolStripMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.RiskSimulatorToolStripMenuItem.Text = "Risk Simulator"
        '
        'TalkToEVACNZToolStripMenuItem
        '
        Me.TalkToEVACNZToolStripMenuItem.CheckOnClick = True
        Me.TalkToEVACNZToolStripMenuItem.Name = "TalkToEVACNZToolStripMenuItem"
        Me.TalkToEVACNZToolStripMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.TalkToEVACNZToolStripMenuItem.Text = "Talk to EvacuatioNZ"
        '
        'mnuNewFile
        '
        Me.mnuNewFile.Name = "mnuNewFile"
        Me.mnuNewFile.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N), System.Windows.Forms.Keys)
        Me.mnuNewFile.Size = New System.Drawing.Size(251, 22)
        Me.mnuNewFile.Text = "&New"
        '
        'mnuISO9705
        '
        Me.mnuISO9705.Name = "mnuISO9705"
        Me.mnuISO9705.Size = New System.Drawing.Size(251, 22)
        Me.mnuISO9705.Text = "New ISO9705 Simulation"
        '
        'OpenBaseModelToolStripMenuItem
        '
        Me.OpenBaseModelToolStripMenuItem.Name = "OpenBaseModelToolStripMenuItem"
        Me.OpenBaseModelToolStripMenuItem.Size = New System.Drawing.Size(251, 22)
        Me.OpenBaseModelToolStripMenuItem.Text = "Open Base Model"
        '
        'SaveBaseToolStripMenuItem
        '
        Me.SaveBaseToolStripMenuItem.Name = "SaveBaseToolStripMenuItem"
        Me.SaveBaseToolStripMenuItem.Size = New System.Drawing.Size(251, 22)
        Me.SaveBaseToolStripMenuItem.Text = "Save Base Model"
        '
        'SaveBaseModelAsAsToolStripMenuItem
        '
        Me.SaveBaseModelAsAsToolStripMenuItem.Name = "SaveBaseModelAsAsToolStripMenuItem"
        Me.SaveBaseModelAsAsToolStripMenuItem.Size = New System.Drawing.Size(251, 22)
        Me.SaveBaseModelAsAsToolStripMenuItem.Text = "Save Base Model As"
        '
        'LoadIterationToolStripMenuItem
        '
        Me.LoadIterationToolStripMenuItem.Name = "LoadIterationToolStripMenuItem"
        Me.LoadIterationToolStripMenuItem.Size = New System.Drawing.Size(251, 22)
        Me.LoadIterationToolStripMenuItem.Text = "Load Input*.xml File"
        '
        'mnuSetRiskdataFolder
        '
        Me.mnuSetRiskdataFolder.Name = "mnuSetRiskdataFolder"
        Me.mnuSetRiskdataFolder.Size = New System.Drawing.Size(251, 22)
        Me.mnuSetRiskdataFolder.Text = "Set Default Parent Riskdata Folder"
        '
        'UtilitiesToolStripMenuItem
        '
        Me.UtilitiesToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OpenBRANZFIREModFileToolStripMenuItem, Me.ChangeFireDatabaseFileToolStripMenuItem, Me.ChangeMaterialsDatabaseFilethermalmdbToolStripMenuItem, Me.DevelopmentKeyToolStripMenuItem, Me.EvacuatioNZSettingsToolStripMenuItem, Me.BatchFilesToolStripMenuItem})
        Me.UtilitiesToolStripMenuItem.Name = "UtilitiesToolStripMenuItem"
        Me.UtilitiesToolStripMenuItem.Size = New System.Drawing.Size(251, 22)
        Me.UtilitiesToolStripMenuItem.Text = "Utilities"
        '
        'OpenBRANZFIREModFileToolStripMenuItem
        '
        Me.OpenBRANZFIREModFileToolStripMenuItem.Name = "OpenBRANZFIREModFileToolStripMenuItem"
        Me.OpenBRANZFIREModFileToolStripMenuItem.Size = New System.Drawing.Size(318, 22)
        Me.OpenBRANZFIREModFileToolStripMenuItem.Text = "Open BRANZFIRE Mod File"
        Me.OpenBRANZFIREModFileToolStripMenuItem.Visible = False
        '
        'ChangeFireDatabaseFileToolStripMenuItem
        '
        Me.ChangeFireDatabaseFileToolStripMenuItem.Name = "ChangeFireDatabaseFileToolStripMenuItem"
        Me.ChangeFireDatabaseFileToolStripMenuItem.Size = New System.Drawing.Size(318, 22)
        Me.ChangeFireDatabaseFileToolStripMenuItem.Text = "Change Fire Database File (fire.mdb)"
        '
        'ChangeMaterialsDatabaseFilethermalmdbToolStripMenuItem
        '
        Me.ChangeMaterialsDatabaseFilethermalmdbToolStripMenuItem.Name = "ChangeMaterialsDatabaseFilethermalmdbToolStripMenuItem"
        Me.ChangeMaterialsDatabaseFilethermalmdbToolStripMenuItem.Size = New System.Drawing.Size(318, 22)
        Me.ChangeMaterialsDatabaseFilethermalmdbToolStripMenuItem.Text = "Change Materials Database File (thermal.mdb)"
        '
        'DevelopmentKeyToolStripMenuItem
        '
        Me.DevelopmentKeyToolStripMenuItem.Name = "DevelopmentKeyToolStripMenuItem"
        Me.DevelopmentKeyToolStripMenuItem.Size = New System.Drawing.Size(318, 22)
        Me.DevelopmentKeyToolStripMenuItem.Text = "Development Key"
        '
        'EvacuatioNZSettingsToolStripMenuItem
        '
        Me.EvacuatioNZSettingsToolStripMenuItem.Name = "EvacuatioNZSettingsToolStripMenuItem"
        Me.EvacuatioNZSettingsToolStripMenuItem.Size = New System.Drawing.Size(318, 22)
        Me.EvacuatioNZSettingsToolStripMenuItem.Text = "EvacuatioNZ settings"
        '
        'BatchFilesToolStripMenuItem
        '
        Me.BatchFilesToolStripMenuItem.Name = "BatchFilesToolStripMenuItem"
        Me.BatchFilesToolStripMenuItem.Size = New System.Drawing.Size(318, 22)
        Me.BatchFilesToolStripMenuItem.Text = "Batch Files"
        '
        'mnuExport
        '
        Me.mnuExport.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCSV, Me.mnuExcel, Me.mnuGlassExcel, Me.mnuBREAK1, Me.TempToolStripMenuItem, Me.BoundaryNodeTemperaturesToolStripMenuItem})
        Me.mnuExport.Name = "mnuExport"
        Me.mnuExport.Size = New System.Drawing.Size(251, 22)
        Me.mnuExport.Text = "Export"
        '
        'mnuCSV
        '
        Me.mnuCSV.Name = "mnuCSV"
        Me.mnuCSV.Size = New System.Drawing.Size(271, 22)
        Me.mnuCSV.Text = "Export Results to .CSV File "
        Me.mnuCSV.Visible = False
        '
        'mnuExcel
        '
        Me.mnuExcel.Name = "mnuExcel"
        Me.mnuExcel.Size = New System.Drawing.Size(271, 22)
        Me.mnuExcel.Text = "Export Results to Microsoft Excel"
        '
        'mnuGlassExcel
        '
        Me.mnuGlassExcel.Enabled = False
        Me.mnuGlassExcel.Name = "mnuGlassExcel"
        Me.mnuGlassExcel.Size = New System.Drawing.Size(271, 22)
        Me.mnuGlassExcel.Text = "Export Glass Breakage Results to Excel"
        Me.mnuGlassExcel.Visible = False
        '
        'mnuBREAK1
        '
        Me.mnuBREAK1.Name = "mnuBREAK1"
        Me.mnuBREAK1.Size = New System.Drawing.Size(271, 22)
        Me.mnuBREAK1.Text = "Create BREAK1 input file"
        Me.mnuBREAK1.Visible = False
        '
        'TempToolStripMenuItem
        '
        Me.TempToolStripMenuItem.Enabled = False
        Me.TempToolStripMenuItem.Name = "TempToolStripMenuItem"
        Me.TempToolStripMenuItem.Size = New System.Drawing.Size(271, 22)
        Me.TempToolStripMenuItem.Text = "temp"
        Me.TempToolStripMenuItem.Visible = False
        '
        'BoundaryNodeTemperaturesToolStripMenuItem
        '
        Me.BoundaryNodeTemperaturesToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UpperWallToolStripMenuItem, Me.CeilingToolStripMenuItem2})
        Me.BoundaryNodeTemperaturesToolStripMenuItem.Name = "BoundaryNodeTemperaturesToolStripMenuItem"
        Me.BoundaryNodeTemperaturesToolStripMenuItem.Size = New System.Drawing.Size(271, 22)
        Me.BoundaryNodeTemperaturesToolStripMenuItem.Text = "Boundary node temperatures"
        '
        'UpperWallToolStripMenuItem
        '
        Me.UpperWallToolStripMenuItem.Name = "UpperWallToolStripMenuItem"
        Me.UpperWallToolStripMenuItem.Size = New System.Drawing.Size(129, 22)
        Me.UpperWallToolStripMenuItem.Text = "upper wall"
        '
        'CeilingToolStripMenuItem2
        '
        Me.CeilingToolStripMenuItem2.Name = "CeilingToolStripMenuItem2"
        Me.CeilingToolStripMenuItem2.Size = New System.Drawing.Size(129, 22)
        Me.CeilingToolStripMenuItem2.Text = "ceiling"
        '
        '_mnuseparator_1
        '
        Me._mnuseparator_1.Name = "_mnuseparator_1"
        Me._mnuseparator_1.Size = New System.Drawing.Size(248, 6)
        '
        'mnuPageSetup
        '
        Me.mnuPageSetup.Name = "mnuPageSetup"
        Me.mnuPageSetup.Size = New System.Drawing.Size(251, 22)
        Me.mnuPageSetup.Text = "Page Setup"
        '
        'mnuPrint
        '
        Me.mnuPrint.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuPrinter, Me.mnuPrintGraph, Me.mnuTextFile})
        Me.mnuPrint.Name = "mnuPrint"
        Me.mnuPrint.Size = New System.Drawing.Size(251, 22)
        Me.mnuPrint.Text = "&Print"
        Me.mnuPrint.Visible = False
        '
        'mnuPrinter
        '
        Me.mnuPrinter.Name = "mnuPrinter"
        Me.mnuPrinter.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.P), System.Windows.Forms.Keys)
        Me.mnuPrinter.Size = New System.Drawing.Size(274, 22)
        Me.mnuPrinter.Text = "Summary of Input and Results"
        '
        'mnuPrintGraph
        '
        Me.mnuPrintGraph.Name = "mnuPrintGraph"
        Me.mnuPrintGraph.Size = New System.Drawing.Size(274, 22)
        Me.mnuPrintGraph.Text = "Print Current Graph"
        Me.mnuPrintGraph.Visible = False
        '
        'mnuTextFile
        '
        Me.mnuTextFile.Enabled = False
        Me.mnuTextFile.Name = "mnuTextFile"
        Me.mnuTextFile.Size = New System.Drawing.Size(274, 22)
        Me.mnuTextFile.Text = "Print to File"
        Me.mnuTextFile.Visible = False
        '
        '_mnuseparator1_2
        '
        Me._mnuseparator1_2.Name = "_mnuseparator1_2"
        Me._mnuseparator1_2.Size = New System.Drawing.Size(248, 6)
        '
        'mnuExit
        '
        Me.mnuExit.Name = "mnuExit"
        Me.mnuExit.Size = New System.Drawing.Size(251, 22)
        Me.mnuExit.Text = "E&xit"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(251, 22)
        Me.ToolStripMenuItem1.Text = "Save Output to XML"
        Me.ToolStripMenuItem1.Visible = False
        '
        'mnuFileOpen
        '
        Me.mnuFileOpen.Name = "mnuFileOpen"
        Me.mnuFileOpen.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
        Me.mnuFileOpen.Size = New System.Drawing.Size(251, 22)
        Me.mnuFileOpen.Text = "&Open Existing Model"
        Me.mnuFileOpen.Visible = False
        '
        'mnuSaveModelData
        '
        Me.mnuSaveModelData.Name = "mnuSaveModelData"
        Me.mnuSaveModelData.Size = New System.Drawing.Size(251, 22)
        Me.mnuSaveModelData.Text = "&Save Model"
        Me.mnuSaveModelData.Visible = False
        '
        'mnuBatchFiles
        '
        Me.mnuBatchFiles.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuSelectFolder, Me.mnuRunBatch, Me.mnuUpdateVersionInputFiles})
        Me.mnuBatchFiles.Name = "mnuBatchFiles"
        Me.mnuBatchFiles.Size = New System.Drawing.Size(251, 22)
        Me.mnuBatchFiles.Text = "Batch Files"
        Me.mnuBatchFiles.Visible = False
        '
        'mnuSelectFolder
        '
        Me.mnuSelectFolder.Name = "mnuSelectFolder"
        Me.mnuSelectFolder.Size = New System.Drawing.Size(252, 22)
        Me.mnuSelectFolder.Text = "Select Batch File Folder"
        '
        'mnuRunBatch
        '
        Me.mnuRunBatch.Name = "mnuRunBatch"
        Me.mnuRunBatch.Size = New System.Drawing.Size(252, 22)
        Me.mnuRunBatch.Text = "Run Batch Files"
        '
        'mnuUpdateVersionInputFiles
        '
        Me.mnuUpdateVersionInputFiles.Name = "mnuUpdateVersionInputFiles"
        Me.mnuUpdateVersionInputFiles.Size = New System.Drawing.Size(252, 22)
        Me.mnuUpdateVersionInputFiles.Text = "Rewrite Input Files in Batch Folder"
        Me.mnuUpdateVersionInputFiles.Visible = False
        '
        'mnuInputs
        '
        Me.mnuInputs.Image = Global.BRISK.My.Resources.Resources.pic1338490349_tv
        Me.mnuInputs.Name = "mnuInputs"
        Me.mnuInputs.Size = New System.Drawing.Size(78, 20)
        Me.mnuInputs.Text = "Console"
        '
        'mnuRoom
        '
        Me.mnuRoom.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RoomsToolStripMenuItem, Me.SurfaceMaterialsDatabaseToolStripMenuItem})
        Me.mnuRoom.Image = Global.BRISK.My.Resources.Resources.pic17
        Me.mnuRoom.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.mnuRoom.Name = "mnuRoom"
        Me.mnuRoom.Size = New System.Drawing.Size(106, 20)
        Me.mnuRoom.Text = "R&oom Design"
        '
        'RoomsToolStripMenuItem
        '
        Me.RoomsToolStripMenuItem.Name = "RoomsToolStripMenuItem"
        Me.RoomsToolStripMenuItem.Size = New System.Drawing.Size(215, 22)
        Me.RoomsToolStripMenuItem.Text = "Rooms"
        '
        'SurfaceMaterialsDatabaseToolStripMenuItem
        '
        Me.SurfaceMaterialsDatabaseToolStripMenuItem.Name = "SurfaceMaterialsDatabaseToolStripMenuItem"
        Me.SurfaceMaterialsDatabaseToolStripMenuItem.Size = New System.Drawing.Size(215, 22)
        Me.SurfaceMaterialsDatabaseToolStripMenuItem.Text = "Surface Materials Database"
        '
        'VentsToolStripMenuItem
        '
        Me.VentsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.WallVentsToolStripMenuItem, Me.CeilingVentsNewToolStripMenuItem, Me.MechanicalFansToolStripMenuItem, Me.ToolStripMenuItem2, Me.CeilingVentsToolStripMenuItem})
        Me.VentsToolStripMenuItem.Image = Global.BRISK.My.Resources.Resources.picExit
        Me.VentsToolStripMenuItem.Name = "VentsToolStripMenuItem"
        Me.VentsToolStripMenuItem.Size = New System.Drawing.Size(91, 20)
        Me.VentsToolStripMenuItem.Text = "Ventilation"
        '
        'WallVentsToolStripMenuItem
        '
        Me.WallVentsToolStripMenuItem.Name = "WallVentsToolStripMenuItem"
        Me.WallVentsToolStripMenuItem.Size = New System.Drawing.Size(190, 22)
        Me.WallVentsToolStripMenuItem.Text = "Wall Vents"
        '
        'CeilingVentsNewToolStripMenuItem
        '
        Me.CeilingVentsNewToolStripMenuItem.Name = "CeilingVentsNewToolStripMenuItem"
        Me.CeilingVentsNewToolStripMenuItem.Size = New System.Drawing.Size(190, 22)
        Me.CeilingVentsNewToolStripMenuItem.Text = "Ceiling Vents"
        '
        'MechanicalFansToolStripMenuItem
        '
        Me.MechanicalFansToolStripMenuItem.Name = "MechanicalFansToolStripMenuItem"
        Me.MechanicalFansToolStripMenuItem.Size = New System.Drawing.Size(190, 22)
        Me.MechanicalFansToolStripMenuItem.Text = "Mechanical Fans (old)"
        Me.MechanicalFansToolStripMenuItem.Visible = False
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(190, 22)
        Me.ToolStripMenuItem2.Text = "Mech Ventilation"
        '
        'CeilingVentsToolStripMenuItem
        '
        Me.CeilingVentsToolStripMenuItem.Name = "CeilingVentsToolStripMenuItem"
        Me.CeilingVentsToolStripMenuItem.Size = New System.Drawing.Size(190, 22)
        Me.CeilingVentsToolStripMenuItem.Text = "Ceiling Vents"
        Me.CeilingVentsToolStripMenuItem.Visible = False
        '
        'mnuDetection
        '
        Me.mnuDetection.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SprinklersHeatDetectorsToolStripMenuItem, Me.SmokeDetectorsToolStripMenuItem, Me.SmokeDetectorsnewToolStripMenuItem})
        Me.mnuDetection.Image = Global.BRISK.My.Resources.Resources.sprinkler
        Me.mnuDetection.Name = "mnuDetection"
        Me.mnuDetection.Size = New System.Drawing.Size(75, 20)
        Me.mnuDetection.Text = "Sensors"
        '
        'SprinklersHeatDetectorsToolStripMenuItem
        '
        Me.SprinklersHeatDetectorsToolStripMenuItem.Name = "SprinklersHeatDetectorsToolStripMenuItem"
        Me.SprinklersHeatDetectorsToolStripMenuItem.Size = New System.Drawing.Size(214, 22)
        Me.SprinklersHeatDetectorsToolStripMenuItem.Text = "Sprinklers / Heat Detectors"
        '
        'SmokeDetectorsToolStripMenuItem
        '
        Me.SmokeDetectorsToolStripMenuItem.Name = "SmokeDetectorsToolStripMenuItem"
        Me.SmokeDetectorsToolStripMenuItem.Size = New System.Drawing.Size(214, 22)
        Me.SmokeDetectorsToolStripMenuItem.Text = "Smoke Detectors (old)"
        Me.SmokeDetectorsToolStripMenuItem.Visible = False
        '
        'SmokeDetectorsnewToolStripMenuItem
        '
        Me.SmokeDetectorsnewToolStripMenuItem.Name = "SmokeDetectorsnewToolStripMenuItem"
        Me.SmokeDetectorsnewToolStripMenuItem.Size = New System.Drawing.Size(214, 22)
        Me.SmokeDetectorsnewToolStripMenuItem.Text = "Smoke Detectors"
        '
        'mnuFireSpec
        '
        Me.mnuFireSpec.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SelectFireToolStripMenuItem, Me.AdditionalCombustionParametersToolStripMenuItem, Me.COSootToolStripMenuItem, Me.FireObjectDatabaseToolStripMenuItem1})
        Me.mnuFireSpec.Image = Global.BRISK.My.Resources.Resources.pic31
        Me.mnuFireSpec.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.mnuFireSpec.Name = "mnuFireSpec"
        Me.mnuFireSpec.Size = New System.Drawing.Size(125, 20)
        Me.mnuFireSpec.Text = "Fire Specification"
        '
        'SelectFireToolStripMenuItem
        '
        Me.SelectFireToolStripMenuItem.Name = "SelectFireToolStripMenuItem"
        Me.SelectFireToolStripMenuItem.Size = New System.Drawing.Size(260, 22)
        Me.SelectFireToolStripMenuItem.Text = "Select Fire"
        '
        'AdditionalCombustionParametersToolStripMenuItem
        '
        Me.AdditionalCombustionParametersToolStripMenuItem.Name = "AdditionalCombustionParametersToolStripMenuItem"
        Me.AdditionalCombustionParametersToolStripMenuItem.Size = New System.Drawing.Size(260, 22)
        Me.AdditionalCombustionParametersToolStripMenuItem.Text = "Additional Combustion Parameters"
        '
        'COSootToolStripMenuItem
        '
        Me.COSootToolStripMenuItem.Name = "COSootToolStripMenuItem"
        Me.COSootToolStripMenuItem.Size = New System.Drawing.Size(260, 22)
        Me.COSootToolStripMenuItem.Text = "CO/Soot "
        '
        'FireObjectDatabaseToolStripMenuItem1
        '
        Me.FireObjectDatabaseToolStripMenuItem1.Name = "FireObjectDatabaseToolStripMenuItem1"
        Me.FireObjectDatabaseToolStripMenuItem1.Size = New System.Drawing.Size(260, 22)
        Me.FireObjectDatabaseToolStripMenuItem1.Text = "Fire Object Database"
        '
        'mnuTenability
        '
        Me.mnuTenability.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TenabilityParametersToolStripMenuItem, Me.FEDEgressPathToolStripMenuItem})
        Me.mnuTenability.Image = Global.BRISK.My.Resources.Resources.Radiation
        Me.mnuTenability.Name = "mnuTenability"
        Me.mnuTenability.Size = New System.Drawing.Size(86, 20)
        Me.mnuTenability.Text = "Tenability"
        '
        'TenabilityParametersToolStripMenuItem
        '
        Me.TenabilityParametersToolStripMenuItem.Name = "TenabilityParametersToolStripMenuItem"
        Me.TenabilityParametersToolStripMenuItem.Size = New System.Drawing.Size(187, 22)
        Me.TenabilityParametersToolStripMenuItem.Text = "Tenability parameters"
        '
        'FEDEgressPathToolStripMenuItem
        '
        Me.FEDEgressPathToolStripMenuItem.Name = "FEDEgressPathToolStripMenuItem"
        Me.FEDEgressPathToolStripMenuItem.Size = New System.Drawing.Size(187, 22)
        Me.FEDEgressPathToolStripMenuItem.Text = "FED egress path"
        '
        'mnuFlameSpread
        '
        Me.mnuFlameSpread.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SetingsToolStripMenuItem, Me.MaterialConeFileToolStripMenuItem})
        Me.mnuFlameSpread.Image = Global.BRISK.My.Resources.Resources.pic1338490003_55
        Me.mnuFlameSpread.Name = "mnuFlameSpread"
        Me.mnuFlameSpread.Size = New System.Drawing.Size(106, 20)
        Me.mnuFlameSpread.Text = "Flame Spread"
        '
        'SetingsToolStripMenuItem
        '
        Me.SetingsToolStripMenuItem.Name = "SetingsToolStripMenuItem"
        Me.SetingsToolStripMenuItem.Size = New System.Drawing.Size(190, 22)
        Me.SetingsToolStripMenuItem.Text = "Flame Spread Settings"
        '
        'MaterialConeFileToolStripMenuItem
        '
        Me.MaterialConeFileToolStripMenuItem.Name = "MaterialConeFileToolStripMenuItem"
        Me.MaterialConeFileToolStripMenuItem.Size = New System.Drawing.Size(190, 22)
        Me.MaterialConeFileToolStripMenuItem.Text = "Material Cone File"
        '
        'mnuOptions
        '
        Me.mnuOptions.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCreateConeFile, Me.mnuseparator4, Me.mnuPrintout, Me.mnuRunTimeGraphs, Me.mnuGraphsVisible, Me._mnuseparator5_1, Me.mnuSimulation, Me.mnuDmpFile, Me.SolversToolStripMenuItem, Me.AmbientConditionsToolStripMenuItem, Me.CompartmentEffectsToolStripMenuItem, Me.PostflashoverBehaviourToolStripMenuItem, Me.ModelPhysicsToolStripMenuItem})
        Me.mnuOptions.Image = Global.BRISK.My.Resources.Resources.pic73
        Me.mnuOptions.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.mnuOptions.Name = "mnuOptions"
        Me.mnuOptions.Size = New System.Drawing.Size(105, 20)
        Me.mnuOptions.Text = "Misc Settings"
        '
        'mnuCreateConeFile
        '
        Me.mnuCreateConeFile.Name = "mnuCreateConeFile"
        Me.mnuCreateConeFile.Size = New System.Drawing.Size(233, 22)
        Me.mnuCreateConeFile.Text = "Create Material Cone Data File"
        Me.mnuCreateConeFile.Visible = False
        '
        'mnuseparator4
        '
        Me.mnuseparator4.Name = "mnuseparator4"
        Me.mnuseparator4.Size = New System.Drawing.Size(230, 6)
        '
        'mnuPrintout
        '
        Me.mnuPrintout.Name = "mnuPrintout"
        Me.mnuPrintout.Size = New System.Drawing.Size(233, 22)
        Me.mnuPrintout.Text = "Printout Options"
        Me.mnuPrintout.Visible = False
        '
        'mnuRunTimeGraphs
        '
        Me.mnuRunTimeGraphs.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuGraphsOff, Me.mnuGraphsOn_HRR})
        Me.mnuRunTimeGraphs.Name = "mnuRunTimeGraphs"
        Me.mnuRunTimeGraphs.Size = New System.Drawing.Size(233, 22)
        Me.mnuRunTimeGraphs.Text = "Run Time Graphs"
        Me.mnuRunTimeGraphs.Visible = False
        '
        'mnuGraphsOff
        '
        Me.mnuGraphsOff.Checked = True
        Me.mnuGraphsOff.CheckOnClick = True
        Me.mnuGraphsOff.CheckState = System.Windows.Forms.CheckState.Checked
        Me.mnuGraphsOff.Name = "mnuGraphsOff"
        Me.mnuGraphsOff.Size = New System.Drawing.Size(170, 22)
        Me.mnuGraphsOff.Text = "Off"
        '
        'mnuGraphsOn_HRR
        '
        Me.mnuGraphsOn_HRR.CheckOnClick = True
        Me.mnuGraphsOn_HRR.Name = "mnuGraphsOn_HRR"
        Me.mnuGraphsOn_HRR.Size = New System.Drawing.Size(170, 22)
        Me.mnuGraphsOn_HRR.Text = "Heat Release Rate "
        '
        'mnuGraphsVisible
        '
        Me.mnuGraphsVisible.Checked = True
        Me.mnuGraphsVisible.CheckState = System.Windows.Forms.CheckState.Checked
        Me.mnuGraphsVisible.Name = "mnuGraphsVisible"
        Me.mnuGraphsVisible.Size = New System.Drawing.Size(233, 22)
        Me.mnuGraphsVisible.Text = "Graphs Visible"
        Me.mnuGraphsVisible.Visible = False
        '
        '_mnuseparator5_1
        '
        Me._mnuseparator5_1.Name = "_mnuseparator5_1"
        Me._mnuseparator5_1.Size = New System.Drawing.Size(230, 6)
        '
        'mnuSimulation
        '
        Me.mnuSimulation.Name = "mnuSimulation"
        Me.mnuSimulation.Size = New System.Drawing.Size(233, 22)
        Me.mnuSimulation.Text = "Describe Project"
        '
        'mnuDmpFile
        '
        Me.mnuDmpFile.Name = "mnuDmpFile"
        Me.mnuDmpFile.Size = New System.Drawing.Size(233, 22)
        Me.mnuDmpFile.Text = "Dump File"
        Me.mnuDmpFile.Visible = False
        '
        'SolversToolStripMenuItem
        '
        Me.SolversToolStripMenuItem.Name = "SolversToolStripMenuItem"
        Me.SolversToolStripMenuItem.Size = New System.Drawing.Size(233, 22)
        Me.SolversToolStripMenuItem.Text = "Solvers"
        '
        'AmbientConditionsToolStripMenuItem
        '
        Me.AmbientConditionsToolStripMenuItem.Name = "AmbientConditionsToolStripMenuItem"
        Me.AmbientConditionsToolStripMenuItem.Size = New System.Drawing.Size(233, 22)
        Me.AmbientConditionsToolStripMenuItem.Text = "Ambient Conditions"
        '
        'CompartmentEffectsToolStripMenuItem
        '
        Me.CompartmentEffectsToolStripMenuItem.Name = "CompartmentEffectsToolStripMenuItem"
        Me.CompartmentEffectsToolStripMenuItem.Size = New System.Drawing.Size(233, 22)
        Me.CompartmentEffectsToolStripMenuItem.Text = "Compartment Effects"
        '
        'PostflashoverBehaviourToolStripMenuItem
        '
        Me.PostflashoverBehaviourToolStripMenuItem.Name = "PostflashoverBehaviourToolStripMenuItem"
        Me.PostflashoverBehaviourToolStripMenuItem.Size = New System.Drawing.Size(233, 22)
        Me.PostflashoverBehaviourToolStripMenuItem.Text = "Postflashover Behaviour"
        '
        'ModelPhysicsToolStripMenuItem
        '
        Me.ModelPhysicsToolStripMenuItem.Name = "ModelPhysicsToolStripMenuItem"
        Me.ModelPhysicsToolStripMenuItem.Size = New System.Drawing.Size(233, 22)
        Me.ModelPhysicsToolStripMenuItem.Text = "Model Physics"
        '
        'mnuView
        '
        Me.mnuView.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuViewInput, Me.mnuViewRTB, Me.mnuView_CVent, Me.mnuWVent, Me.mnuViewEndPoints})
        Me.mnuView.Image = Global.BRISK.My.Resources.Resources.View
        Me.mnuView.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.mnuView.Name = "mnuView"
        Me.mnuView.Size = New System.Drawing.Size(60, 20)
        Me.mnuView.Text = "&View"
        '
        'mnuViewInput
        '
        Me.mnuViewInput.Name = "mnuViewInput"
        Me.mnuViewInput.Size = New System.Drawing.Size(220, 22)
        Me.mnuViewInput.Text = "View Input"
        '
        'mnuViewRTB
        '
        Me.mnuViewRTB.Name = "mnuViewRTB"
        Me.mnuViewRTB.Size = New System.Drawing.Size(220, 22)
        Me.mnuViewRTB.Text = "View Results"
        '
        'mnuView_CVent
        '
        Me.mnuView_CVent.Name = "mnuView_CVent"
        Me.mnuView_CVent.Size = New System.Drawing.Size(220, 22)
        Me.mnuView_CVent.Text = "View Ceiling Vent Flow Data"
        '
        'mnuWVent
        '
        Me.mnuWVent.Name = "mnuWVent"
        Me.mnuWVent.Size = New System.Drawing.Size(220, 22)
        Me.mnuWVent.Text = "View Wall Vent Flow Data"
        '
        'mnuViewEndPoints
        '
        Me.mnuViewEndPoints.Name = "mnuViewEndPoints"
        Me.mnuViewEndPoints.Size = New System.Drawing.Size(220, 22)
        Me.mnuViewEndPoints.Text = "View End-Points "
        Me.mnuViewEndPoints.Visible = False
        '
        'mnuGraphs
        '
        Me.mnuGraphs.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuMultiGraphs, Me.mnuLayerHeight, Me.mnuGasTemperatures, Me.mnuHeat, Me.mnuMassLossGraph, Me.mnuMassLeft, Me.mnuPlumeGraph, Me.mnuWallFlow, Me.mnuVentFlowGraph, Me.mnuSpecies, Me.mnuFEDgraph, Me.mnuVisibilityGraph, Me.mnuVentFires, Me.mnuTargetRadiationGraph, Me.mnuDetectorGraph, Me.mnuSurfaceTemperatures, Me.mnuSurfaceFluxes, Me.mnuConvectHTC, Me.mnuAST, Me.ToolStripMenuItem18, Me.mnuPressure, Me.mnuCeilingJet, Me.mnufloorspread, Me.mnuQuintiereGraph, Me.mnuConeHRR, Me.mnuSPR, Me.mnuGER, Me.mnuUnconstrainedHeat, Me.mnuGlassTempGraph, Me.mnuNormalisedHeatLoad, Me.ToolStripMenuItem23, Me.ToolStripMenuItem28, Me.ToolStripMenuItem29, Me.mnuPlotWoodBurningRate, Me.ToolStripMenuItem31})
        Me.mnuGraphs.Image = Global.BRISK.My.Resources.Resources.pic11
        Me.mnuGraphs.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.mnuGraphs.Name = "mnuGraphs"
        Me.mnuGraphs.Size = New System.Drawing.Size(131, 20)
        Me.mnuGraphs.Text = "Single Run Graphs"
        '
        'mnuMultiGraphs
        '
        Me.mnuMultiGraphs.Name = "mnuMultiGraphs"
        Me.mnuMultiGraphs.Size = New System.Drawing.Size(272, 22)
        Me.mnuMultiGraphs.Text = "Show Multiple Graphs"
        Me.mnuMultiGraphs.Visible = False
        '
        'mnuLayerHeight
        '
        Me.mnuLayerHeight.Name = "mnuLayerHeight"
        Me.mnuLayerHeight.Size = New System.Drawing.Size(272, 22)
        Me.mnuLayerHeight.Text = "Layer Height"
        '
        'mnuGasTemperatures
        '
        Me.mnuGasTemperatures.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuLayerTemp, Me.mnuGraphLowerLayerTemp, Me.mnuCeilingJetTemp, Me.mnuMaxCJettemp})
        Me.mnuGasTemperatures.Name = "mnuGasTemperatures"
        Me.mnuGasTemperatures.Size = New System.Drawing.Size(272, 22)
        Me.mnuGasTemperatures.Text = "Layer Temperatures"
        '
        'mnuLayerTemp
        '
        Me.mnuLayerTemp.Name = "mnuLayerTemp"
        Me.mnuLayerTemp.Size = New System.Drawing.Size(212, 22)
        Me.mnuLayerTemp.Text = "Upper Layer"
        '
        'mnuGraphLowerLayerTemp
        '
        Me.mnuGraphLowerLayerTemp.Name = "mnuGraphLowerLayerTemp"
        Me.mnuGraphLowerLayerTemp.Size = New System.Drawing.Size(212, 22)
        Me.mnuGraphLowerLayerTemp.Text = "Lower Layer"
        '
        'mnuCeilingJetTemp
        '
        Me.mnuCeilingJetTemp.Name = "mnuCeilingJetTemp"
        Me.mnuCeilingJetTemp.Size = New System.Drawing.Size(212, 22)
        Me.mnuCeilingJetTemp.Text = "Ceiling Jet at Link Position"
        Me.mnuCeilingJetTemp.Visible = False
        '
        'mnuMaxCJettemp
        '
        Me.mnuMaxCJettemp.Name = "mnuMaxCJettemp"
        Me.mnuMaxCJettemp.Size = New System.Drawing.Size(212, 22)
        Me.mnuMaxCJettemp.Text = "Maximum Ceiling Jet"
        Me.mnuMaxCJettemp.Visible = False
        '
        'mnuHeat
        '
        Me.mnuHeat.Name = "mnuHeat"
        Me.mnuHeat.Size = New System.Drawing.Size(272, 22)
        Me.mnuHeat.Text = "Heat Release Rate"
        '
        'mnuMassLossGraph
        '
        Me.mnuMassLossGraph.Name = "mnuMassLossGraph"
        Me.mnuMassLossGraph.Size = New System.Drawing.Size(272, 22)
        Me.mnuMassLossGraph.Text = "Mass Loss Rate"
        '
        'mnuMassLeft
        '
        Me.mnuMassLeft.Name = "mnuMassLeft"
        Me.mnuMassLeft.Size = New System.Drawing.Size(272, 22)
        Me.mnuMassLeft.Text = "Mass Loss Total"
        '
        'mnuPlumeGraph
        '
        Me.mnuPlumeGraph.Name = "mnuPlumeGraph"
        Me.mnuPlumeGraph.Size = New System.Drawing.Size(272, 22)
        Me.mnuPlumeGraph.Text = "Plume Flow"
        '
        'mnuWallFlow
        '
        Me.mnuWallFlow.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuwallupper, Me.mnuwalllower})
        Me.mnuWallFlow.Name = "mnuWallFlow"
        Me.mnuWallFlow.Size = New System.Drawing.Size(272, 22)
        Me.mnuWallFlow.Text = "Wall Convection Flow"
        Me.mnuWallFlow.Visible = False
        '
        'mnuwallupper
        '
        Me.mnuwallupper.Name = "mnuwallupper"
        Me.mnuwallupper.Size = New System.Drawing.Size(151, 22)
        Me.mnuwallupper.Text = "to Upper Layer"
        '
        'mnuwalllower
        '
        Me.mnuwalllower.Name = "mnuwalllower"
        Me.mnuwalllower.Size = New System.Drawing.Size(151, 22)
        Me.mnuwalllower.Text = "to Lower Layer"
        '
        'mnuVentFlowGraph
        '
        Me.mnuVentFlowGraph.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFlowtoLower, Me.mnuFlowtoUpper, Me.mnuFlowOut})
        Me.mnuVentFlowGraph.Name = "mnuVentFlowGraph"
        Me.mnuVentFlowGraph.Size = New System.Drawing.Size(272, 22)
        Me.mnuVentFlowGraph.Text = "Vent Flow"
        Me.mnuVentFlowGraph.Visible = False
        '
        'mnuFlowtoLower
        '
        Me.mnuFlowtoLower.Name = "mnuFlowtoLower"
        Me.mnuFlowtoLower.Size = New System.Drawing.Size(231, 22)
        Me.mnuFlowtoLower.Text = "Net Mass Flow to Lower Layer"
        '
        'mnuFlowtoUpper
        '
        Me.mnuFlowtoUpper.Name = "mnuFlowtoUpper"
        Me.mnuFlowtoUpper.Size = New System.Drawing.Size(231, 22)
        Me.mnuFlowtoUpper.Text = "Net Mass Flow to Upper Layer"
        '
        'mnuFlowOut
        '
        Me.mnuFlowOut.Name = "mnuFlowOut"
        Me.mnuFlowOut.Size = New System.Drawing.Size(231, 22)
        Me.mnuFlowOut.Text = "Volume Flow to Outside"
        '
        'mnuSpecies
        '
        Me.mnuSpecies.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuOxygenGraph, Me.mnuCO2Graph, Me.mnuCOGraph, Me.mnuHCNgraph, Me.mnuH2OGraph, Me.mnuOD, Me.mnuExtinction, Me.mnuUnburnedFuel})
        Me.mnuSpecies.Name = "mnuSpecies"
        Me.mnuSpecies.Size = New System.Drawing.Size(272, 22)
        Me.mnuSpecies.Text = "Species Concentrations"
        '
        'mnuOxygenGraph
        '
        Me.mnuOxygenGraph.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuO2Upper, Me.mnuO2Lower})
        Me.mnuOxygenGraph.Name = "mnuOxygenGraph"
        Me.mnuOxygenGraph.Size = New System.Drawing.Size(187, 22)
        Me.mnuOxygenGraph.Text = "Oxygen %"
        '
        'mnuO2Upper
        '
        Me.mnuO2Upper.Name = "mnuO2Upper"
        Me.mnuO2Upper.Size = New System.Drawing.Size(137, 22)
        Me.mnuO2Upper.Text = "Upper Layer"
        '
        'mnuO2Lower
        '
        Me.mnuO2Lower.Name = "mnuO2Lower"
        Me.mnuO2Lower.Size = New System.Drawing.Size(137, 22)
        Me.mnuO2Lower.Text = "Lower Layer"
        '
        'mnuCO2Graph
        '
        Me.mnuCO2Graph.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCO2Upper, Me.mnuCO2Lower})
        Me.mnuCO2Graph.Name = "mnuCO2Graph"
        Me.mnuCO2Graph.Size = New System.Drawing.Size(187, 22)
        Me.mnuCO2Graph.Text = "CO2 %"
        '
        'mnuCO2Upper
        '
        Me.mnuCO2Upper.Name = "mnuCO2Upper"
        Me.mnuCO2Upper.Size = New System.Drawing.Size(137, 22)
        Me.mnuCO2Upper.Text = "Upper Layer"
        '
        'mnuCO2Lower
        '
        Me.mnuCO2Lower.Name = "mnuCO2Lower"
        Me.mnuCO2Lower.Size = New System.Drawing.Size(137, 22)
        Me.mnuCO2Lower.Text = "Lower Layer"
        '
        'mnuCOGraph
        '
        Me.mnuCOGraph.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCOUpper, Me.mnuCOLower})
        Me.mnuCOGraph.Name = "mnuCOGraph"
        Me.mnuCOGraph.Size = New System.Drawing.Size(187, 22)
        Me.mnuCOGraph.Text = "CO (ppm)"
        '
        'mnuCOUpper
        '
        Me.mnuCOUpper.Name = "mnuCOUpper"
        Me.mnuCOUpper.Size = New System.Drawing.Size(137, 22)
        Me.mnuCOUpper.Text = "Upper Layer"
        '
        'mnuCOLower
        '
        Me.mnuCOLower.Name = "mnuCOLower"
        Me.mnuCOLower.Size = New System.Drawing.Size(137, 22)
        Me.mnuCOLower.Text = "Lower Layer"
        '
        'mnuHCNgraph
        '
        Me.mnuHCNgraph.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuHCNupper, Me.mnuHCNlower})
        Me.mnuHCNgraph.Name = "mnuHCNgraph"
        Me.mnuHCNgraph.Size = New System.Drawing.Size(187, 22)
        Me.mnuHCNgraph.Text = "HCN (ppm)"
        '
        'mnuHCNupper
        '
        Me.mnuHCNupper.Name = "mnuHCNupper"
        Me.mnuHCNupper.Size = New System.Drawing.Size(137, 22)
        Me.mnuHCNupper.Text = "Upper Layer"
        '
        'mnuHCNlower
        '
        Me.mnuHCNlower.Name = "mnuHCNlower"
        Me.mnuHCNlower.Size = New System.Drawing.Size(137, 22)
        Me.mnuHCNlower.Text = "Lower Layer"
        '
        'mnuH2OGraph
        '
        Me.mnuH2OGraph.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuH2OUpper, Me.mnuH2OLower})
        Me.mnuH2OGraph.Name = "mnuH2OGraph"
        Me.mnuH2OGraph.Size = New System.Drawing.Size(187, 22)
        Me.mnuH2OGraph.Text = "H2O %"
        '
        'mnuH2OUpper
        '
        Me.mnuH2OUpper.Name = "mnuH2OUpper"
        Me.mnuH2OUpper.Size = New System.Drawing.Size(137, 22)
        Me.mnuH2OUpper.Text = "Upper Layer"
        '
        'mnuH2OLower
        '
        Me.mnuH2OLower.Name = "mnuH2OLower"
        Me.mnuH2OLower.Size = New System.Drawing.Size(137, 22)
        Me.mnuH2OLower.Text = "Lower Layer"
        '
        'mnuOD
        '
        Me.mnuOD.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuODupper, Me.mnuODlower, Me.mnuODcjet})
        Me.mnuOD.Name = "mnuOD"
        Me.mnuOD.Size = New System.Drawing.Size(187, 22)
        Me.mnuOD.Text = "Optical Density"
        '
        'mnuODupper
        '
        Me.mnuODupper.Name = "mnuODupper"
        Me.mnuODupper.Size = New System.Drawing.Size(223, 22)
        Me.mnuODupper.Text = "Upper Layer"
        '
        'mnuODlower
        '
        Me.mnuODlower.Name = "mnuODlower"
        Me.mnuODlower.Size = New System.Drawing.Size(223, 22)
        Me.mnuODlower.Text = "Lower Layer"
        '
        'mnuODcjet
        '
        Me.mnuODcjet.Name = "mnuODcjet"
        Me.mnuODcjet.Size = New System.Drawing.Size(223, 22)
        Me.mnuODcjet.Text = "Ceiling Jet (Smoke Detector)"
        Me.mnuODcjet.Visible = False
        '
        'mnuExtinction
        '
        Me.mnuExtinction.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuEXTu, Me.mnuExtl, Me.mnuExtcjet})
        Me.mnuExtinction.Name = "mnuExtinction"
        Me.mnuExtinction.Size = New System.Drawing.Size(187, 22)
        Me.mnuExtinction.Text = "Extinction Coefficient"
        '
        'mnuEXTu
        '
        Me.mnuEXTu.Name = "mnuEXTu"
        Me.mnuEXTu.Size = New System.Drawing.Size(223, 22)
        Me.mnuEXTu.Text = "Upper Layer"
        '
        'mnuExtl
        '
        Me.mnuExtl.Name = "mnuExtl"
        Me.mnuExtl.Size = New System.Drawing.Size(223, 22)
        Me.mnuExtl.Text = "Lower Layer"
        '
        'mnuExtcjet
        '
        Me.mnuExtcjet.Name = "mnuExtcjet"
        Me.mnuExtcjet.Size = New System.Drawing.Size(223, 22)
        Me.mnuExtcjet.Text = "Ceiling Jet (Smoke Detector)"
        Me.mnuExtcjet.Visible = False
        '
        'mnuUnburnedFuel
        '
        Me.mnuUnburnedFuel.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuTUHCUpper, Me.mnuTUHVLower})
        Me.mnuUnburnedFuel.Name = "mnuUnburnedFuel"
        Me.mnuUnburnedFuel.Size = New System.Drawing.Size(187, 22)
        Me.mnuUnburnedFuel.Text = "Unburned Fuel"
        '
        'mnuTUHCUpper
        '
        Me.mnuTUHCUpper.Name = "mnuTUHCUpper"
        Me.mnuTUHCUpper.Size = New System.Drawing.Size(137, 22)
        Me.mnuTUHCUpper.Text = "Upper Layer"
        '
        'mnuTUHVLower
        '
        Me.mnuTUHVLower.Name = "mnuTUHVLower"
        Me.mnuTUHVLower.Size = New System.Drawing.Size(137, 22)
        Me.mnuTUHVLower.Text = "Lower Layer"
        '
        'mnuFEDgraph
        '
        Me.mnuFEDgraph.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFEDgases, Me.mnuFEDradiation})
        Me.mnuFEDgraph.Name = "mnuFEDgraph"
        Me.mnuFEDgraph.Size = New System.Drawing.Size(272, 22)
        Me.mnuFEDgraph.Text = "FED"
        '
        'mnuFEDgases
        '
        Me.mnuFEDgases.Name = "mnuFEDgases"
        Me.mnuFEDgases.Size = New System.Drawing.Size(165, 22)
        Me.mnuFEDgases.Text = "Asphyxiant Gases"
        '
        'mnuFEDradiation
        '
        Me.mnuFEDradiation.Name = "mnuFEDradiation"
        Me.mnuFEDradiation.Size = New System.Drawing.Size(165, 22)
        Me.mnuFEDradiation.Text = "Thermal"
        '
        'mnuVisibilityGraph
        '
        Me.mnuVisibilityGraph.Name = "mnuVisibilityGraph"
        Me.mnuVisibilityGraph.Size = New System.Drawing.Size(272, 22)
        Me.mnuVisibilityGraph.Text = "Visibility"
        '
        'mnuVentFires
        '
        Me.mnuVentFires.Name = "mnuVentFires"
        Me.mnuVentFires.Size = New System.Drawing.Size(272, 22)
        Me.mnuVentFires.Text = "Vent Fires"
        '
        'mnuTargetRadiationGraph
        '
        Me.mnuTargetRadiationGraph.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuRadiationTarget, Me.mnuRadiationFloor})
        Me.mnuTargetRadiationGraph.Name = "mnuTargetRadiationGraph"
        Me.mnuTargetRadiationGraph.Size = New System.Drawing.Size(272, 22)
        Me.mnuTargetRadiationGraph.Text = "Radiation"
        '
        'mnuRadiationTarget
        '
        Me.mnuRadiationTarget.Name = "mnuRadiationTarget"
        Me.mnuRadiationTarget.Size = New System.Drawing.Size(107, 22)
        Me.mnuRadiationTarget.Text = "Target"
        '
        'mnuRadiationFloor
        '
        Me.mnuRadiationFloor.Name = "mnuRadiationFloor"
        Me.mnuRadiationFloor.Size = New System.Drawing.Size(107, 22)
        Me.mnuRadiationFloor.Text = "Floor"
        '
        'mnuDetectorGraph
        '
        Me.mnuDetectorGraph.Name = "mnuDetectorGraph"
        Me.mnuDetectorGraph.Size = New System.Drawing.Size(272, 22)
        Me.mnuDetectorGraph.Text = "Detector/Sprinkler Temp"
        '
        'mnuSurfaceTemperatures
        '
        Me.mnuSurfaceTemperatures.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCeilingTempGraph, Me.mnuWallTempGraph, Me.mnuGraphLowerWallTemp, Me.mnuGraphFloorTemp})
        Me.mnuSurfaceTemperatures.Name = "mnuSurfaceTemperatures"
        Me.mnuSurfaceTemperatures.Size = New System.Drawing.Size(272, 22)
        Me.mnuSurfaceTemperatures.Text = "Surface Temperatures"
        '
        'mnuCeilingTempGraph
        '
        Me.mnuCeilingTempGraph.Name = "mnuCeilingTempGraph"
        Me.mnuCeilingTempGraph.Size = New System.Drawing.Size(132, 22)
        Me.mnuCeilingTempGraph.Text = "Ceiling"
        '
        'mnuWallTempGraph
        '
        Me.mnuWallTempGraph.Name = "mnuWallTempGraph"
        Me.mnuWallTempGraph.Size = New System.Drawing.Size(132, 22)
        Me.mnuWallTempGraph.Text = "Upper Wall"
        '
        'mnuGraphLowerWallTemp
        '
        Me.mnuGraphLowerWallTemp.Name = "mnuGraphLowerWallTemp"
        Me.mnuGraphLowerWallTemp.Size = New System.Drawing.Size(132, 22)
        Me.mnuGraphLowerWallTemp.Text = "Lower Wall"
        '
        'mnuGraphFloorTemp
        '
        Me.mnuGraphFloorTemp.Name = "mnuGraphFloorTemp"
        Me.mnuGraphFloorTemp.Size = New System.Drawing.Size(132, 22)
        Me.mnuGraphFloorTemp.Text = "Floor"
        '
        'mnuSurfaceFluxes
        '
        Me.mnuSurfaceFluxes.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.IncidentHeatFluxesToolStripMenuItem, Me.NetHeatFluxesToolStripMenuItem})
        Me.mnuSurfaceFluxes.Name = "mnuSurfaceFluxes"
        Me.mnuSurfaceFluxes.Size = New System.Drawing.Size(272, 22)
        Me.mnuSurfaceFluxes.Text = "Surface Heat Fluxes"
        '
        'IncidentHeatFluxesToolStripMenuItem
        '
        Me.IncidentHeatFluxesToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem10, Me.ToolStripMenuItem11, Me.ToolStripMenuItem12, Me.ToolStripMenuItem13})
        Me.IncidentHeatFluxesToolStripMenuItem.Name = "IncidentHeatFluxesToolStripMenuItem"
        Me.IncidentHeatFluxesToolStripMenuItem.Size = New System.Drawing.Size(205, 22)
        Me.IncidentHeatFluxesToolStripMenuItem.Text = "Incident radiant heat flux"
        '
        'ToolStripMenuItem10
        '
        Me.ToolStripMenuItem10.Name = "ToolStripMenuItem10"
        Me.ToolStripMenuItem10.Size = New System.Drawing.Size(132, 22)
        Me.ToolStripMenuItem10.Text = "Ceiling"
        '
        'ToolStripMenuItem11
        '
        Me.ToolStripMenuItem11.Name = "ToolStripMenuItem11"
        Me.ToolStripMenuItem11.Size = New System.Drawing.Size(132, 22)
        Me.ToolStripMenuItem11.Text = "Upper Wall"
        '
        'ToolStripMenuItem12
        '
        Me.ToolStripMenuItem12.Name = "ToolStripMenuItem12"
        Me.ToolStripMenuItem12.Size = New System.Drawing.Size(132, 22)
        Me.ToolStripMenuItem12.Text = "Lower Wall"
        '
        'ToolStripMenuItem13
        '
        Me.ToolStripMenuItem13.Name = "ToolStripMenuItem13"
        Me.ToolStripMenuItem13.Size = New System.Drawing.Size(132, 22)
        Me.ToolStripMenuItem13.Text = "Floor"
        '
        'NetHeatFluxesToolStripMenuItem
        '
        Me.NetHeatFluxesToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem6, Me.ToolStripMenuItem7, Me.ToolStripMenuItem8, Me.ToolStripMenuItem9})
        Me.NetHeatFluxesToolStripMenuItem.Name = "NetHeatFluxesToolStripMenuItem"
        Me.NetHeatFluxesToolStripMenuItem.Size = New System.Drawing.Size(205, 22)
        Me.NetHeatFluxesToolStripMenuItem.Text = "Net total heat flux"
        '
        'ToolStripMenuItem6
        '
        Me.ToolStripMenuItem6.Name = "ToolStripMenuItem6"
        Me.ToolStripMenuItem6.Size = New System.Drawing.Size(132, 22)
        Me.ToolStripMenuItem6.Text = "Ceiling"
        '
        'ToolStripMenuItem7
        '
        Me.ToolStripMenuItem7.Name = "ToolStripMenuItem7"
        Me.ToolStripMenuItem7.Size = New System.Drawing.Size(132, 22)
        Me.ToolStripMenuItem7.Text = "Upper wall"
        '
        'ToolStripMenuItem8
        '
        Me.ToolStripMenuItem8.Name = "ToolStripMenuItem8"
        Me.ToolStripMenuItem8.Size = New System.Drawing.Size(132, 22)
        Me.ToolStripMenuItem8.Text = "Lower Wall"
        '
        'ToolStripMenuItem9
        '
        Me.ToolStripMenuItem9.Name = "ToolStripMenuItem9"
        Me.ToolStripMenuItem9.Size = New System.Drawing.Size(132, 22)
        Me.ToolStripMenuItem9.Text = "Floor"
        '
        'mnuConvectHTC
        '
        Me.mnuConvectHTC.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem24, Me.ToolStripMenuItem25, Me.ToolStripMenuItem26, Me.ToolStripMenuItem27})
        Me.mnuConvectHTC.Name = "mnuConvectHTC"
        Me.mnuConvectHTC.Size = New System.Drawing.Size(272, 22)
        Me.mnuConvectHTC.Text = "Convective Heat Transfer Coefficients"
        '
        'ToolStripMenuItem24
        '
        Me.ToolStripMenuItem24.Name = "ToolStripMenuItem24"
        Me.ToolStripMenuItem24.Size = New System.Drawing.Size(132, 22)
        Me.ToolStripMenuItem24.Text = "Ceiling"
        '
        'ToolStripMenuItem25
        '
        Me.ToolStripMenuItem25.Name = "ToolStripMenuItem25"
        Me.ToolStripMenuItem25.Size = New System.Drawing.Size(132, 22)
        Me.ToolStripMenuItem25.Text = "Upper Wall"
        '
        'ToolStripMenuItem26
        '
        Me.ToolStripMenuItem26.Name = "ToolStripMenuItem26"
        Me.ToolStripMenuItem26.Size = New System.Drawing.Size(132, 22)
        Me.ToolStripMenuItem26.Text = "Lower Wall"
        '
        'ToolStripMenuItem27
        '
        Me.ToolStripMenuItem27.Name = "ToolStripMenuItem27"
        Me.ToolStripMenuItem27.Size = New System.Drawing.Size(132, 22)
        Me.ToolStripMenuItem27.Text = "Floor"
        '
        'mnuAST
        '
        Me.mnuAST.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem14, Me.ToolStripMenuItem15, Me.ToolStripMenuItem16, Me.ToolStripMenuItem17})
        Me.mnuAST.Name = "mnuAST"
        Me.mnuAST.Size = New System.Drawing.Size(272, 22)
        Me.mnuAST.Text = "Adiabatic Surface Temperatures"
        '
        'ToolStripMenuItem14
        '
        Me.ToolStripMenuItem14.Name = "ToolStripMenuItem14"
        Me.ToolStripMenuItem14.Size = New System.Drawing.Size(132, 22)
        Me.ToolStripMenuItem14.Text = "Ceiling"
        '
        'ToolStripMenuItem15
        '
        Me.ToolStripMenuItem15.Name = "ToolStripMenuItem15"
        Me.ToolStripMenuItem15.Size = New System.Drawing.Size(132, 22)
        Me.ToolStripMenuItem15.Text = "Upper Wall"
        '
        'ToolStripMenuItem16
        '
        Me.ToolStripMenuItem16.Name = "ToolStripMenuItem16"
        Me.ToolStripMenuItem16.Size = New System.Drawing.Size(132, 22)
        Me.ToolStripMenuItem16.Text = "Lower Wall"
        '
        'ToolStripMenuItem17
        '
        Me.ToolStripMenuItem17.Name = "ToolStripMenuItem17"
        Me.ToolStripMenuItem17.Size = New System.Drawing.Size(132, 22)
        Me.ToolStripMenuItem17.Text = "Floor"
        '
        'ToolStripMenuItem18
        '
        Me.ToolStripMenuItem18.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem19, Me.ToolStripMenuItem20, Me.ToolStripMenuItem21, Me.ToolStripMenuItem22, Me.uwallchardepth, Me.ceilingchardepth})
        Me.ToolStripMenuItem18.Name = "ToolStripMenuItem18"
        Me.ToolStripMenuItem18.Size = New System.Drawing.Size(272, 22)
        Me.ToolStripMenuItem18.Text = "Surface Internal Temperatures"
        '
        'ToolStripMenuItem19
        '
        Me.ToolStripMenuItem19.Name = "ToolStripMenuItem19"
        Me.ToolStripMenuItem19.Size = New System.Drawing.Size(179, 22)
        Me.ToolStripMenuItem19.Text = "Ceiling"
        '
        'ToolStripMenuItem20
        '
        Me.ToolStripMenuItem20.Name = "ToolStripMenuItem20"
        Me.ToolStripMenuItem20.Size = New System.Drawing.Size(179, 22)
        Me.ToolStripMenuItem20.Text = "Upper Wall"
        '
        'ToolStripMenuItem21
        '
        Me.ToolStripMenuItem21.Name = "ToolStripMenuItem21"
        Me.ToolStripMenuItem21.Size = New System.Drawing.Size(179, 22)
        Me.ToolStripMenuItem21.Text = "Lower Wall"
        '
        'ToolStripMenuItem22
        '
        Me.ToolStripMenuItem22.Name = "ToolStripMenuItem22"
        Me.ToolStripMenuItem22.Size = New System.Drawing.Size(179, 22)
        Me.ToolStripMenuItem22.Text = "Floor"
        '
        'uwallchardepth
        '
        Me.uwallchardepth.Name = "uwallchardepth"
        Me.uwallchardepth.Size = New System.Drawing.Size(179, 22)
        Me.uwallchardepth.Text = "Wall  - char depth"
        '
        'ceilingchardepth
        '
        Me.ceilingchardepth.Name = "ceilingchardepth"
        Me.ceilingchardepth.Size = New System.Drawing.Size(179, 22)
        Me.ceilingchardepth.Text = "Ceiling - char depth"
        '
        'mnuPressure
        '
        Me.mnuPressure.Name = "mnuPressure"
        Me.mnuPressure.Size = New System.Drawing.Size(272, 22)
        Me.mnuPressure.Text = "Pressure"
        '
        'mnuCeilingJet
        '
        Me.mnuCeilingJet.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem3, Me.ToolStripMenuItem4, Me.ToolStripMenuItem5})
        Me.mnuCeilingJet.Name = "mnuCeilingJet"
        Me.mnuCeilingJet.Size = New System.Drawing.Size(272, 22)
        Me.mnuCeilingJet.Text = "Ceiling Jet"
        '
        'ToolStripMenuItem3
        '
        Me.ToolStripMenuItem3.Name = "ToolStripMenuItem3"
        Me.ToolStripMenuItem3.Size = New System.Drawing.Size(298, 22)
        Me.ToolStripMenuItem3.Text = "Ceiling Jet Temperature at Sensor Location"
        '
        'ToolStripMenuItem4
        '
        Me.ToolStripMenuItem4.Name = "ToolStripMenuItem4"
        Me.ToolStripMenuItem4.Size = New System.Drawing.Size(298, 22)
        Me.ToolStripMenuItem4.Text = "Ceiling Jet Velocity at Sensor Location"
        '
        'ToolStripMenuItem5
        '
        Me.ToolStripMenuItem5.Name = "ToolStripMenuItem5"
        Me.ToolStripMenuItem5.Size = New System.Drawing.Size(298, 22)
        Me.ToolStripMenuItem5.Text = "Max Ceiling Jet Temperature"
        '
        'mnufloorspread
        '
        Me.mnufloorspread.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFloor_YP, Me.mnuFloor_YB, Me.mnuFloor_X})
        Me.mnufloorspread.Name = "mnufloorspread"
        Me.mnufloorspread.Size = New System.Drawing.Size(272, 22)
        Me.mnufloorspread.Text = "Floor Flame Spread"
        Me.mnufloorspread.Visible = False
        '
        'mnuFloor_YP
        '
        Me.mnuFloor_YP.Name = "mnuFloor_YP"
        Me.mnuFloor_YP.Size = New System.Drawing.Size(147, 22)
        Me.mnuFloor_YP.Text = "Floor Upward"
        '
        'mnuFloor_YB
        '
        Me.mnuFloor_YB.Name = "mnuFloor_YB"
        Me.mnuFloor_YB.Size = New System.Drawing.Size(147, 22)
        Me.mnuFloor_YB.Text = "Floor Burnout"
        '
        'mnuFloor_X
        '
        Me.mnuFloor_X.Name = "mnuFloor_X"
        Me.mnuFloor_X.Size = New System.Drawing.Size(147, 22)
        Me.mnuFloor_X.Text = "Floor Lateral"
        '
        'mnuQuintiereGraph
        '
        Me.mnuQuintiereGraph.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuYPyrolysis, Me.mnuY_Burnout, Me.mnuXPyrolysis, Me.mnuZPyrolysis, Me.mnuYvelocity, Me.mnuXvelocity})
        Me.mnuQuintiereGraph.Name = "mnuQuintiereGraph"
        Me.mnuQuintiereGraph.Size = New System.Drawing.Size(272, 22)
        Me.mnuQuintiereGraph.Text = "Wall-Ceiling Flame Spread"
        '
        'mnuYPyrolysis
        '
        Me.mnuYPyrolysis.Name = "mnuYPyrolysis"
        Me.mnuYPyrolysis.Size = New System.Drawing.Size(198, 22)
        Me.mnuYPyrolysis.Text = "Y_Pyrolysis Front"
        '
        'mnuY_Burnout
        '
        Me.mnuY_Burnout.Name = "mnuY_Burnout"
        Me.mnuY_Burnout.Size = New System.Drawing.Size(198, 22)
        Me.mnuY_Burnout.Text = "Y_Burnout Front"
        '
        'mnuXPyrolysis
        '
        Me.mnuXPyrolysis.Name = "mnuXPyrolysis"
        Me.mnuXPyrolysis.Size = New System.Drawing.Size(198, 22)
        Me.mnuXPyrolysis.Text = "X_Pyrolysis Front"
        '
        'mnuZPyrolysis
        '
        Me.mnuZPyrolysis.Name = "mnuZPyrolysis"
        Me.mnuZPyrolysis.Size = New System.Drawing.Size(198, 22)
        Me.mnuZPyrolysis.Text = "Z_Pyrolysis Front"
        '
        'mnuYvelocity
        '
        Me.mnuYvelocity.Name = "mnuYvelocity"
        Me.mnuYvelocity.Size = New System.Drawing.Size(198, 22)
        Me.mnuYvelocity.Text = "Upward Spread Velocity"
        '
        'mnuXvelocity
        '
        Me.mnuXvelocity.Name = "mnuXvelocity"
        Me.mnuXvelocity.Size = New System.Drawing.Size(198, 22)
        Me.mnuXvelocity.Text = "Lateral Spread Velocity"
        '
        'mnuConeHRR
        '
        Me.mnuConeHRR.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuConeHRRwall, Me.mnuConeHRRceiling, Me.mnuConeHRRfloor, Me.mnuWallIgn, Me.mnuCeilingIgn, Me.mnuFloorIgn, Me.mnuWallHoG, Me.mnuCeilingHoG, Me.mnuFloorHoG})
        Me.mnuConeHRR.Name = "mnuConeHRR"
        Me.mnuConeHRR.Size = New System.Drawing.Size(272, 22)
        Me.mnuConeHRR.Text = "Material Cone Data"
        '
        'mnuConeHRRwall
        '
        Me.mnuConeHRRwall.Name = "mnuConeHRRwall"
        Me.mnuConeHRRwall.Size = New System.Drawing.Size(200, 22)
        Me.mnuConeHRRwall.Text = "Wall Lining HRR "
        '
        'mnuConeHRRceiling
        '
        Me.mnuConeHRRceiling.Name = "mnuConeHRRceiling"
        Me.mnuConeHRRceiling.Size = New System.Drawing.Size(200, 22)
        Me.mnuConeHRRceiling.Text = "Ceiling Lining HRR"
        '
        'mnuConeHRRfloor
        '
        Me.mnuConeHRRfloor.Name = "mnuConeHRRfloor"
        Me.mnuConeHRRfloor.Size = New System.Drawing.Size(200, 22)
        Me.mnuConeHRRfloor.Text = "Floor Lining HRR"
        Me.mnuConeHRRfloor.Visible = False
        '
        'mnuWallIgn
        '
        Me.mnuWallIgn.Name = "mnuWallIgn"
        Me.mnuWallIgn.Size = New System.Drawing.Size(200, 22)
        Me.mnuWallIgn.Text = "Wall Ignition Data"
        '
        'mnuCeilingIgn
        '
        Me.mnuCeilingIgn.Name = "mnuCeilingIgn"
        Me.mnuCeilingIgn.Size = New System.Drawing.Size(200, 22)
        Me.mnuCeilingIgn.Text = "Ceiling Ignition Data"
        '
        'mnuFloorIgn
        '
        Me.mnuFloorIgn.Name = "mnuFloorIgn"
        Me.mnuFloorIgn.Size = New System.Drawing.Size(200, 22)
        Me.mnuFloorIgn.Text = "Floor Ignition Data"
        Me.mnuFloorIgn.Visible = False
        '
        'mnuWallHoG
        '
        Me.mnuWallHoG.Name = "mnuWallHoG"
        Me.mnuWallHoG.Size = New System.Drawing.Size(200, 22)
        Me.mnuWallHoG.Text = "Wall HoG Correlation"
        '
        'mnuCeilingHoG
        '
        Me.mnuCeilingHoG.Name = "mnuCeilingHoG"
        Me.mnuCeilingHoG.Size = New System.Drawing.Size(200, 22)
        Me.mnuCeilingHoG.Text = "Ceiling HoG Correlation"
        '
        'mnuFloorHoG
        '
        Me.mnuFloorHoG.Name = "mnuFloorHoG"
        Me.mnuFloorHoG.Size = New System.Drawing.Size(200, 22)
        Me.mnuFloorHoG.Text = "Floor HoG Correlation"
        Me.mnuFloorHoG.Visible = False
        '
        'mnuSPR
        '
        Me.mnuSPR.Name = "mnuSPR"
        Me.mnuSPR.Size = New System.Drawing.Size(272, 22)
        Me.mnuSPR.Text = "SPR"
        '
        'mnuGER
        '
        Me.mnuGER.Name = "mnuGER"
        Me.mnuGER.Size = New System.Drawing.Size(272, 22)
        Me.mnuGER.Text = "GER"
        '
        'mnuUnconstrainedHeat
        '
        Me.mnuUnconstrainedHeat.Name = "mnuUnconstrainedHeat"
        Me.mnuUnconstrainedHeat.Size = New System.Drawing.Size(272, 22)
        Me.mnuUnconstrainedHeat.Text = "Unconstrained HRR"
        '
        'mnuGlassTempGraph
        '
        Me.mnuGlassTempGraph.Name = "mnuGlassTempGraph"
        Me.mnuGlassTempGraph.Size = New System.Drawing.Size(272, 22)
        Me.mnuGlassTempGraph.Text = "Glass Temperatures"
        Me.mnuGlassTempGraph.Visible = False
        '
        'mnuNormalisedHeatLoad
        '
        Me.mnuNormalisedHeatLoad.Name = "mnuNormalisedHeatLoad"
        Me.mnuNormalisedHeatLoad.Size = New System.Drawing.Size(272, 22)
        Me.mnuNormalisedHeatLoad.Text = "Normalised Heat Load"
        Me.mnuNormalisedHeatLoad.Visible = False
        '
        'ToolStripMenuItem23
        '
        Me.ToolStripMenuItem23.Name = "ToolStripMenuItem23"
        Me.ToolStripMenuItem23.Size = New System.Drawing.Size(272, 22)
        Me.ToolStripMenuItem23.Text = "NHL Furnace"
        Me.ToolStripMenuItem23.Visible = False
        '
        'ToolStripMenuItem28
        '
        Me.ToolStripMenuItem28.Name = "ToolStripMenuItem28"
        Me.ToolStripMenuItem28.Size = New System.Drawing.Size(272, 22)
        Me.ToolStripMenuItem28.Text = "Furnace Surface Temp"
        Me.ToolStripMenuItem28.Visible = False
        '
        'ToolStripMenuItem29
        '
        Me.ToolStripMenuItem29.Name = "ToolStripMenuItem29"
        Me.ToolStripMenuItem29.Size = New System.Drawing.Size(272, 22)
        Me.ToolStripMenuItem29.Text = "Furnace Net Heat Flux"
        Me.ToolStripMenuItem29.Visible = False
        '
        'mnuPlotWoodBurningRate
        '
        Me.mnuPlotWoodBurningRate.Name = "mnuPlotWoodBurningRate"
        Me.mnuPlotWoodBurningRate.Size = New System.Drawing.Size(272, 22)
        Me.mnuPlotWoodBurningRate.Text = "Surface Wood Burning Rate"
        '
        'ToolStripMenuItem31
        '
        Me.ToolStripMenuItem31.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.burningrate, Me.areashrinkage})
        Me.ToolStripMenuItem31.Name = "ToolStripMenuItem31"
        Me.ToolStripMenuItem31.Size = New System.Drawing.Size(272, 22)
        Me.ToolStripMenuItem31.Text = "Fuel response effects"
        '
        'burningrate
        '
        Me.burningrate.Name = "burningrate"
        Me.burningrate.Size = New System.Drawing.Size(175, 22)
        Me.burningrate.Text = "Burning rate"
        '
        'areashrinkage
        '
        Me.areashrinkage.Name = "areashrinkage"
        Me.areashrinkage.Size = New System.Drawing.Size(175, 22)
        Me.areashrinkage.Text = "Fuel area shrinkage"
        '
        'mnuSmokeView
        '
        Me.mnuSmokeView.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ViewToolStripMenuItem, Me.CreateSmokeViewDataFileToolStripMenuItem})
        Me.mnuSmokeView.Image = Global.BRISK.My.Resources.Resources.GEOMTRY1
        Me.mnuSmokeView.Name = "mnuSmokeView"
        Me.mnuSmokeView.Size = New System.Drawing.Size(95, 20)
        Me.mnuSmokeView.Text = "Smokeview"
        '
        'ViewToolStripMenuItem
        '
        Me.ViewToolStripMenuItem.Name = "ViewToolStripMenuItem"
        Me.ViewToolStripMenuItem.Size = New System.Drawing.Size(220, 22)
        Me.ViewToolStripMenuItem.Text = "View"
        Me.ViewToolStripMenuItem.Visible = False
        '
        'CreateSmokeViewDataFileToolStripMenuItem
        '
        Me.CreateSmokeViewDataFileToolStripMenuItem.Name = "CreateSmokeViewDataFileToolStripMenuItem"
        Me.CreateSmokeViewDataFileToolStripMenuItem.Size = New System.Drawing.Size(220, 22)
        Me.CreateSmokeViewDataFileToolStripMenuItem.Text = "Create SmokeView Data File"
        Me.CreateSmokeViewDataFileToolStripMenuItem.Visible = False
        '
        'mnuHelp
        '
        Me.mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuHelpContents, Me.mnuTechRef, Me.mnuVerification, Me.mnuEULA, Me.mnuVersion, Me.mnuAnotherSep, Me.mnuAbout})
        Me.mnuHelp.Image = Global.BRISK.My.Resources.Resources.picHelpsymbol
        Me.mnuHelp.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(60, 20)
        Me.mnuHelp.Text = "&Help"
        '
        'mnuHelpContents
        '
        Me.mnuHelpContents.Enabled = False
        Me.mnuHelpContents.Name = "mnuHelpContents"
        Me.mnuHelpContents.ShortcutKeys = System.Windows.Forms.Keys.F1
        Me.mnuHelpContents.Size = New System.Drawing.Size(254, 22)
        Me.mnuHelpContents.Text = "Contents"
        Me.mnuHelpContents.Visible = False
        '
        'mnuTechRef
        '
        Me.mnuTechRef.Image = Global.BRISK.My.Resources.Resources.Reader16
        Me.mnuTechRef.Name = "mnuTechRef"
        Me.mnuTechRef.Size = New System.Drawing.Size(254, 22)
        Me.mnuTechRef.Text = "Technical Reference Manual (PDF)"
        '
        'mnuVerification
        '
        Me.mnuVerification.Enabled = False
        Me.mnuVerification.Image = Global.BRISK.My.Resources.Resources.Reader16
        Me.mnuVerification.Name = "mnuVerification"
        Me.mnuVerification.Size = New System.Drawing.Size(254, 22)
        Me.mnuVerification.Text = "Verification Data (PDF)"
        Me.mnuVerification.Visible = False
        '
        'mnuEULA
        '
        Me.mnuEULA.Name = "mnuEULA"
        Me.mnuEULA.Size = New System.Drawing.Size(254, 22)
        Me.mnuEULA.Text = "End User License Agreement"
        '
        'mnuVersion
        '
        Me.mnuVersion.Name = "mnuVersion"
        Me.mnuVersion.Size = New System.Drawing.Size(254, 22)
        Me.mnuVersion.Text = "Version Release Notes"
        '
        'mnuAnotherSep
        '
        Me.mnuAnotherSep.Name = "mnuAnotherSep"
        Me.mnuAnotherSep.Size = New System.Drawing.Size(251, 6)
        '
        'mnuAbout
        '
        Me.mnuAbout.Name = "mnuAbout"
        Me.mnuAbout.Size = New System.Drawing.Size(254, 22)
        Me.mnuAbout.Text = "&About B-RISK"
        '
        'mnuStop1
        '
        Me.mnuStop1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuTerminate, Me.mnuContinue, Me.mnuCancelBatch})
        Me.mnuStop1.Image = Global.BRISK.My.Resources.Resources.Abort
        Me.mnuStop1.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.mnuStop1.Name = "mnuStop1"
        Me.mnuStop1.Size = New System.Drawing.Size(59, 20)
        Me.mnuStop1.Text = "Stop"
        Me.mnuStop1.Visible = False
        '
        'mnuTerminate
        '
        Me.mnuTerminate.Name = "mnuTerminate"
        Me.mnuTerminate.Size = New System.Drawing.Size(207, 22)
        Me.mnuTerminate.Text = "Terminate the Simulation"
        '
        'mnuContinue
        '
        Me.mnuContinue.Name = "mnuContinue"
        Me.mnuContinue.Size = New System.Drawing.Size(207, 22)
        Me.mnuContinue.Text = "Continue Simulation"
        '
        'mnuCancelBatch
        '
        Me.mnuCancelBatch.Name = "mnuCancelBatch"
        Me.mnuCancelBatch.Size = New System.Drawing.Size(207, 22)
        Me.mnuCancelBatch.Text = "Cancel Batch Files"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewToolStripMenuItem, Me.OpenToolStripMenuItem, Me.toolStripSeparator, Me.SaveToolStripMenuItem, Me.SaveAsToolStripMenuItem, Me.toolStripSeparator1, Me.PrintToolStripMenuItem, Me.PrintPreviewToolStripMenuItem, Me.toolStripSeparator3, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        Me.FileToolStripMenuItem.Visible = False
        '
        'NewToolStripMenuItem
        '
        Me.NewToolStripMenuItem.Image = CType(resources.GetObject("NewToolStripMenuItem.Image"), System.Drawing.Image)
        Me.NewToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.NewToolStripMenuItem.Name = "NewToolStripMenuItem"
        Me.NewToolStripMenuItem.Size = New System.Drawing.Size(143, 22)
        Me.NewToolStripMenuItem.Text = "&New"
        '
        'OpenToolStripMenuItem
        '
        Me.OpenToolStripMenuItem.Image = CType(resources.GetObject("OpenToolStripMenuItem.Image"), System.Drawing.Image)
        Me.OpenToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.OpenToolStripMenuItem.Name = "OpenToolStripMenuItem"
        Me.OpenToolStripMenuItem.Size = New System.Drawing.Size(143, 22)
        Me.OpenToolStripMenuItem.Text = "&Open"
        '
        'toolStripSeparator
        '
        Me.toolStripSeparator.Name = "toolStripSeparator"
        Me.toolStripSeparator.Size = New System.Drawing.Size(140, 6)
        '
        'SaveToolStripMenuItem
        '
        Me.SaveToolStripMenuItem.Image = CType(resources.GetObject("SaveToolStripMenuItem.Image"), System.Drawing.Image)
        Me.SaveToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.SaveToolStripMenuItem.Name = "SaveToolStripMenuItem"
        Me.SaveToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
        Me.SaveToolStripMenuItem.Size = New System.Drawing.Size(143, 22)
        Me.SaveToolStripMenuItem.Text = "&Save"
        '
        'SaveAsToolStripMenuItem
        '
        Me.SaveAsToolStripMenuItem.Name = "SaveAsToolStripMenuItem"
        Me.SaveAsToolStripMenuItem.Size = New System.Drawing.Size(143, 22)
        Me.SaveAsToolStripMenuItem.Text = "Save &As"
        '
        'toolStripSeparator1
        '
        Me.toolStripSeparator1.Name = "toolStripSeparator1"
        Me.toolStripSeparator1.Size = New System.Drawing.Size(140, 6)
        '
        'PrintToolStripMenuItem
        '
        Me.PrintToolStripMenuItem.Image = CType(resources.GetObject("PrintToolStripMenuItem.Image"), System.Drawing.Image)
        Me.PrintToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.PrintToolStripMenuItem.Name = "PrintToolStripMenuItem"
        Me.PrintToolStripMenuItem.Size = New System.Drawing.Size(143, 22)
        Me.PrintToolStripMenuItem.Text = "&Print"
        '
        'PrintPreviewToolStripMenuItem
        '
        Me.PrintPreviewToolStripMenuItem.Image = CType(resources.GetObject("PrintPreviewToolStripMenuItem.Image"), System.Drawing.Image)
        Me.PrintPreviewToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.PrintPreviewToolStripMenuItem.Name = "PrintPreviewToolStripMenuItem"
        Me.PrintPreviewToolStripMenuItem.Size = New System.Drawing.Size(143, 22)
        Me.PrintPreviewToolStripMenuItem.Text = "Print Pre&view"
        '
        'toolStripSeparator3
        '
        Me.toolStripSeparator3.Name = "toolStripSeparator3"
        Me.toolStripSeparator3.Size = New System.Drawing.Size(140, 6)
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(143, 22)
        Me.ExitToolStripMenuItem.Text = "E&xit"
        '
        'EditToolStripMenuItem
        '
        Me.EditToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UndoToolStripMenuItem, Me.RedoToolStripMenuItem, Me.toolStripSeparator4, Me.CutToolStripMenuItem, Me.CopyToolStripMenuItem, Me.PasteToolStripMenuItem, Me.toolStripSeparator5, Me.SelectAllToolStripMenuItem})
        Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
        Me.EditToolStripMenuItem.Size = New System.Drawing.Size(39, 20)
        Me.EditToolStripMenuItem.Text = "&Edit"
        Me.EditToolStripMenuItem.Visible = False
        '
        'UndoToolStripMenuItem
        '
        Me.UndoToolStripMenuItem.Name = "UndoToolStripMenuItem"
        Me.UndoToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Z), System.Windows.Forms.Keys)
        Me.UndoToolStripMenuItem.Size = New System.Drawing.Size(144, 22)
        Me.UndoToolStripMenuItem.Text = "&Undo"
        '
        'RedoToolStripMenuItem
        '
        Me.RedoToolStripMenuItem.Name = "RedoToolStripMenuItem"
        Me.RedoToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Y), System.Windows.Forms.Keys)
        Me.RedoToolStripMenuItem.Size = New System.Drawing.Size(144, 22)
        Me.RedoToolStripMenuItem.Text = "&Redo"
        '
        'toolStripSeparator4
        '
        Me.toolStripSeparator4.Name = "toolStripSeparator4"
        Me.toolStripSeparator4.Size = New System.Drawing.Size(141, 6)
        '
        'CutToolStripMenuItem
        '
        Me.CutToolStripMenuItem.Image = CType(resources.GetObject("CutToolStripMenuItem.Image"), System.Drawing.Image)
        Me.CutToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.CutToolStripMenuItem.Name = "CutToolStripMenuItem"
        Me.CutToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.X), System.Windows.Forms.Keys)
        Me.CutToolStripMenuItem.Size = New System.Drawing.Size(144, 22)
        Me.CutToolStripMenuItem.Text = "Cu&t"
        '
        'CopyToolStripMenuItem
        '
        Me.CopyToolStripMenuItem.Image = CType(resources.GetObject("CopyToolStripMenuItem.Image"), System.Drawing.Image)
        Me.CopyToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.CopyToolStripMenuItem.Name = "CopyToolStripMenuItem"
        Me.CopyToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.CopyToolStripMenuItem.Size = New System.Drawing.Size(144, 22)
        Me.CopyToolStripMenuItem.Text = "&Copy"
        '
        'PasteToolStripMenuItem
        '
        Me.PasteToolStripMenuItem.Image = CType(resources.GetObject("PasteToolStripMenuItem.Image"), System.Drawing.Image)
        Me.PasteToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.PasteToolStripMenuItem.Name = "PasteToolStripMenuItem"
        Me.PasteToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.V), System.Windows.Forms.Keys)
        Me.PasteToolStripMenuItem.Size = New System.Drawing.Size(144, 22)
        Me.PasteToolStripMenuItem.Text = "&Paste"
        '
        'toolStripSeparator5
        '
        Me.toolStripSeparator5.Name = "toolStripSeparator5"
        Me.toolStripSeparator5.Size = New System.Drawing.Size(141, 6)
        '
        'SelectAllToolStripMenuItem
        '
        Me.SelectAllToolStripMenuItem.Name = "SelectAllToolStripMenuItem"
        Me.SelectAllToolStripMenuItem.Size = New System.Drawing.Size(144, 22)
        Me.SelectAllToolStripMenuItem.Text = "Select &All"
        '
        'ToolsToolStripMenuItem
        '
        Me.ToolsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CustomizeToolStripMenuItem, Me.OptionsToolStripMenuItem})
        Me.ToolsToolStripMenuItem.Name = "ToolsToolStripMenuItem"
        Me.ToolsToolStripMenuItem.Size = New System.Drawing.Size(47, 20)
        Me.ToolsToolStripMenuItem.Text = "&Tools"
        Me.ToolsToolStripMenuItem.Visible = False
        '
        'CustomizeToolStripMenuItem
        '
        Me.CustomizeToolStripMenuItem.Name = "CustomizeToolStripMenuItem"
        Me.CustomizeToolStripMenuItem.Size = New System.Drawing.Size(130, 22)
        Me.CustomizeToolStripMenuItem.Text = "&Customize"
        '
        'OptionsToolStripMenuItem
        '
        Me.OptionsToolStripMenuItem.Name = "OptionsToolStripMenuItem"
        Me.OptionsToolStripMenuItem.Size = New System.Drawing.Size(130, 22)
        Me.OptionsToolStripMenuItem.Text = "&Options"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ContentsToolStripMenuItem, Me.IndexToolStripMenuItem, Me.SearchToolStripMenuItem, Me.toolStripSeparator6, Me.AboutToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "&Help"
        Me.HelpToolStripMenuItem.Visible = False
        '
        'ContentsToolStripMenuItem
        '
        Me.ContentsToolStripMenuItem.Name = "ContentsToolStripMenuItem"
        Me.ContentsToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.ContentsToolStripMenuItem.Text = "&Contents"
        '
        'IndexToolStripMenuItem
        '
        Me.IndexToolStripMenuItem.Name = "IndexToolStripMenuItem"
        Me.IndexToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.IndexToolStripMenuItem.Text = "&Index"
        '
        'SearchToolStripMenuItem
        '
        Me.SearchToolStripMenuItem.Name = "SearchToolStripMenuItem"
        Me.SearchToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.SearchToolStripMenuItem.Text = "&Search"
        '
        'toolStripSeparator6
        '
        Me.toolStripSeparator6.Name = "toolStripSeparator6"
        Me.toolStripSeparator6.Size = New System.Drawing.Size(119, 6)
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.AboutToolStripMenuItem.Text = "&About..."
        '
        'CeilingToolStripMenuItem1
        '
        Me.CeilingToolStripMenuItem1.Name = "CeilingToolStripMenuItem1"
        Me.CeilingToolStripMenuItem1.Size = New System.Drawing.Size(152, 22)
        Me.CeilingToolStripMenuItem1.Text = "Ceiling"
        '
        'AverageToolStripMenuItem
        '
        Me.AverageToolStripMenuItem.Name = "AverageToolStripMenuItem"
        Me.AverageToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.AverageToolStripMenuItem.Text = "Average"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'FolderBrowserDialog1
        '
        Me.FolderBrowserDialog1.RootFolder = System.Environment.SpecialFolder.ApplicationData
        '
        'ToolStrip1
        '
        Me.ToolStrip1.BackColor = System.Drawing.SystemColors.MenuBar
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewToolStripButton, Me.OpenToolStripButton, Me.SaveToolStripButton, Me.PrintToolStripButton, Me.ToolStripButton5, Me.toolStripSeparator2, Me.mnuRun, Me.mnuStop})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 24)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1249, 25)
        Me.ToolStrip1.Stretch = True
        Me.ToolStrip1.TabIndex = 23
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'NewToolStripButton
        '
        Me.NewToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.NewToolStripButton.Image = CType(resources.GetObject("NewToolStripButton.Image"), System.Drawing.Image)
        Me.NewToolStripButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.NewToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.NewToolStripButton.Name = "NewToolStripButton"
        Me.NewToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.NewToolStripButton.Text = "&New"
        Me.NewToolStripButton.ToolTipText = "New Base Model"
        '
        'OpenToolStripButton
        '
        Me.OpenToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.OpenToolStripButton.Image = CType(resources.GetObject("OpenToolStripButton.Image"), System.Drawing.Image)
        Me.OpenToolStripButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.OpenToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.OpenToolStripButton.Name = "OpenToolStripButton"
        Me.OpenToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.OpenToolStripButton.Text = "&Open"
        Me.OpenToolStripButton.ToolTipText = "Open Base Model"
        '
        'SaveToolStripButton
        '
        Me.SaveToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.SaveToolStripButton.Image = CType(resources.GetObject("SaveToolStripButton.Image"), System.Drawing.Image)
        Me.SaveToolStripButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.SaveToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.SaveToolStripButton.Name = "SaveToolStripButton"
        Me.SaveToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.SaveToolStripButton.Text = "&Save"
        Me.SaveToolStripButton.ToolTipText = "Save Base Model"
        '
        'PrintToolStripButton
        '
        Me.PrintToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.PrintToolStripButton.Image = CType(resources.GetObject("PrintToolStripButton.Image"), System.Drawing.Image)
        Me.PrintToolStripButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.PrintToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.PrintToolStripButton.Name = "PrintToolStripButton"
        Me.PrintToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.PrintToolStripButton.Text = "&Print"
        Me.PrintToolStripButton.ToolTipText = "Print Input & Results"
        '
        'ToolStripButton5
        '
        Me.ToolStripButton5.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton5.Image = Global.BRISK.My.Resources.Resources.picExcel16
        Me.ToolStripButton5.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.ToolStripButton5.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton5.Name = "ToolStripButton5"
        Me.ToolStripButton5.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton5.Text = "ToolStripButton5"
        Me.ToolStripButton5.ToolTipText = "Send Output to Excel File"
        '
        'toolStripSeparator2
        '
        Me.toolStripSeparator2.Name = "toolStripSeparator2"
        Me.toolStripSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'mnuRun
        '
        Me.mnuRun.Image = Global.BRISK.My.Resources.Resources.Go
        Me.mnuRun.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.mnuRun.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.mnuRun.Name = "mnuRun"
        Me.mnuRun.Size = New System.Drawing.Size(48, 22)
        Me.mnuRun.Text = "&Run"
        Me.mnuRun.ToolTipText = "Start Simulation"
        Me.mnuRun.Visible = False
        '
        'mnuStop
        '
        Me.mnuStop.Image = Global.BRISK.My.Resources.Resources.picStop
        Me.mnuStop.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.mnuStop.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.mnuStop.Name = "mnuStop"
        Me.mnuStop.Size = New System.Drawing.Size(51, 22)
        Me.mnuStop.Text = "Stop"
        Me.mnuStop.ToolTipText = "Stop Simulation"
        Me.mnuStop.Visible = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripStatusLabel2, Me.ToolStripStatusLabel3, Me.ToolStripStatusLabel4, Me.ToolStripProgressBar1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 580)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1249, 22)
        Me.StatusStrip1.TabIndex = 25
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.BackColor = System.Drawing.SystemColors.MenuBar
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(120, 17)
        Me.ToolStripStatusLabel1.Text = "ToolStripStatusLabel1"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.BackColor = System.Drawing.SystemColors.MenuBar
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(120, 17)
        Me.ToolStripStatusLabel2.Text = "ToolStripStatusLabel2"
        '
        'ToolStripStatusLabel3
        '
        Me.ToolStripStatusLabel3.BackColor = System.Drawing.SystemColors.MenuBar
        Me.ToolStripStatusLabel3.LinkBehavior = System.Windows.Forms.LinkBehavior.AlwaysUnderline
        Me.ToolStripStatusLabel3.Name = "ToolStripStatusLabel3"
        Me.ToolStripStatusLabel3.Size = New System.Drawing.Size(16, 17)
        Me.ToolStripStatusLabel3.Text = "   "
        '
        'ToolStripStatusLabel4
        '
        Me.ToolStripStatusLabel4.BackColor = System.Drawing.SystemColors.MenuBar
        Me.ToolStripStatusLabel4.ForeColor = System.Drawing.SystemColors.InfoText
        Me.ToolStripStatusLabel4.Name = "ToolStripStatusLabel4"
        Me.ToolStripStatusLabel4.Size = New System.Drawing.Size(978, 17)
        Me.ToolStripStatusLabel4.Spring = True
        Me.ToolStripStatusLabel4.Text = "   "
        Me.ToolStripStatusLabel4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ToolStripProgressBar1
        '
        Me.ToolStripProgressBar1.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.ToolStripProgressBar1.Name = "ToolStripProgressBar1"
        Me.ToolStripProgressBar1.Size = New System.Drawing.Size(200, 16)
        Me.ToolStripProgressBar1.Visible = False
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 500
        '
        'Process1
        '
        Me.Process1.StartInfo.Domain = ""
        Me.Process1.StartInfo.LoadUserProfile = False
        Me.Process1.StartInfo.Password = Nothing
        Me.Process1.StartInfo.StandardErrorEncoding = Nothing
        Me.Process1.StartInfo.StandardOutputEncoding = Nothing
        Me.Process1.StartInfo.UserName = ""
        Me.Process1.SynchronizingObject = Me
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.ChartRuntime4, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.ChartRuntime2, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.ChartRuntime3, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.ChartRuntime1, 0, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(26, 80)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(897, 468)
        Me.TableLayoutPanel1.TabIndex = 29
        Me.TableLayoutPanel1.Visible = False
        '
        'ChartRuntime4
        '
        ChartArea5.AxisX.Title = "Time (sec)"
        ChartArea5.AxisY.Title = "Heat Release Rate (kW)"
        ChartArea5.Name = "ChartArea1"
        Me.ChartRuntime4.ChartAreas.Add(ChartArea5)
        Me.ChartRuntime4.Dock = System.Windows.Forms.DockStyle.Fill
        Legend5.Name = "Legend1"
        Me.ChartRuntime4.Legends.Add(Legend5)
        Me.ChartRuntime4.Location = New System.Drawing.Point(451, 237)
        Me.ChartRuntime4.Name = "ChartRuntime4"
        Series5.ChartArea = "ChartArea1"
        Series5.Legend = "Legend1"
        Series5.Name = "Series1"
        Me.ChartRuntime4.Series.Add(Series5)
        Me.ChartRuntime4.Size = New System.Drawing.Size(443, 228)
        Me.ChartRuntime4.TabIndex = 38
        Me.ChartRuntime4.Visible = False
        '
        'ChartRuntime2
        '
        ChartArea6.AxisX.Title = "Time (sec)"
        ChartArea6.AxisY.Title = "Heat Release Rate (kW)"
        ChartArea6.Name = "ChartArea1"
        Me.ChartRuntime2.ChartAreas.Add(ChartArea6)
        Me.ChartRuntime2.Dock = System.Windows.Forms.DockStyle.Fill
        Legend6.Name = "Legend1"
        Me.ChartRuntime2.Legends.Add(Legend6)
        Me.ChartRuntime2.Location = New System.Drawing.Point(451, 3)
        Me.ChartRuntime2.Name = "ChartRuntime2"
        Series6.ChartArea = "ChartArea1"
        Series6.Legend = "Legend1"
        Series6.Name = "Series1"
        Me.ChartRuntime2.Series.Add(Series6)
        Me.ChartRuntime2.Size = New System.Drawing.Size(443, 228)
        Me.ChartRuntime2.TabIndex = 37
        Me.ChartRuntime2.Visible = False
        '
        'ChartRuntime3
        '
        ChartArea7.AxisX.Title = "Time (sec)"
        ChartArea7.AxisY.Title = "Heat Release Rate (kW)"
        ChartArea7.Name = "ChartArea1"
        Me.ChartRuntime3.ChartAreas.Add(ChartArea7)
        Me.ChartRuntime3.Dock = System.Windows.Forms.DockStyle.Fill
        Legend7.Name = "Legend1"
        Me.ChartRuntime3.Legends.Add(Legend7)
        Me.ChartRuntime3.Location = New System.Drawing.Point(3, 237)
        Me.ChartRuntime3.Name = "ChartRuntime3"
        Series7.ChartArea = "ChartArea1"
        Series7.Legend = "Legend1"
        Series7.Name = "Series1"
        Me.ChartRuntime3.Series.Add(Series7)
        Me.ChartRuntime3.Size = New System.Drawing.Size(442, 228)
        Me.ChartRuntime3.TabIndex = 35
        Me.ChartRuntime3.Visible = False
        '
        'ChartRuntime1
        '
        ChartArea8.AxisX.Title = "Time (sec)"
        ChartArea8.AxisY.Title = "Heat Release Rate (kW)"
        ChartArea8.Name = "ChartArea1"
        Me.ChartRuntime1.ChartAreas.Add(ChartArea8)
        Me.ChartRuntime1.Dock = System.Windows.Forms.DockStyle.Fill
        Legend8.Name = "Legend1"
        Me.ChartRuntime1.Legends.Add(Legend8)
        Me.ChartRuntime1.Location = New System.Drawing.Point(3, 3)
        Me.ChartRuntime1.Name = "ChartRuntime1"
        Series8.ChartArea = "ChartArea1"
        Series8.Legend = "Legend1"
        Series8.Name = "Series1"
        Me.ChartRuntime1.Series.Add(Series8)
        Me.ChartRuntime1.Size = New System.Drawing.Size(442, 228)
        Me.ChartRuntime1.TabIndex = 36
        Me.ChartRuntime1.Visible = False
        '
        'MDIFrmMain
        '
        Me.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.ClientSize = New System.Drawing.Size(1249, 602)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.MainMenu1)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.Location = New System.Drawing.Point(11, 57)
        Me.Name = "MDIFrmMain"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "B-RISK (DESIGN FIRE TOOL)"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.TableLayoutPanel1.ResumeLayout(False)
        CType(Me.ChartRuntime4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ChartRuntime2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ChartRuntime3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ChartRuntime1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents PrintDocument1 As System.Drawing.Printing.PrintDocument
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents mnuInputs As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents NewToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents OpenToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents SaveToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents PrintToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel3 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel4 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents Process1 As System.Diagnostics.Process
    Friend WithEvents UserModeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NZBCVM2ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents ChartRuntime3 As System.Windows.Forms.DataVisualization.Charting.Chart
    Friend WithEvents ChartRuntime4 As System.Windows.Forms.DataVisualization.Charting.Chart
    Friend WithEvents ChartRuntime2 As System.Windows.Forms.DataVisualization.Charting.Chart
    Friend WithEvents ChartRuntime1 As System.Windows.Forms.DataVisualization.Charting.Chart
    Friend WithEvents RiskSimulatorToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents toolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripButton5 As System.Windows.Forms.ToolStripButton
    Friend WithEvents TempToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LoadIterationToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OpenBaseModelToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents VentsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents WallVentsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CeilingVentsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MechanicalFansToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuDetection As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SprinklersHeatDetectorsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SmokeDetectorsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuRun As System.Windows.Forms.ToolStripButton
    Friend WithEvents mnuStop As System.Windows.Forms.ToolStripButton
    Friend WithEvents mnuSmokeView As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ViewToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CreateSmokeViewDataFileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuTenability As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TenabilityParametersToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FEDEgressPathToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SelectFireToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AdditionalCombustionParametersToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents COSootToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FireObjectDatabaseToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SurfaceMaterialsDatabaseToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFlameSpread As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SetingsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MaterialConeFileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SolversToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AmbientConditionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CompartmentEffectsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PostflashoverBehaviourToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ModelPhysicsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SaveBaseToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SmokeDetectorsnewToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UtilitiesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OpenBRANZFIREModFileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SaveBaseModelAsAsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCeilingJet As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem3 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem4 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem5 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ChangeFireDatabaseFileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NewToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OpenToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents toolStripSeparator As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents SaveToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SaveAsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents toolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents PrintToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PrintPreviewToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents toolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EditToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UndoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RedoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents toolStripSeparator4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents CutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CopyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PasteToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents toolStripSeparator5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents SelectAllToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CustomizeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OptionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ContentsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents IndexToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SearchToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents toolStripSeparator6 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RoomsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuUnconstrainedHeat As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CeilingVentsNewToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuNormalisedHeatLoad As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSurfaceFluxes As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSurfaceTemperatures As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSetRiskdataFolder As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents IncidentHeatFluxesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NetHeatFluxesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem6 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem7 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem8 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem9 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem10 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem11 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem12 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem13 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAST As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem14 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem15 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem16 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem17 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem18 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem19 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem20 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem21 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem22 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuConvectHTC As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem24 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem25 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem26 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem27 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem23 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CeilingToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AverageToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem28 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem29 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ChangeMaterialsDatabaseFilethermalmdbToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DevelopmentKeyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem30 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TalkToEVACNZToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EvacuatioNZSettingsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents uwallchardepth As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ceilingchardepth As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BoundaryNodeTemperaturesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UpperWallToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CeilingToolStripMenuItem2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuPlotWoodBurningRate As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem31 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents burningrate As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents areashrinkage As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BatchFilesToolStripMenuItem As ToolStripMenuItem
#End Region
End Class