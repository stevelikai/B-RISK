Option Strict Off
Option Explicit On

'Imports System.Globalization
'Imports System.Security.Permissions
'Imports System.Threading

'<Assembly: SecurityPermission(SecurityAction.RequestMinimum, ControlThread:=True)> 

Module GLOBAL_Renamed
    '============================================================================'
    'A PROGRAM FOR MODELLING SMOKE DEVELOPMENT AND FIRE GROWTH IN A ROOM         '
    'By Colleen Wade'                               '
    '============================================================================'
    '                                                                            '
    ' Visual Basic global constant file.                                         '
    '
    'edits:
    ''31 January 2008 CW
    '============================================================================'
    Public Const DefaultAbsTolerance As Double = 0.000001
    Public FuelResponseEffects As Boolean = False
    Public autosavepdf As Boolean = False
    Public autosaveXL As Boolean = False
    Public FlameExtinctionModel As Boolean = True
    Public RCNone As Boolean

    'Global variables used in the program
    Public Const Version As String = "2018.051"
    Public Const DevKeyCode As String = "jam4sam16"
    Public DevKey As Boolean = False
    Public TalkToEVACNZ As Boolean
    Public OutputInterval As Integer
    Public NumberIterations As Integer
    Public IterationsCompleted As Integer
    Public enzsimtime As Double
    Public brisksimtime As Double
    Public roomcolor(0 To 11) As System.Drawing.Color
    Public wallpercent As Integer
    Public ceilingpercent As Integer
    Public wallHFSlimit As Double
    Public wallVFSlimit As Double
    Public TEeurocode As Single
    Public ExcessFuelFactor As Double
    Public HRRatFO As Double
    Public burnmode As Boolean 'used in wood crib PF model
    Public NHLmax(0 To 4) As Double
    Public NHLte(,) As Double
    Public upgrade As Boolean
    Public ventid As Integer

    Public CLT_instant As Boolean = False
    Public wall_char(,) As Double
    Public ceil_char(,) As Double
    Public CLTwallpercent As Integer
    Public CLTceilingpercent As Integer
    Public Lamella As Single
    Public IntegralModel As Boolean
    Public KineticModel As Boolean
    Public CLTQcrit As Double
    Public CLTflameflux As Double
    Public CLTLoG As Double
    Public CLTigtemp As Double
    Public CLTcalibrationfactor As Double
    Public Lamella1 As Single
    Public Lamella2 As Single
    Public DebondTemp As Double
    Public mcNumberIterations As Integer = 1 'number of monte carlo iterations
    Public furnaceNHL(,) As Double
    Public furnaceST(,) As Double
    Public furnaceqnet(,) As Double
    Public CeilingNode(,,) As Double
    Public UWallNode(,,) As Double
    Public LWallNode(,,) As Double
    Public FloorNode(,,) As Double
    Public elementcounter As Integer
    Public CeilingElementMF(,,) As Double
    Public UWallElementMF(,,) As Double
    Public LWallElementMF(,,) As Double
    Public CeilingCharResidue(,) As Double
    Public CeilingResidualMass(,) As Double
    Public CeilingWoodMLR(,) As Double
    Public CeilingWoodMLR_tot() As Double
    Public UWallCharResidue(,) As Double
    Public LWallCharResidue(,) As Double
    Public item_location(,) As Single 'details of items in fire room
    Public vectorlength(,) As Single 'length of midpoint vectors from source to target items, ignition time in 2nd dimension
    Public targetdistance(,) As Single 'length of the vector from centre of a source to the nearest part of a target object
    Public targetxpoint(,) As Single 'x coordinate of the nearest part of a target object
    Public targetypoint(,) As Single 'y coordinate of the nearest part of a target object
    Public itemtowalldistance() As Single 'the least distance between an object and a wall surface 
    Public n_max As Integer 'number of fuel objects
    Public itemcount() As Integer 'number of items in each iteration
    Public FlagSimStop As Boolean = True 'flag to indicate if mc simulation underway
    Public NumItems As Integer
    Public NumVents As Integer
    Public NumCVents As Integer
    Public VM2 As Boolean
    Public sprink_mode As Integer
    Public sd_mode As Boolean
    Public mv_mode As Boolean 'mech vent
    Public firstitem As Integer
    Public firstitemtemp As Integer
    Public itime() As Single
    Public sprmode_store() As Integer
    Public alphaTfitted(,) As Single
    Public timeequiv(,) As Double
    Public ignmode() As String 'store ignition mode of secondary items
    Public autopopulate As Boolean = True
    Public calc_sprdist As Boolean = False
    Public calc_sddist As Boolean = False
    Public ignitetargets As Boolean = False
    Public terminate_fo As Boolean = False
    Public terminate_fuelgone As Boolean = False
    Public modGER As Boolean = False
    Public usepowerlawdesignfire As Boolean = False
    Public useT2fire As Boolean = True
    Public VentilationLimitFlag As Boolean = False
    Public flashover_time As Double
    Public FEDCO As Boolean
    Public FEDPath(,) As Single
    Public fandata(,) As Single
    Public energy(,) As Single
    Public ModelVersion As Single
    Public mc_go As Boolean
    Public boo_startup As Boolean = True
    Public FLED_Actual As Single
    Public mc_ULTemp(,,) As Double
    Public mc_LayerHeight(,,) As Double
    Public mc_visi(,,) As Double
    Public mc_FEDgas(,,) As Double
    Public mc_FEDheat(,,) As Double
    Public mc_fanflowrate(,) As Double
    Public mc_fanIDreliability(,) As Double
    Public mc_fanstarttime(,) As Double
    Public mc_fanpressurelimit(,) As Double
    Public mc_SDRadialDist(,) As Double
    Public mc_SDOD(,) As Double
    Public mc_SDZ(,) As Double
    Public mc_SDcharlength(,) As Double
    Public mc_RadialDist(,) As Double
    Public mc_RTI(,) As Double
    Public mc_cfactor(,) As Double
    Public mc_acttemp(,) As Double
    Public mc_dist(,) As Double
    Public mc_waterdensity(,) As Double
    Public mc_ventprob(,) As Double
    Public mc_HOreliability(,) As Double
    Public mc_vent_area(,) As Double
    Public mc_vent_height(,) As Double
    Public mc_vent_width(,) As Double
    Public mc_room_width(,) As Double
    Public mc_room_length(,) As Double
    Public mc_integrity(,) As Double
    Public mc_maxopening(,) As Double
    Public mc_maxopeningtime(,) As Double
    Public mc_gastemp(,) As Double
    Public mc_integrity2(,) As Double
    Public mc_maxopening2(,) As Double
    Public mc_maxopeningtime2(,) As Double
    Public mc_gastemp2(,) As Double
    Public mc_item_hoc(,) As Double
    Public mc_item_soot(,) As Double
    Public mc_item_co2(,) As Double
    Public mc_item_lhog(,) As Double
    Public mc_item_RLF(,) As Double
    Public mc_item_hrrua(,) As Double
    Public AutoOpenVent(,,) As Boolean 'flag for when we use opening doors option
    Public AutoOpenCVent(,,) As Boolean 'flag for when we use opening doors option
    Public HRR_threshold(,,) As Double 'fire size for human detection
    Public HRR_threshold2(,,) As Double 'fire size for human detection
    Public HRR_threshold_time(,,) As Double 'time fire size threshold is reached
    Public HRR_threshold_time2(,,) As Double 'time fire size threshold is reached
    Public HRRFlag(,,) As Integer
    Public HRRFlag2(,,) As Integer
    Public trigger_ventopendelay(,,) As Double 'time delay following smoke detection activation
    Public trigger_ventopendelay2(,,) As Double 'time delay following smoke detection activation
    Public trigger_ventopenduration(,,) As Double 'time the vent remains open
    Public trigger_ventopenduration2(,,) As Double 'time the vent remains open
    Public HRR_ventopendelay(,,) As Double 'time delay following smoke detection activation
    Public HRR_ventopendelay2(,,) As Double 'time delay following smoke detection activation
    Public HRR_ventopenduration(,,) As Double 'time the vent remains open
    Public HRR_ventopenduration2(,,) As Double 'time the vent remains open
    Public SDtriggerroom(,,) As Integer 'the room to be used for the SD trigger
    Public SDtriggerroom2(,,) As Integer 'the room to be used for the SD trigger
    Public trigger_device(,,,) As Boolean
    Public trigger_device2(,,,) As Boolean
    Public oldVentCloseTime() As Single
    Public VentBreakTimeClose(,,) As Double
    Public CVentBreakTimeClose(,,) As Double
    Public AlphaT As Single
    Public PeakHRR As Single
    Public originalpeakhrr As Single
    Public gTimeX As Double
    Public NumSprinklers As Integer
    Public NumSmokeDetectors As Integer
    Public NumFans As Integer
    Public itcounter As Integer = 1
    Public Const vbModal As Short = 1
    Public Const minSDdepth As Double = 0 'the smoke depth must be greater than this depth for the SD to activate, consider effect both in room of origin and outside room of origin before changing
    Public Const ProgramTitle As String = "B-RISK DESIGN FIRE TOOL"
    Public Const gcs_folder_ext As String = "\B-RISK"

    'Public Const TrialPeriod As Short = 90
    Public oldbasefile As String = ""
    Public basefile As String = ""
    Public RiskDataDirectory As String
    Public DefaultRiskDataDirectory As String
    Public MasterDirectory As String
    Public ApplicationPath As String
    Public UserAppDataFolder As String
    Public CommonFilesFolder As String
    Public DocsFolder As String
    Public DataFolder As String
    Public UserPersonalDataFolder As String
    Public BatchFileFolder As String
    Public CVentlogfile As String
    Public WVentlogfile As String
    Public WVentlogfile2 As String
    Public DataDirectory As String
    Public DBDirectory As String
    Public ProjectDirectory As String

    Public FireDatabaseName As String
    Public MaterialsDatabaseName As String
    Public nowallflow As Boolean
    Public ventlog As Boolean

    Public Const DefaultFireDatabaseName As String = "\dbases\fire.mdb"
    Public Const DefaultMaterialsDatabaseName As String = "\dbases\thermal.mdb"
    Public Const techref_file As String = "techguide.pdf"
    Public Const userguide_file As String = "\docs\userguide.pdf"
    Public Const verification_file As String = "verification.pdf"
    Public Const gcs_UOC_XML As String = "http://www.civil.canterbury.ac.nz/spearpoint/HRR_Database/HRR_Database.xml"
    Public Const gcs_UOC_XSLT As String = "http://www.civil.canterbury.ac.nz/spearpoint/HRR_Database/hrrt_branzfire.xslt"
    Public distbackcolor As System.Drawing.Color = Color.LightSalmon
    Public distnobackcolor As System.Drawing.Color = Color.White
    Public Const flashover_crit_flux = 20 'kW/m2
    Public Const flashover_crit_temp = 773 'K
    Public FOFluxCriteria As Boolean
    Public Const CO2YieldPF As Double = 1.19
    Public Const H2OYieldPF As Double = 0.578
    Public Const SootYieldPF As Double = 0.015
    Public Const gcd_Machine_Error As Double = 0.000001 'machine error or decimal place accuracy
    Public Const Smallvalue As Double = 0.0000001 'to avoid numerical problems with the species concentrations
    Public Const cjJET As Short = 0
    Public Const cjAlpert As Short = 1
    Public Const vbJanssens As Short = 0
    Public Const vbFTP As Short = 1
    Public Const DefaultPrinterFontName As String = "Courier New"
    Public Const DefaultPrinterFontSize As String = "8"

    Public Const ParticleExtinction As Single = 8790 'm2/kg
    Public Const deltaHair As Double = 3 'kJ/g
    Public Const OutsideConvCoeff As Double = 5 'was 5 W/m2K Convective heat transfer outside room
    Public Const StefanBoltzmann As Double = 0.0000000566961 'W/m2K4 Stefan Boltzmann constant
    Public Const PI As Double = 3.14159265 'Pi
    Public Const Gas_Constant As Double = 8.3145 'kJ/kmol K universal gas constant
    Public Const Gas_Constant_Air As Double = 0.2871 'kJ/kgK
    Public Const Flametemp As Double = 1250 'K
    Public Const G As Double = 9.807 'm/s2  gravitational constant
    Public Const EntrainmentCoefficient As Double = 0.076 'plume entrainment coefficient
    Public Const ReferenceDensity As Double = 1.225 ' kg/m3  density of air at sea level
    Public Const ReferenceTemp As Double = 288.15 'temp corresponding to density of air at sea level
    Public Const O2Ambient As Double = 0.2313 'mass fraction of oxygen in air
    Public Const CO2Ambient As Double = 0.0005 'mass fractionv of CO2 in air
    Public Const SpecificHeat_fuel As Double = 1.7 'propane
    Public Const MolecularWeightAir As Double = 29 'molecular weight of air
    Public Const MolecularWeightCO2 As Double = 44 'molecular weight of CO2
    Public Const MolecularWeightCO As Double = 28 'molecular weight of CO
    Public Const MolecularWeightH2O As Double = 18 'molecular weight of H2O
    Public Const MolecularWeightHCN As Double = 27 'molecular weight of HCN
    Public Const MolecularWeightO2 As Double = 32 'molecular weight of O2
    Public Const MolecularWeightN2 As Double = 28 'molecular weight of O2
    Public Const MolecularWeightC As Double = 12 'molecular weight of O2

    Public Const URL1 As String = "https://link.springer.com/article/10.1007/s10694-009-0081-0" 'Fire Technology - Investigation of Sprinkler Sprays on Fire Induced Doorway Flows
    Public Const URL2 As String = "https://link.springer.com/article/10.1007%2Fs10694-018-0714-2" 'Fire Technology - Predicting the Fire Dynamics of Exposed Timber Surfaces in Compartments Using a Two-Zone Model

    Public Const Atm_Pressure As Double = 101.325 'kPa
    Public Const gamma As Double = 1.4
    Public Const gammax As Double = 0.684731456
    Public Const gamcut As Double = 0.528281788 'from ventcf2a
    Public Const QSpread As Double = 30 'kW/m2  heat flux from the flame ahead of the pyrolysis region
    Public Const MaxNumberRooms As Integer = 12
    Public Const IgCriteria As Single = 30 'criteria for determine ignition time from cone data
    Public Const MaxConeCurves As Short = 6
    Public Const SDNormalDefault As Single = 0.097 '1/m
    Public Const SDhighDefault As Single = 0.055 '1/m
    Public Const SDvhighDefault As Single = 0.013 '1/m
    Public Const gcs_RTI_cspr As Single = 95
    Public Const gcs_RTI_rspr As Single = 36
    Public Const gcs_RTI_hd As Single = 30
    Public Const gcs_cfactor As Single = 0.4
    Public Const gcs_sprdist As Single = 0.02
    Public Const gcs_DischargeCoeff As Double = 0.68 'discharge coefficient of the vent
    Public Const gcs_SprinkCoolCoeff As Single = 1 'cooling coefficient of the sprinkler - see Crocker et al, Fire Technology, 46, 347–362, 2010
    'gcs_SprinkCoolCoeff = 0.84 Tyco LFII residential pendent sprinkler (TY2234)
    Public timeunit As Integer = 60
    Public chartemp As Double
    Public sprnum_prob(0 To 3) As Decimal
    Public SprReliability As Decimal
    Public SDReliability As Decimal
    Public FanReliability As Decimal

    Public SprSuppressionProb As Decimal
    Public SprCooling As Decimal
    Public StoreHeight As Single
    Public NumOperatingSpr As Integer
    'Public sprfile As String
    Public gb_first_time_vent As Boolean
    Public fueltype As String
    Public batch As Boolean
    Public OtherMessages As String
    Public gRegisteredUser As Boolean
    Public UseOneCurve As Boolean
    Public wallparam(,) As Double
    Public CVentAuto(,,) As Boolean
    Public enzbreakflag(,,) As Boolean 'wall vent
    Public breakflag(,,) As Boolean 'wall vent
    Public breakflag2(,,) As Boolean 'ceiling vent
    Public HOFlag(,,) As Boolean
    Public HOactive(,,) As Boolean
    Public FANactive() As Boolean
    Public calcFRR As Boolean
    Public cpux() As Double
    Public cplx() As Double
    Public IgnCorrelation As Short
    Public cjModel As Short
    Public PrinterFontName As String
    Public PrinterFontSize As String
    Public LEsolver As String = ""
    Public PessimiseCombWall As Boolean
    Public gsDatabase As String
    Public gsRecordSource As String
    'Global PagePrinter As PPrinter
    Public StartOccupied As Integer
    Public MaxNumberVents As Short
    Public MaxNumberCVents As Short
    Public EndOccupied As Integer
    Public UserName As String
    Public thermal As String
    Public Flashover As Boolean
    Public FLED As Single
    Public HoC_fuel As Double
    Public NewHoC_fuel As Double
    Public InitialFuelMass As Double
    'Public FuelHeatofGasification As Single
    Public FuelSurfaceArea As Single
    Public nC As Single
    Public nH As Single
    Public nO As Single
    Public nN As Single
    Public Enhance As Boolean
    Public useCLTmodel As Boolean
    Public Fuel_Thickness As Double
    Public fuelmasswithCLT As Double
    Public Stick_Spacing As Double
    Public Cribheight As Double
    Public ExtractRate() As Single
    Public NumberFans() As Short
    Public ExtractStartTime() As Single
    Public FanElevation() As Double
    Public MaxPressure() As Single
    Public Extract() As Boolean
    Public UseFanCurve() As Boolean
    Public fanon() As Boolean
    Public WallSurface() As String 'material description
    Public WallSubstrate() As String 'material description
    Public CeilingSurface() As String 'material description
    Public CeilingSubstrate() As String 'material description
    Public FloorSubstrate() As String 'material description
    Public FloorSurface() As String 'material description
    Public HaveCeilingSubstrate() As Boolean
    Public HaveWallSubstrate() As Boolean
    Public HaveFloorSubstrate() As Boolean
    Public HaveSD() As Boolean
    Public SDinside() As Boolean
    Public SpecifyOD() As Boolean
    Public SDFlag() As Short
    Public SDFlag2() As Short
    Public SDFlagSD() As Short
    Public SDtransit(,) As Double
    Public SDTime() As Single
    Public SDTimeSD() As Single
    Public progdir As String
    Public RoomLength() As Double 'array 1 to number of rooms
    Public RoomDescription() As String 'array 1 to number of rooms
    Public FloorElevation() As Double 'array 1 to number of rooms
    Public RoomAbsX() As Single 'array 1 to number of rooms
    Public RoomAbsY() As Single 'array 1 to number of rooms
    Public RoomWidth() As Double 'array 1 to number of rooms
    Public RoomHeight() As Double 'array 1 to number of rooms
    Public MinStudHeight() As Double 'array 1 to number of rooms
    Public RoomVolume() As Double 'array 1 to number of rooms
    'Public WallInsideBiot As Double
    'Public CeilingInsideBiot As Double
    'Public FloorInsideBiot As Double
    Public WallOutsideBiot() As Double
    Public CeilingOutsideBiot() As Double
    Public FloorOutsideBiot() As Double
    Public WallFourier() As Double
    Public CeilingFourier() As Double
    Public FloorFourier() As Double
    Public AlphaWall() As Double
    Public AlphaCeiling() As Double
    Public AlphaFloor() As Double
    Public WallDeltaX() As Double
    Public CeilingDeltaX() As Double
    Public FloorDeltaX() As Double
    Public CeilingConductivity() As Double
    Public CeilingSpecificHeat() As Double
    Public CeilingDensity() As Double
    Public CeilingSubConductivity() As Double
    Public CeilingSubSpecificHeat() As Double
    Public CeilingSubDensity() As Double
    Public FloorConductivity() As Double
    Public FloorSpecificHeat() As Double
    Public FloorDensity() As Double
    Public WallConductivity() As Double
    Public WallSpecificHeat() As Double
    Public WallDensity() As Double
    Public WallThickness() As Double
    Public CeilingThickness() As Double
    Public WallSubConductivity() As Double
    Public FloorSubConductivity() As Double
    Public WallSubSpecificHeat() As Double
    Public FloorSubSpecificHeat() As Double
    Public WallSubDensity() As Double
    Public FloorSubDensity() As Double
    Public WallSubThickness() As Double
    Public CeilingSubThickness() As Double
    Public FloorSubThickness() As Double
    Public FloorThickness() As Double
    Public NumberRooms As Integer
    Public NumberVents(,) As Short
    Public NumberCVents(,) As Short
    Public Highest_Vent() As Short
    Public Lowest_Vent() As Short
    Public RoomFloorArea() As Double
    Public SootMassGen() As Single
    Public InteriorTemp As Double
    Public ExteriorTemp As Double
    Public RelativeHumidity As Single
    Public EmissionCoefficient As Single
    Public StoichAFratio As Double
    Public preCO As Single
    Public postCO As Single
    Public preSoot As Single
    Public postSoot As Single
    Public sootmode As Boolean
    Public comode As Boolean
    Public SootAlpha As Single
    Public SootEpsilon As Single
    'Public MLUnitArea As Single
    Public NumberObjects As Short 'number of burning objects
    'Public RadiantLossFraction As Single
    Public NewRadiantLossFraction() As Single
    Public SootFactor As Single
    Public MonitorHeight As Single
    Public Activity As String
    Public illumination As Boolean
    Public g_post As Boolean
    Public corner As Short
    Public Description As String
    Public JobNumber As String
    Public opendatafile As String
    Public DataFile As String
    Public PlumeModel As String
    Public ConfirmfileName As Short
    Public NumberDataPoints() As Short
    Public NumberTimeSteps As Integer
    Public SimTime As Integer
    Public DisplayInterval As Integer
    Public ExcelInterval As Integer
    Public RTI() As Single
    Public Sprinkler(,) As Boolean
    Public DetectorType() As Short
    Public cfactor() As Single
    Public WaterSprayDensity() As Single
    Public SprinkDistance() As Single
    Public SmokeOD() As Single
    Public SDdelay() As Single
    Public DetSensitivity() As Single
    Public SDRadialDist() As Single
    Public SDdepth() As Single
    Public RadialDistance() As Double
    Public ActuationTemp() As Single
    Public FEDEndPoint As Single
    Public VisibilityEndPoint As Single
    Public ConvectEndPoint As Single
    Public TempEndPoint As Single
    Public TargetEndPoint As Single
    Public runtime As Double
    Public stepcount As Integer
    Public ERRNO As Short
    Public FlameVelocityFlag As Short
    Public FloorVelocityFlag As Short
    Public flagstop As Short
    Public NeutralPlaneFlag As Short
    Public burner_id As Integer
    Public WallIgniteObject() As Short
    Public CeilingIgniteObject() As Short
    Public WallIgniteFlag() As Short 'use with quintiere's room corner model
    Public CeilingIgniteFlag() As Short 'use with quintiere's room corner model
    Public FloorIgniteFlag() As Short 'use with quintiere's room corner model
    Public YPFlag As Short 'use with quintiere's room corner model
    Public YBFlag As Short 'use with quintiere's room corner model
    Public AreaFlag As Short
    Public BurnerFlameLength As Single
    Public ComboTopLeftIndex As Short
    Public ComboTopRightIndex As Short
    Public ComboBottomLeftIndex As Short
    Public ComboBottomRightIndex As Short
    Public SprinklerTime As Single
    Public HDTime As Single
    Public SprinklerHRR As Single
    Public SprinklerFlag As Short
    Public HDFlag As Short
    Public FHflag As Boolean
    Public fireroom As Integer
    Public graphroom As Short
    Public HRR_total As Double
    Public soffitheight(,,) As Single
    Public COYield As Single
    Public EnergyYield() As Single
    Public NewEnergyYield() As Single
    Public HCNYield() As Single
    Public HCNuserYield() As Single
    Public ObjLabel() As String
    Public ObjLength() As Single
    Public ObjWidth() As Single
    Public ObjHeight() As Single
    Public ObjDimX() As Single
    Public ObjDimY() As Single
    Public Item1X() As Single
    Public Item1Y() As Single
    Public ObjElevation() As Single
    Public ObjCRF() As Single
    Public ObjFTPindexpilot() As Single
    Public ObjFTPlimitpilot() As Single
    Public ObjCRFauto() As Single
    Public ObjFTPindexauto() As Single
    Public ObjFTPlimitauto() As Single
    Public CO2Yield() As Single
    Public SootYield() As Single
    Public WaterVaporYield() As Single
    Public ObjectDescription() As String
    Public ObjectMass() As Double
    Public ObjectItemID() As Integer
    Public ObjectIgnTime() As Double
    Public ObjectLHoG() As Single
    Public ObjectMLUA(,) As Single
    Public ObjectIgnMode() As String
    Public ObjectRLF() As Single
    Public ObjectWindEffect() As Single
    Public ObjectPyrolysisOption() As Integer
    Public ObjectPoolFBMLR() As Double
    Public ObjectPoolVapTemp() As Double
    Public ObjectPoolDensity() As Double
    Public ObjectPoolRamp() As Double
    Public ObjectPoolVolume() As Double
    Public ObjectPoolDiameter() As Double
    Public FireHeight() As Single
    Public FireLocation() As Short
    Public CVentArea(,,) As Double
    Public AutoFlag() As Boolean
    Public VentHeight(,,) As Double
    Public VentFace(,,) As Integer
    Public VentOffset(,,) As Single
    Public spillplume(,,) As Short
    Public spillplumemodel(,,) As Short
    Public spillbalconyprojection(,,) As Single
    Public VentWidth(,,) As Double
    Public FRintegrity(,,) As Double
    Public FRMaxOpening(,,) As Double
    Public FRMaxOpeningTime(,,) As Double
    Public FRintegrity2(,,) As Double
    Public FRMaxOpening2(,,) As Double
    Public FRMaxOpeningTime2(,,) As Double
    Public FRgastemp2(,,) As Double
    Public FRgastemp(,,) As Double
    Public FRcriteria(,,) As Integer
    Public FRcriteria2(,,) As Integer
    Public FRfaildata(,,) As Double 'wall vent
    Public FRfaildata2(,,) As Double 'ceiling vent
    Public WallLength1(,,) As Single
    Public WallLength2(,,) As Single
    Public VentSillHeight(,,) As Double
    Public VentCD(,,) As Double
    Public VentProb(,,) As Single
    Public HOReliability(,,) As Single
    Public VentOpenTime(,,) As Single
    Public oldVentOpenTime(,,) As Single
    Public oldCVentOpenTime(,,) As Single
    Public VentBreakTime(,,) As Single
    Public CVentBreakTime(,,) As Single
    Public VentCloseTime(,,) As Single
    Public Downstand(,,) As Single
    Public GLASSconductivity(,,) As Single
    Public GLASSemissivity(,,) As Single
    Public GLASSexpansion(,,) As Single
    Public AutoBreakGlass(,,) As Boolean
    Public SpillPlumeBalc(,,) As Boolean
    Public GlassFlameFlux(,,) As Boolean
    Public GLASSthickness(,,) As Single
    Public GLASSdistance(,,) As Single
    Public GLASSFalloutTime(,,) As Single
    Public GLASSshading(,,) As Single
    Public GLASSbreakingstress(,,) As Single
    Public GLASSalpha(,,) As Single
    Public GlassYoungsModulus(,,) As Single
    Public GLASSTempHistory(,,,,) As Double
    Public GLASSOtherHistory(,,,,) As Double
    Public CVentOpenTime(,,) As Double
    Public CVentCloseTime(,,) As Double
    Public CVentDC(,,) As Double
    Public WAfloor_flag(,,) As Short
    Public HeatReleaseData(,,) As Double 'user input heat release
    Public MLRData(,,) As Double 'user input mass loss rate data
    Public tim(,) As Double 'time
    Public layerheight(,) As Double 'height layer interface above floor
    Public UpperVolume(,) As Double 'upper layer volume
    Public uppertemp(,) As Double 'average temp of the upper layer
    Public lowertemp(,) As Double 'lower layer temperature
    Public upperemissivity(,) As Double 'upper layer emissivity
    Public IgWallTemp(,) As Double 'wall temperatures adjacent to burner with quintiere's room corner model
    Public IgFloorTemp(,) As Double 'wall temperatures adjacent to burner with quintiere's room corner model
    Public IgCeilingTemp(,) As Double 'wall temperatures adjacent to burner with quintiere's room corner model
    Public Upperwalltemp(,) As Double 'upper wall temp
    Public UnexposedLowerwalltemp(,) As Double 'unexposed lower wall temp
    Public UnexposedUpperwalltemp(,) As Double 'unexposed upper wall temp
    Public UnexposedCeilingtemp(,) As Double 'unexposed ceiling temp
    Public UnexposedFloortemp(,) As Double 'unexposed floor temp
    Public LowerWallTemp(,) As Double 'lower wall temp
    Public CeilingTemp(,) As Double 'ceiling temp
    Public FloorTemp(,) As Double 'floor temp
    Public RoomPressure(,) As Double 'pressure
    Public HeatRelease(,,) As Double 'heat release of fire
    Public FuelMassLossRate(,) As Double 'fuel mass loss rate kg/s

    Public FuelBurningRate(,,) As Double 'fuel burning rate kg/s
    Public FBMLR_poolfire As Double 'freeburn mass loss rate kg/m2/s for pool fire
    Public pandia_poolfire As Double 'pan diameter for pool fire (m)
    Public fuelamount As Double 'mass of liquid for pool fire (kg)

    Public WoodBurningRate() As Double 'save wall + ceiling contributions in fire room to total kg/s
    Public FuelHoC() As Single
    Public CJetTemp(,,) As Double
    Public TotalFuel() As Double
    Public qloss() As Double 'heat losses to the ceiling and upper walls
    Public massplumeflow(,) As Double 'mass flow rate of air entrained into plume
    Public VentMassOut(,) As Double 'mass flow of hot gases out the vent
    Public FlowToUpper(,) As Double 'net mass flow of hot gases to the upper layer due to vent flow, door mixing and entrainment
    Public FlowToLower(,) As Double 'net mass flow of hot gases to the lower layer due to vent flow, door mixing and entrainment
    Public WallFlowtoUpper(,) As Double 'net mass flow of hot gases to the lower layer due to vent flow, door mixing and entrainment
    Public WallFlowtoLower(,) As Double 'net mass flow of hot gases to the lower layer due to vent flow, door mixing and entrainment
    Public UFlowToOutside(,) As Double 'mass flow of hot gases from the upper layer to outside
    Public FlowGER(,) As Double 'mass flow of hot gases from the upper layer to outside
    Public SPR(,) As Double 'smoke production rate (m2/s)
    Public ventfire(,) As Double 'total hrr in vent fire
    Public TUHC(,,) As Double 'mass fraction of total unburned hydrocarbons
    Public CO2MassFraction(,,) As Double 'mass fraction of CO2 in the upper layer
    Public CO2VolumeFraction(,,) As Double 'volume fraction of CO2 in the upper layer
    Public H2OVolumeFraction(,,) As Double 'volume fraction of H2O in the upper layer
    Public HCNVolumeFraction(,,) As Double 'volume fraction of HCN in the upper layer
    Public COMassFraction(,,) As Double 'mass fraction of CO in the upper layer
    Public SootMassFraction(,,) As Double 'mass fraction of soot
    Public HCNMassFraction(,,) As Double 'mass fraction of HCN
    Public H2OMassFraction(,,) As Double 'mass fraction of water vapor
    Public COVolumeFraction(,,) As Double 'volume fraction of CO in the upper layer
    Public OD_upper(,) As Double
    Public OD_lower(,) As Double
    Public Visibility(,) As Double
    Public O2MassFraction(,,) As Double 'mass fraction of O2 in the upper layer
    Public O2VolumeFraction(,,) As Double 'volume fraction of O2 in the upper layer
    Public FEDSum(,) As Double 'FED for incapacitation
    Public FEDRadSum(,) As Double 'FED for incapacitation
    Public SurfaceRad(,) As Double 'FED for incapacitation
    Public Target(,) As Double 'incident radiant flux on floor
    Public LinkTemp(,) As Double 'temperature of the detector or sprinkler link
    Public SprinkTemp(,) As Double
    Public SmokeConcentration(,) As Double 'at the detector
    Public SmokeConcentrationSD(,) As Double 'at the detector
    Public OD_outside(,) As Double 'at the detector
    Public OD_outsideSD(,) As Double 'at the detector
    Public OD_inside(,) As Double 'at the detector
    Public OD_insideSD(,) As Double 'at the detector
    Public CeilingSlope() As Boolean
    Public FanAutoStart() As Boolean
    Public TwoZones() As Boolean
    Public GlobalER() As Double 'global equivalence ratio
    Public absorb(,,) As Double
    Public ventflow(,,) As Double 'cw 18/3/2010
    Public VentLogFlag As Short

    'Room Corner Simulation Variables
    Public WallHeatFlux As Single 'heat flux to the wall material
    Public CeilingHeatFlux As Single 'heat flux to the ceiling material
    Public PeakWallHRR(,) As Double 'peak rate of heat release in cone calorimeter at 50 kW/m2
    Public PeakCeilingHRR(,) As Double 'peak rate of heat release in cone calorimeter at 50 kW/m2
    Public PeakFloorHRR(,) As Double 'peak rate of heat release in cone calorimeter at 50 kW/m2
    Public FlameAreaConstant As Single
    Public BurnerWidth As Single
    Public Qburner As Double
    Public WallAreaBurner As Single
    Public CornerTime() As Short
    Public ConeXW(,,) As Double
    Public ConeYW(,,) As Double
    Public ConeXF(,,) As Double
    Public ConeYF(,,) As Double
    Public ConeXC(,,) As Double
    Public ConeYC(,,) As Double
    Public ExtFlux_W(,) As Double
    Public ExtFlux_C(,) As Double
    Public ExtFlux_F(,) As Double
    Public ExtFlux() As Double
    Public ConePeak() As Double
    Public ConeHoC() As Single
    Public ConeSEA() As Single
    Public NormalHRR_W(,) As Double
    Public NormalHRR_C(,) As Double
    Public NormalHRR_F(,) As Double
    Public ConeNumber_W(,) As Integer
    Public ConeNumber_C(,) As Integer
    Public ConeNumber_F(,) As Integer
    Public CurveNumber_W() As Short
    Public CurveNumber_F() As Short
    Public CurveNumber_C() As Short
    Public wallhigh As Single 'time for cone data curve
    Public ceilinghigh As Single 'time for cone data curve
    Public floorhigh As Single 'time for cone data curve
    Public flamespread() As Double
    Public product() As Double
    Public IgTempW() As Double
    Public IgTempC() As Double
    Public IgTempF() As Double
    Public ThermalInertiaWall() As Double
    Public ThermalInertiaCeiling() As Double
    Public ThermalInertiaFloor() As Double
    Public WallConeDataFile() As String
    Public FloorConeDataFile() As String
    Public CeilingConeDataFile() As String
    Public WallSootYield() As Double
    Public CeilingSootYield() As Double
    Public FloorSootYield() As Double
    Public WallH2OYield() As Double
    Public CeilingH2OYield() As Double
    Public FloorH2OYield() As Double
    Public WallHCNYield() As Double
    Public CeilingHCNYield() As Double
    Public FloorHCNYield() As Double
    Public WallCO2Yield() As Double
    Public CeilingCO2Yield() As Double
    Public FloorCO2Yield() As Double
    Public FirstTime() As Boolean

    'Quintiere's Model
    Public BurnerDiameter As Double
    Public WallIgniteTime() As Double
    Public WallIgniteStep() As Short
    Public FloorIgniteTime() As Double
    Public CeilingIgniteTime() As Double
    Public CeilingIgniteStep() As Short
    Public FloorIgniteStep() As Short
    Public WallFlameSpreadParameter() As Double
    Public FloorFlameSpreadParameter() As Double
    Public WallTSMin() As Double
    Public FloorTSMin() As Double
    Public WallEffectiveHeatofCombustion() As Double
    Public WallFTP() As Double
    Public Walln() As Single
    Public WallQCrit() As Double
    Public CeilingQCrit() As Double
    Public FloorQCrit() As Double
    Public CeilingFTP() As Double
    Public FloorFTP() As Double
    Public Ceilingn() As Single
    Public Floorn() As Single
    Public WallHeatofGasification() As Double
    Public CeilingEffectiveHeatofCombustion() As Double
    Public FloorEffectiveHeatofCombustion() As Double
    Public CeilingHeatofGasification() As Double
    Public FloorHeatofGasification() As Double
    Public X_pyrolysis(,) As Double
    Public Xf_pyrolysis(,) As Double
    Public Z_pyrolysis(,) As Double
    Public FlameVelocity(,,) As Double
    Public FloorVelocity(,,) As Double
    Public Y_pyrolysis(,) As Double
    Public Yf_pyrolysis(,) As Double
    'Global X_burnout() As Double
    Public Y_burnout(,) As Double
    Public Yf_burnout(,) As Double
    Public timeH As Double 'time at which the pyrolysis front reaches the ceiling
    'Global timeH2 As Double 'time at which burnout reaches the ceiling
    'Global tbo As Double 'sec initial burnout time
    Public AreaUnderWallCurve() As Double 'area under the hrr curve (MJ/m2)
    Public AreaUnderCeilingCurve() As Double 'area under the hrr curve (MJ/m2)
    Public AreaUnderFloorCurve() As Double 'area under the hrr curve (MJ/m2)
    Public FlameLengthPower As Double
    Public QFB As Double 'burner flame flux
    Public flagcount As Short
    Public pyrolarea(,,) As Double
    Public delta_area(,,) As Double
    Public QPyrol As Double

    'CLT stuff
    Public QWall1 As Double
    Public QCeiling1 As Double
    Public QFloor1 As Double

    'adjacent room ignition
    Public IgniteNextRoom As Boolean

    '4 wall heat transfer radiation model
    Public F(,,) As Double 'view factors
    Public TransmissionFactor(,,) As Double
    Public Surface_Emissivity(,) As Double
    Public QCeiling(,) As Double
    Public QUpperWall(,) As Double
    Public QLowerWall(,) As Double
    Public QFloor(,) As Double
    Public QCeilingAST(,,) As Double 'room,inc rad flux=0 conv coeff=1, net rad flux=1,stepcount 
    Public QUpperWallAST(,,) As Double
    Public QLowerWallAST(,,) As Double
    Public QFloorAST(,,) As Double

    Public TotalLosses(,) As Double
    Public NHL(,,) As Double
    Public radfireUpper As Double 'temp variable for holding rad heat absorbed by upper layer due to point source fire
    Public radfireLower As Double
    Public Timestep As Double
    Public Wallnodes As Short
    Public Ceilingnodes As Short
    Public Floornodes As Short

    Public Error_Control As Double
    Public Error_Control_ventflow As Double
    Public MW_air As Double
    Public MolecularWeightFuel As Double
    Public SpecificHeat_air As Double '1.023 'specific heat of air kJ/kgK

    'constants to represent the file-access mode
    Public Const savefile As Short = 1
    Public Const loadfile As Short = 2
    Public Const replacefile As Short = 1
    Public Const readfile As Short = 3
    Public Const addtofile As Short = 3
    Public Const randomfile As Short = 4
    Public Const binaryfile As Short = 5

    'constants to represent error conditions
    Public Const err_deviceunavailable As Short = 68
    Public Const err_disknotready As Short = 71
    Public Const err_filealreadyexists As Short = 58
    Public Const err_toomanyfiles As Short = 67
    Public Const err_renameacrossdisks As Short = 74
    Public Const err_path_fileaccesserror As Short = 75
    Public Const err_deviceio As Short = 57
    Public Const err_diskfull As Short = 61
    Public Const err_badfilename As Short = 64
    Public Const err_badfilenameornumber As Short = 52
    Public Const err_filenotfound As Short = 53
    Public Const err_pathdoesnotexist As Short = 76
    Public Const err_badfilemode As Short = 54
    Public Const err_filealreadyopen As Short = 55
    Public Const err_inputpastendoffile As Short = 62
    Public Const err_overflow As Short = 6

    'MsgBox parameters
    Public Const MB_OK As Short = 0 ' OK button only
    Public Const MB_OKCANCEL As Short = 1 ' OK and Cancel buttons
    Public Const MB_ABORTRETRYIGNORE As Short = 2 ' Abort, Retry, and Ignore buttons
    Public Const MB_YESNOCANCEL As Short = 3 ' Yes, No, and Cancel buttons
    Public Const MB_YESNO As Short = 4 ' Yes and No buttons
    Public Const MB_RETRYCANCEL As Short = 5 ' Retry and Cancel buttons

    Public Const MB_ICONSTOP As Short = 16
    Public Const MB_STOP As Short = 16 'Critical message
    Public Const MB_ICONQUESTION As Short = 32 ' Warning query
    Public Const MB_ICONEXCLAMATION As Short = 48
    Public Const MB_EXCLAIM As Short = 48 ' Warning message
    Public Const MB_ICONINFORMATION As Short = 64 ' Information message

    Public Const MB_DEFBUTTON1 As Short = 0 ' First button is default
    Public Const MB_DEFBUTTON2 As Short = 256 ' Second button is default
    Public Const MB_DEFBUTTON3 As Short = 512 ' Third button is default

    'MsgBox return values
    Public Const IDOK As Short = 1 ' OK button pressed
    Public Const IDCANCEL As Short = 2 ' Cancel button pressed
    Public Const IDABORT As Short = 3 ' Abort button pressed
    Public Const IDRETRY As Short = 4 ' Retry button pressed
    Public Const IDIGNORE As Short = 5 ' Ignore button pressed
    Public Const IDYES As Short = 6 ' Yes button pressed
    Public Const IDNO As Short = 7 ' No button pressed

    '================='
    '                 '
    ' Property values '
    '                 '
    '================='

    ' BackColor, ForeColor, FillColor (standard RGB colors: form, controls)
    Public Const BLACK As Integer = &H0
    Public Const RED As Integer = &HFF
    Public Const GREEN As Integer = &HFF00
    Public Const YELLOW As Integer = &HFFFF
    Public Const BLUE As Integer = &HFF0000
    Public Const MAGENTA As Integer = &HFF00FF
    Public Const CYAN As Integer = &HFFFF00
    Public Const WHITE As Integer = &HFFFFFF

    ' System Colors
    Public Const SCROLL_BARS As Integer = &H80000000 ' Scroll-bars gray area.
    Public Const DESKTOP As Integer = &H80000001 ' Desktop.
    Public Const ACTIVE_TITLE_BAR As Integer = &H80000002 ' Active window caption.
    Public Const INACTIVE_TITLE_BAR As Integer = &H80000003 ' Inactive window caption.
    Public Const MENU_BAR As Integer = &H80000004 ' Menu background.
    Public Const WINDOW_BACKGROUND As Integer = &H80000005 ' Window background.
    Public Const WINDOW_FRAME As Integer = &H80000006 ' Window frame.
    Public Const MENU_TEXT As Integer = &H80000007 ' Text in menus.
    Public Const WINDOW_TEXT As Integer = &H80000008 ' Text in windows.
    Public Const TITLE_BAR_TEXT As Integer = &H80000009 ' Text in caption, size box, scroll-bar arrow box..
    Public Const ACTIVE_BORDER As Integer = &H8000000A ' Active window border.
    Public Const INACTIVE_BORDER As Integer = &H8000000B ' Inactive window border.
    Public Const APPLICATION_WORKSPACE As Integer = &H8000000C ' Background color of multiple document interface (MDI) applications.
    Public Const HIGHLIGHT As Integer = &H8000000D ' Items selected item in a control.
    Public Const HIGHLIGHT_TEXT As Integer = &H8000000E ' Text of item selected in a control.
    Public Const BUTTON_FACE As Integer = &H8000000F ' Face shading on command buttons.
    Public Const BUTTON_SHADOW As Integer = &H80000010 ' Edge shading on command buttons.
    Public Const GRAY_TEXT As Integer = &H80000011 ' Grayed (disabled) text.  This color is set to 0 if the current display driver does not support a solid gray color.
    Public Const BUTTON_TEXT As Integer = &H80000012 ' Text on push buttons.

    ' BorderStyle (form, label, picture box, text box)
    Public Const NONE As Short = 0 ' 0 - None
    Public Const FIXED_SINGLE As Short = 1 ' 1 - Fixed Single
    Public Const SIZABLE As Short = 2 ' 2 - Sizable (Forms only)
    Public Const FIXED_DOUBLE As Short = 3 ' 3 - Fixed Double (Forms only)

    ' MousePointer (form, controls)
    'UPGRADE_NOTE: Default was upgraded to Default_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public Const Default_Renamed As Short = 0 ' 0 - Default
    Public Const ARROW As Short = 1 ' 1 - Arrow
    Public Const CROSSHAIR As Short = 2 ' 2 - Cross
    Public Const IBEAM As Short = 3 ' 3 - I-Beam
    Public Const ICON_POINTER As Short = 4 ' 4 - Icon
    Public Const SIZE_POINTER As Short = 5 ' 5 - Size
    Public Const SIZE_NE_SW As Short = 6 ' 6 - Size NE SW
    Public Const SIZE_N_S As Short = 7 ' 7 - Size N S
    Public Const SIZE_NW_SE As Short = 8 ' 8 - Size NW SE
    Public Const SIZE_W_E As Short = 9 ' 9 - Size W E
    Public Const UP_ARROW As Short = 10 ' 10 - Up Arrow
    Public Const HOURGLASS As Short = 11 ' 11 - Hourglass
    Public Const NO_DROP As Short = 12 ' 12 - No drop

    ' ScrollBar (text box)
    ' Global Const NONE As Integer = 0                ' 0 - None
    Public Const HORIZONTAL As Short = 1 ' 1 - Horizontal
    Public Const VERTICAL As Short = 2 ' 2 - Vertical
    Public Const BOTH As Short = 3 ' 3 - Both

    ' Value (check box)
    Public Const UNCHECKED As Short = 0 ' 0 - Unchecked
    Public Const CHECKED As Short = 1 ' 1 - Checked
    Public Const GRAYED As Short = 2 ' 2 - Grayed

    ' WindowState (form)
    Public Const NORMAL As Short = 0 ' 0 - Normal
    Public Const MINIMIZED As Short = 1 ' 1 - Minimized
    Public Const MAXIMIZED As Short = 2 ' 2 - Maximized

    'Common Dialog Control
    'Action Property
    Public Const DLG_FILE_OPEN As Short = 1
    Public Const DLG_FILE_SAVE As Short = 2
    Public Const DLG_COLOR As Short = 3
    Public Const DLG_FONT As Short = 4
    Public Const DLG_PRINT As Short = 5
    Public Const DLG_HELP As Short = 6

    'File Open/Save Dialog Flags
    Public Const OFN_READONLY As Integer = &H1
    Public Const OFN_OVERWRITEPROMPT As Integer = &H2
    Public Const OFN_HIDEREADONLY As Integer = &H4
    Public Const OFN_NOCHANGEDIR As Integer = &H8
    Public Const OFN_SHOWHELP As Integer = &H10
    Public Const OFN_NOVALIDATE As Integer = &H100
    Public Const OFN_ALLOWMULTISELECT As Integer = &H200
    Public Const OFN_EXTENSIONDIFFERENT As Integer = &H400
    Public Const OFN_PATHMUSTEXIST As Integer = &H800
    Public Const OFN_FILEMUSTEXIST As Integer = &H1000
    Public Const OFN_CREATEPROMPT As Integer = &H2000
    Public Const OFN_SHAREAWARE As Integer = &H4000
    Public Const OFN_NOREADONLYRETURN As Integer = &H8000
End Module