Option Strict Off
Option Explicit On
Module Glass
	Dim m_bFirstRunFlag As Boolean 'Module first run flag = TRUE if first run completed
    Dim m_bFirstRun(,,) As Boolean 'Vent First run flag = TRUE if first run completed
    Dim m_dblTemps(,,,) As Double 'Through-thickness nodal temperature array
    Dim m_dblFractureTemp(,,) As Double 'Calculated Fracture temperature (K)
	Dim m_dblHmax As Object
	Dim m_dblHmin As Double 'Upper and lower limits for interior heat tx coeff
	Dim m_dblHexterior As Double 'Exterior heat transfer coefficient
    Dim m_dblprevTempUL(,,) As Double 'May be needed if delta time too big
    Dim m_dblRhocp(,,) As Double 'Density*specific heat = k/alpha
    Dim m_dblDeltaX(,,) As Double 'Spacing between nodes
    Dim m_dblOldTimeElapsed(,,) As Double 'Used to calculate delta time
    Dim m_iTimeSteps(,,) As Short 'Number of times the sub has been called
	
    Public Sub Break_Glass(ByRef idc As Integer, ByRef idr As Integer, ByRef idv As Integer, ByRef TimeElapsed As Double, ByRef breakflag(,,) As Boolean, ByRef vheight As Double, ByRef vwidth As Double, ByRef TempUL As Double, ByRef eUL As Double)
        ', FluxtoUpperWall As Double,  LowerEmm As Double, AmbientEmm)
        '==============================================================================
        'Ross Parry, 11/2001 - Canterbury University ME(Fire) Research Project
        '
        'Main Reference: Sincaglia, P. E., and Barnett, J. R., (1997)
        '                "Development of a Glass Window Fracture Model for Zone-Type
        '                Computer Fire Codes" in
        '                J. of Fire Prot. Engr., Vol 8, No. 3, pp 108-118.
        '===============================================================================
        '   glass breakage subroutine - called once at each timestep
        '
        '   TimeElapsed - current elapsed time
        '   vheight - height of vent (m)
        '   vwidth - width of vent (m)
        '   dblTempUL - exposing layer temperature of current compartment
        '   eUL - emissivity of the exposing layer
        '
        '==============================================================================
        'Explicit finite-difference method used to calculate nodal temperatures in
        'one-dimension, through the thickness of the window pane. Average temperature
        'used with Pagni and Joshi's fracture criterion.
        'Heat transfer is by convective and through-thickness radiation absorption.
        '3-band absorptivity used for radiation absorption.
        '====================================================================================

        Dim iCounter As Short
        Dim n As Short 'Number of Nodes
        Dim dblKprime As Double 'Apparent conduction
        Dim dblHinterior As Double 'Internal convection transfer coeff, hI
        Dim dblNewTemps() As Double 'Temperatures calculated this timestep
        Dim dblDeltaTime As Double 'Timestep size
        Dim dblGlassTemp As Double 'Average through-thickness glass temp
        Dim sBreakMsg As Object
        Dim sBreakTitle As String
        Dim sTooThickMsg As Object
        Dim sTooThickTitle As String
        Dim sTooSlowMsg As Object
        Dim sTooSlowTitle As String
        'Dim sStabilityMsg As Object
        'Dim sStabilityTitle As String
        Dim dblTimeCriterion As Double
        Dim dblHalfWidth, dblSigma As Object
        Dim dblBeta As Double
        Dim dblK As Object
        Dim dblE As Double 'Conductivity, Young's modulus

        sBreakMsg = "Vent No. " & idv & " between rooms " & idc & " and " & idr & " has fractured after " & TimeElapsed & " seconds."
        sBreakTitle = "Glass Fracture"
        sTooThickMsg = "Vent No. " & idv & " between rooms " & idc & " and " & idr & " does not have a sufficiently insulated edge for the glass" & " fracture model to be applicable."
        sTooThickTitle = "Window Thickness or Shading Error"
        sTooSlowMsg = "The rate of heating is too slow to calculate the fracture" & " time for vent No. " & idv & " between rooms " & idc & " and " & idr & "."
        sTooSlowTitle = "Window Heated Too Slowly"

        If System.Math.Round(GLASSthickness(idc, idr, idv)) < 10 Then
            n = 10
        Else
            n = System.Math.Round(GLASSthickness(idc, idr, idv)) + 1
        End If

        ReDim dblNewTemps(n) 'Size new temperature array to no. of nodes

        'Size persistant arrays
        If m_bFirstRunFlag = False Then
            ReDim m_bFirstRun(NumberRooms + 1, NumberRooms + 1, MaxNumberVents)
            ReDim m_dblTemps(NumberRooms + 1, NumberRooms + 1, MaxNumberVents, n)
            ReDim m_dblFractureTemp(NumberRooms + 1, NumberRooms + 1, MaxNumberVents)
            ReDim m_dblprevTempUL(NumberRooms + 1, NumberRooms + 1, MaxNumberVents)
            ReDim m_dblRhocp(NumberRooms + 1, NumberRooms + 1, MaxNumberVents)
            ReDim m_dblDeltaX(NumberRooms + 1, NumberRooms + 1, MaxNumberVents)
            ReDim m_dblOldTimeElapsed(NumberRooms + 1, NumberRooms + 1, MaxNumberVents)
            ReDim m_iTimeSteps(NumberRooms + 1, NumberRooms + 1, MaxNumberVents)
            ReDim breakflag(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim GLASSTempHistory(NumberRooms + 1, NumberRooms + 1, MaxNumberVents, n, 1)
            ReDim GLASSOtherHistory(NumberRooms + 1, NumberRooms + 1, MaxNumberVents, 2, 1)
            m_bFirstRunFlag = True
        End If

        If breakflag(idc, idr, idv) = True Then Exit Sub

        'Initialise all persistant variables for this window
        If m_bFirstRun(idc, idr, idv) = False Then
            'Find biggest half-width for earliest fracture time
            If vheight > vwidth Then
                'UPGRADE_WARNING: Couldn't resolve default property of object dblHalfWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                dblHalfWidth = vheight / 2
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object dblHalfWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                dblHalfWidth = vwidth / 2
            End If

            'Set nodes to ambient temperature
            For iCounter = 1 To n
                m_dblTemps(idc, idr, idv, iCounter) = InteriorTemp
            Next

            'Set up the window properties for this vent
            dblSigma = GLASSbreakingstress(idc, idr, idv)
            dblBeta = GLASSexpansion(idc, idr, idv)
            dblK = GLASSconductivity(idc, idr, idv)
            'dblE = 72000
            dblE = GlassYoungsModulus(idc, idr, idv)
            m_dblRhocp(idc, idr, idv) = dblK / GLASSalpha(idc, idr, idv)
            m_dblDeltaX(idc, idr, idv) = GLASSthickness(idc, idr, idv) / 1000 / (n - 1)

            'Pagni and Joshi's fracture criterion:
            m_dblFractureTemp(idc, idr, idv) = InteriorTemp + (1 + 0.001 * GLASSshading(idc, idr, idv) / dblHalfWidth) * dblSigma / (dblE * dblBeta)
            m_dblHmin = 5
            m_dblHmax = 50
            m_dblHexterior = 10
            m_dblprevTempUL(idc, idr, idv) = TempUL
            m_dblOldTimeElapsed(idc, idr, idv) = 0
            m_iTimeSteps(idc, idr, idv) = 1

            'Check applicability criterion: s >= 2L
            If 0.001 * GLASSshading(idc, idr, idv) < 2 * 0.001 * GLASSthickness(idc, idr, idv) Then
                Call MsgBox(sTooThickMsg, MsgBoxStyle.OkOnly, sTooThickTitle)
            End If

            m_bFirstRun(idc, idr, idv) = True
        End If

        'Resize arrays for this timestep
        If m_iTimeSteps(idc, idr, idv) <> 1 Then
            ReDim Preserve GLASSTempHistory(NumberRooms + 1, NumberRooms + 1, MaxNumberVents, n, m_iTimeSteps(idc, idr, idv))
        End If
        ReDim Preserve GLASSOtherHistory(NumberRooms + 1, NumberRooms + 1, MaxNumberVents, 2, m_iTimeSteps(idc, idr, idv))

        'Check applicability criterion: t <= s^2 / alpha  -> Heating speed
        If TimeElapsed > 0.001 * GLASSshading(idc, idr, idv) ^ 2 / GLASSalpha(idc, idr, idv) Then
            Call MsgBox(sTooSlowMsg, MsgBoxStyle.OkOnly, sTooSlowTitle)
            'end sub???
        End If

        'Calculate internal heat transfer coefficient
        If TempUL >= 400 Then
            dblHinterior = m_dblHmax
        ElseIf TempUL <= 300 Then
            dblHinterior = m_dblHmin
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object m_dblHmax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            dblHinterior = m_dblHmin + (m_dblHmax - m_dblHmin) * (TempUL - 300) / 100
        End If

        'Dim insideconvcoeff As Double
        'If stepcount > 1 Then insideconvcoeff = Get_Convection_Coefficient2(ByVal vheight * vwidth, ByVal TempUL, ByVal GLASSTempHistory(idc, idr, idv, 1, m_iTimeSteps(idc, idr, idv) - 1), VERTICAL)
        'dblHinterior = insideconvcoeff

        'Calculate values needed for determining timestep size
        dblKprime = 0.7222 + 0.001583 * (TempUL - 273.15) 'Endry Turzik Correlation
        dblDeltaTime = TimeElapsed - m_dblOldTimeElapsed(idc, idr, idv)

        dblTimeCriterion = m_dblRhocp(idc, idr, idv) * m_dblDeltaX(idc, idr, idv) ^ 2 / (2 * (dblHinterior * m_dblDeltaX(idc, idr, idv) + dblKprime))

        If dblDeltaTime > dblTimeCriterion Then
            Call NodalTemps2(TempUL, dblDeltaTime, Int(dblDeltaTime / dblTimeCriterion + 1), eUL, idc, idr, idv, n)
        Else
            Call NodalTemps2(TempUL, dblDeltaTime, 1, eUL, idc, idr, idv, n)
        End If

        m_dblOldTimeElapsed(idc, idr, idv) = TimeElapsed
        m_dblprevTempUL(idc, idr, idv) = TempUL

        'Calculate the average temperature through the thickness of the glass
        For iCounter = 1 To n
            dblGlassTemp = dblGlassTemp + m_dblTemps(idc, idr, idv, iCounter)
            GLASSTempHistory(idc, idr, idv, iCounter, m_iTimeSteps(idc, idr, idv)) = m_dblTemps(idc, idr, idv, iCounter)
        Next iCounter

        dblGlassTemp = dblGlassTemp / n

        GLASSOtherHistory(idc, idr, idv, 1, m_iTimeSteps(idc, idr, idv)) = TempUL
        GLASSOtherHistory(idc, idr, idv, 2, m_iTimeSteps(idc, idr, idv)) = eUL

        m_iTimeSteps(idc, idr, idv) = m_iTimeSteps(idc, idr, idv) + 1

        'Change flag if glass has fractured
        If dblGlassTemp >= m_dblFractureTemp(idc, idr, idv) And breakflag(idc, idr, idv) = False Then
            breakflag(idc, idr, idv) = True
            Dim Message As String = TimeElapsed.ToString & " sec. Glass in Vent " & CStr(idc) & "-" & CStr(idr) & "-" & CStr(idv) & " fractured."
            frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
        End If

    End Sub
	Function Q(ByRef intRange As Short, ByRef dblTemp As Double, ByRef dblEmissivity As Double, ByVal dblLength As Double) As Double
		'=======================================================================
		'Ross Parry, 11/2001 - Canterbury University ME(Fire) Research Project
		'=======================================================================
		'
		'Calculates radiative heat fluxes from the hot upper layer based on three
		'wavelength bands, q1 : 0    <= lambda <= 2.75 micrometres
		'                  q2 : 2.75 <  lambda <= 4.5  micrometres
		'                  q3 : 4.5  <  lambda
		'========================================================================
		Const SIGMA As Double = 0.0000000567051 'Stefan-Boltzmann (W/(m2K4)
		Const C2 As Double = 14387.69 '(micrometres*kelvin)
		Const LAMBDA1 As Double = 2.75
		Const LAMBDA2 As Double = 4.5
		Const CONVERGESIZE As Double = 0.00000001
		Const MAXITS As Short = 20
		Const RHOBAR As Double = 0.057 'Average reflectivity
		Const GAMMA1 As Double = 35 'Absorption coefficients
		Const gamma2 As Double = 475 '
		
		Dim iCount As Short
		Dim dblZeta(2) As Double
		Dim dblFraction(3) As Double
		Dim dblNew As Double
		Dim iRange As Short
		Dim dblG0 As Double 'Non-Reflected Irradiation
		
		'Calculate dimensionless parameters, zeta
		dblZeta(1) = C2 / (LAMBDA1 * dblTemp)
		dblZeta(2) = C2 / (LAMBDA2 * dblTemp)
		
		'Calculate radiative energy fractions within each
		'wavelength band using converging series formula from Siegel and Howell
		For iRange = 1 To 2
			iCount = 1
			Do 
				dblNew = System.Math.Exp(-iCount * dblZeta(iRange)) / iCount * (dblZeta(iRange) ^ 3 + 3 * dblZeta(iRange) ^ 2 / iCount ^ 2 + 6 * dblZeta(iRange) / iCount ^ 2 + 6 / iCount ^ 3)
				If System.Math.Abs(dblNew) < CONVERGESIZE Then
					Exit Do
				Else : dblFraction(iRange) = dblFraction(iRange) + dblNew
					iCount = iCount + 1
				End If
			Loop While iCount < MAXITS
			'dblFraction(iRange) = dblFraction(iRange) * 15 / PI
			dblFraction(iRange) = dblFraction(iRange) * 15 / (PI ^ 4) 'cw
		Next iRange
		dblFraction(2) = dblFraction(2) - dblFraction(1)
		dblFraction(3) = 1 - dblFraction(2) - dblFraction(1)
		
		dblG0 = dblEmissivity * (1 - RHOBAR) * SIGMA * dblTemp ^ 4 'Non-reflected irradiation
		dblLength = 1.077 * dblLength 'Path length through element
		
		'Calculate radiant energy flow in each wavelength band
		Select Case intRange
			Case 1
				'Calculate q1
				Q = (1 - System.Math.Exp(-GAMMA1 * dblLength)) * dblFraction(1) * dblG0
			Case 2
				'Calculate q2
				Q = (1 - System.Math.Exp(-gamma2 * dblLength)) * dblFraction(2) * dblG0
			Case 3
				'Calculate q3
				Q = dblFraction(3) * dblG0
				'Q = GLASSemissivity() * (dblFraction(3) * dblG0 - SIGMA*st^4) 'cw
			Case Else
				Q = 0 'Error
		End Select
	End Function
	Function CylinderViewFactor(ByRef distance As Object, ByRef hrr As Double) As Object
		
		'=======================================================================
		'Ross Parry, 12/2001 - Canterbury University ME(Fire) Research Project
		'=======================================================================
		'
		'Calculates view factor from a flame to a vertical surface.
		'Worst case scenario - radiation to surface opposite base of cylindrical flame
		
		'View Factor Reference: Siegel, R. and Howell, J. R. (1992), "Thermal Radiation
		'Heat Transfer: Third Edition", Taylor & Francis, Washington, p1033
		'
		'Distance: Radial distance between the centre of the fire and the window
		'HRR: Fire heat release rate.
		'========================================================================
		
		Dim dblDiameter, dblRadius As Double
		Dim dblQcone As Double 'Peak RHR Cone Calorimeter
		Dim dblHeight As Double
		Dim dblL, dblH As Double
		Dim dblX, dblY As Double
		Dim dblTan2, dblTan1, dblTan3 As Double
		
		'dblQcone = 500      'kW/m2 = lower-bound typical value -> max radiation
		dblQcone = 130 'kW/m2 = lower-bound typical value -> max radiation
		
		dblDiameter = (4 * hrr / dblQcone / PI) ^ 0.5
		
		'View Factor approx 1 if flame very near window
		
		'UPGRADE_WARNING: Couldn't resolve default property of object distance. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If distance <= (dblDiameter / 2) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object CylinderViewFactor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CylinderViewFactor = 1
			Exit Function
		End If
		
		dblHeight = -1.02 * dblDiameter + 0.235 * hrr ^ 0.4 'Heskstad's flame height
		'correlation
		'lets take a position helf the flame height instead of the base of the flame and then double the cf
		dblHeight = dblHeight / 2 'cw 29/10/2002
		
		If dblHeight <= 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object CylinderViewFactor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CylinderViewFactor = 0
			Exit Function
		End If
		
		dblRadius = dblDiameter / 2
		dblL = dblHeight / dblRadius
		'UPGRADE_WARNING: Couldn't resolve default property of object distance. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		dblH = distance / dblRadius
		dblX = (1 + dblH) ^ 2 + dblL ^ 2
		dblY = (1 - dblH) ^ 2 + dblL ^ 2
		
		dblTan1 = System.Math.Atan(dblL / (dblH ^ 2 - 1) ^ 0.5)
		dblTan2 = System.Math.Atan((dblX * (dblH - 1) / dblY / (dblH + 1)) ^ 0.5)
		dblTan3 = System.Math.Atan(((dblH - 1) / (dblH + 1)) ^ 0.5)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object CylinderViewFactor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CylinderViewFactor = 1 / (PI * dblH) * dblTan1 + dblL / PI * ((dblX - 2 * dblH) / (dblH * (dblX * dblY) ^ 0.5) * dblTan2 - 1 / dblH * dblTan3)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object CylinderViewFactor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CylinderViewFactor = CylinderViewFactor * 2 'cw doubled it for half the flame height
		
	End Function
	Function FlameEmissivity(ByRef hrr As Double, Optional ByRef kappa As Double = 0.8) As Double
		
		'=======================================================================
		'Ross Parry, 12/2001 - Canterbury University ME(Fire) Research Project
		'=======================================================================
		'
		'Calculates emissivity from a flame.
		'
		'HRR: Fire heat release rate.
		'kappa: Emission coefficient - material dependent, use 0.8 as default
		'========================================================================
		
		Dim dblDiameter As Double
		Dim dblQcone As Double 'Peak RHR Cone Calorimeter
		
		kappa = EmissionCoefficient 'cw added 6/3/2002
		
		dblQcone = 500 'kW/m2 = lower-bound typical value -> max radiation
		dblDiameter = (4 * hrr / dblQcone / PI) ^ 0.5
		
		FlameEmissivity = 1 - System.Math.Exp(-kappa * dblDiameter)
		
	End Function
	Public Function fnHinterior(ByRef TempUL As Double) As Single
		If TempUL >= 400 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object m_dblHmax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fnHinterior = m_dblHmax
		ElseIf TempUL <= 300 Then 
			fnHinterior = m_dblHmin
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object m_dblHmax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fnHinterior = m_dblHmin + (m_dblHmax - m_dblHmin) * (TempUL - 300) / 100
		End If
	End Function
	
	
	Public Sub ResetWindows()
		'==============================================================================
		'Ross Parry, 11/2001 - Canterbury University ME(Fire) Research Project
		'==============================================================================
		'   Toggles FirstRun flag between simulations so that persistant
		'   variables can be recalculated.
		'==============================================================================
		m_bFirstRunFlag = False
	End Sub
	
	
	Public Sub NodalTemps2(ByRef FinalTempUL As Double, ByRef dblOldDeltaTime As Double, ByRef iDivisor As Short, ByRef eUL As Double, ByRef idc As Object, ByRef idr As Object, ByRef idv As Object, ByRef n As Short)
		
		'modified by C Wade 29/10/2002
		
		Dim dblDeltaTime As Double
		Dim Q1, Q2 As Object
        'Dim Q3 As Double 'Band fluxes for middle nodes
		Dim Q1a, Q2a As Object
		Dim Q3a As Double 'Band fluxes for interior and exterior nodes
		Dim iCounter As Short
		Dim iIncrement As Short
		Dim dblTempUL As Double
		Dim dblHinterior As Object
		Dim dblKprime As Double
		Dim dblNewTemps() As Double
		Dim dblViewFactor As Double
		Dim dblEmissivity As Double
		Dim qinfty As Double
		
		ReDim dblNewTemps(n)
		dblDeltaTime = dblOldDeltaTime / iDivisor
		
		For iIncrement = 1 To iDivisor
			'UPGRADE_WARNING: Couldn't resolve default property of object idv. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object idr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object idc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dblTempUL = m_dblprevTempUL(idc, idr, idv) + iIncrement / iDivisor * (FinalTempUL - m_dblprevTempUL(idc, idr, idv))
			
			If dblTempUL >= 400 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object m_dblHmax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object dblHinterior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				dblHinterior = m_dblHmax
				
			ElseIf dblTempUL <= 300 Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object dblHinterior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				dblHinterior = m_dblHmin
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object m_dblHmax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object dblHinterior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				dblHinterior = m_dblHmin + (m_dblHmax - m_dblHmin) * (dblTempUL - 300) / 100
			End If
			
			dblKprime = 0.7222 + 0.001583 * (dblTempUL - 273.15) 'Endry Turzik Correlation
			
			'UPGRADE_WARNING: Couldn't resolve default property of object idv. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object idr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object idc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Q1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Q1 = Qn(1, dblTempUL, eUL, m_dblDeltaX(idc, idr, idv), m_dblTemps(idc, idr, idv, 1), GLASSemissivity(idc, idr, idv))
			'UPGRADE_WARNING: Couldn't resolve default property of object idv. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object idr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object idc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Q2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Q2 = Qn(2, dblTempUL, eUL, m_dblDeltaX(idc, idr, idv), m_dblTemps(idc, idr, idv, 1), GLASSemissivity(idc, idr, idv))
			'Q3 = Qn(3, dblTempUL, eUL, m_dblDeltaX(idc, idr, idv), m_dblTemps(idc, idr, idv, 1), GLASSemissivity(idc, idr, idv))
			'UPGRADE_WARNING: Couldn't resolve default property of object idv. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object idr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object idc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Q1a. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Q1a = Qn(1, dblTempUL, eUL, m_dblDeltaX(idc, idr, idv) / 2, m_dblTemps(idc, idr, idv, 1), GLASSemissivity(idc, idr, idv))
			'UPGRADE_WARNING: Couldn't resolve default property of object idv. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object idr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object idc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Q2a. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Q2a = Qn(2, dblTempUL, eUL, m_dblDeltaX(idc, idr, idv) / 2, m_dblTemps(idc, idr, idv, 1), GLASSemissivity(idc, idr, idv))
			'UPGRADE_WARNING: Couldn't resolve default property of object idv. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object idr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object idc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Q3a = Qn(3, dblTempUL, eUL, m_dblDeltaX(idc, idr, idv) / 2, m_dblTemps(idc, idr, idv, 1), GLASSemissivity(idc, idr, idv))
			
			
			'UPGRADE_WARNING: Couldn't resolve default property of object idv. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object idr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object idc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If GlassFlameFlux(idc, idr, idv) = True Then
				'Add radiative heat transfer from the flame.
				'Assume 1073 K (=800 ºC)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object idv. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object idr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object idc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object CylinderViewFactor(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				dblViewFactor = CylinderViewFactor(GLASSdistance(idc, idr, idv), HeatRelease(idc, m_iTimeSteps(idc, idr, idv), 2))
				'UPGRADE_WARNING: Couldn't resolve default property of object idv. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object idr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object idc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				dblEmissivity = FlameEmissivity(HeatRelease(idc, m_iTimeSteps(idc, idr, idv), 1), 0.8)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object idv. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object idr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object idc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Q1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Q1 = Q1 + dblViewFactor * Qn(1, 1073, dblEmissivity, m_dblDeltaX(idc, idr, idv), m_dblTemps(idc, idr, idv, 1), GLASSemissivity(idc, idr, idv))
				'UPGRADE_WARNING: Couldn't resolve default property of object idv. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object idr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object idc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Q2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Q2 = Q2 + dblViewFactor * Qn(2, 1073, dblEmissivity, m_dblDeltaX(idc, idr, idv), m_dblTemps(idc, idr, idv, 1), GLASSemissivity(idc, idr, idv))
				'Q3 = Q3 + dblViewFactor * Qn(3, 1073, dblEmissivity, m_dblDeltaX(idc, idr, idv), m_dblTemps(idc, idr, idv, 1), GLASSemissivity(idc, idr, idv))
				'UPGRADE_WARNING: Couldn't resolve default property of object idv. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object idr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object idc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Q1a. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Q1a = Q1a + dblViewFactor * Qn(1, 1073, dblEmissivity, m_dblDeltaX(idc, idr, idv) / 2, m_dblTemps(idc, idr, idv, 1), GLASSemissivity(idc, idr, idv))
				'UPGRADE_WARNING: Couldn't resolve default property of object idv. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object idr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object idc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Q2a. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Q2a = Q2a + dblViewFactor * Qn(2, 1073, dblEmissivity, m_dblDeltaX(idc, idr, idv) / 2, m_dblTemps(idc, idr, idv, 1), GLASSemissivity(idc, idr, idv))
				'UPGRADE_WARNING: Couldn't resolve default property of object idv. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object idr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object idc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Q3a = Q3a + dblViewFactor * Qn(3, 1073, dblEmissivity, m_dblDeltaX(idc, idr, idv) / 2, m_dblTemps(idc, idr, idv, 1), GLASSemissivity(idc, idr, idv))
				
				
			End If
			
			
			'New temperature of the interior node
			'UPGRADE_WARNING: Couldn't resolve default property of object idv. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object idr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object idc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Q2a. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Q1a. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object dblHinterior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dblNewTemps(1) = m_dblTemps(idc, idr, idv, 1) + dblDeltaTime * (2 * dblHinterior * (dblTempUL - m_dblTemps(idc, idr, idv, 1)) / (m_dblRhocp(idc, idr, idv) * m_dblDeltaX(idc, idr, idv)) - 2 * dblKprime * (m_dblTemps(idc, idr, idv, 1) - m_dblTemps(idc, idr, idv, 2)) / (m_dblRhocp(idc, idr, idv) * m_dblDeltaX(idc, idr, idv) ^ 2) + (Q1a + Q2a + Q3a) / m_dblRhocp(idc, idr, idv) / m_dblDeltaX(idc, idr, idv) / 2)
			
			'New temperatures of all of the middle nodes
			For iCounter = 2 To (n - 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object idv. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object idr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object idc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Q2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Q1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				dblNewTemps(iCounter) = m_dblTemps(idc, idr, idv, iCounter) + dblDeltaTime * (dblKprime / m_dblRhocp(idc, idr, idv) * (m_dblTemps(idc, idr, idv, iCounter + 1) - 2 * m_dblTemps(idc, idr, idv, iCounter) + m_dblTemps(idc, idr, idv, iCounter - 1)) / m_dblDeltaX(idc, idr, idv) ^ 2 + (Q1 + Q2) / m_dblRhocp(idc, idr, idv) / m_dblDeltaX(idc, idr, idv))
			Next iCounter
			'If stepcount = 260 And idv = 4 Then Stop
			'UPGRADE_WARNING: Couldn't resolve default property of object idv. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object idr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object idc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			qinfty = GLASSemissivity(idc, idr, idv) * StefanBoltzmann * (m_dblTemps(idc, idr, idv, n) ^ 4 - InteriorTemp ^ 4)
			
			'New temperature of exterior node
			'UPGRADE_WARNING: Couldn't resolve default property of object idv. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object idr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object idc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Q2a. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Q1a. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dblNewTemps(n) = m_dblTemps(idc, idr, idv, n) + dblDeltaTime * (-2 * m_dblHexterior * (m_dblTemps(idc, idr, idv, n) - ExteriorTemp) / (m_dblRhocp(idc, idr, idv) * m_dblDeltaX(idc, idr, idv)) + 2 * dblKprime * (m_dblTemps(idc, idr, idv, n - 1) - m_dblTemps(idc, idr, idv, n)) / (m_dblRhocp(idc, idr, idv) * m_dblDeltaX(idc, idr, idv) ^ 2) + (Q1a + Q2a + qinfty) / m_dblRhocp(idc, idr, idv) / m_dblDeltaX(idc, idr, idv) / 2)
			
			For iCounter = 1 To n
				'UPGRADE_WARNING: Couldn't resolve default property of object idv. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object idr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object idc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				m_dblTemps(idc, idr, idv, iCounter) = dblNewTemps(iCounter)
			Next iCounter
			
		Next iIncrement
		
	End Sub
	
	Public Function Qn(ByRef intRange As Short, ByRef dblTemp As Double, ByRef dblEmissivity As Double, ByVal dblLength As Double, ByRef surfacetemp As Double, ByRef glassemm As Single) As Double
		
		'modified by C Wade 29/10/2002
		'=======================================================================
		'Ross Parry, 11/2001 - Canterbury University ME(Fire) Research Project
		'=======================================================================
		'
		'Calculates radiative heat fluxes from the hot upper layer based on three
		'wavelength bands, q1 : 0    <= lambda <= 2.75 micrometres
		'                  q2 : 2.75 <  lambda <= 4.5  micrometres
		'                  q3 : 4.5  <  lambda
		'========================================================================
		Const SIGMA As Double = 0.0000000567051 'Stefan-Boltzmann (W/(m2K4)
		Const C2 As Double = 14387.69 '(micrometres*kelvin)
		Const LAMBDA1 As Double = 2.75
		Const LAMBDA2 As Double = 4.5
		Const CONVERGESIZE As Double = 0.00000001
		Const MAXITS As Short = 20
		Const RHOBAR As Double = 0.057 'Average reflectivity
		Const GAMMA1 As Double = 35 'Absorption coefficients
		Const gamma2 As Double = 475 '
		
		Dim iCount As Short
		Dim dblZeta(2) As Double
		Dim dblFraction(3) As Double
		Dim dblNew As Double
		Dim iRange As Short
		Dim dblG0 As Double 'Non-Reflected Irradiation
		
		'Calculate dimensionless parameters, zeta
		dblZeta(1) = C2 / (LAMBDA1 * dblTemp)
		dblZeta(2) = C2 / (LAMBDA2 * dblTemp)
		
		'Calculate radiative energy fractions within each
		'wavelength band using converging series formula from Siegel and Howell
		For iRange = 1 To 2
			iCount = 1
			Do 
				dblNew = System.Math.Exp(-iCount * dblZeta(iRange)) / iCount * (dblZeta(iRange) ^ 3 + 3 * dblZeta(iRange) ^ 2 / iCount ^ 2 + 6 * dblZeta(iRange) / iCount ^ 2 + 6 / iCount ^ 3)
				If System.Math.Abs(dblNew) < CONVERGESIZE Then
					Exit Do
				Else : dblFraction(iRange) = dblFraction(iRange) + dblNew
					iCount = iCount + 1
				End If
			Loop While iCount < MAXITS
			'dblFraction(iRange) = dblFraction(iRange) * 15 / PI
			dblFraction(iRange) = dblFraction(iRange) * 15 / (PI ^ 4) 'cw
		Next iRange
		dblFraction(2) = dblFraction(2) - dblFraction(1)
		dblFraction(3) = 1 - dblFraction(2) - dblFraction(1)
		
		dblG0 = dblEmissivity * (1 - RHOBAR) * SIGMA * dblTemp ^ 4 'Non-reflected irradiation
		dblLength = 1.077 * dblLength 'Path length through element
		
		'Calculate radiant energy flow in each wavelength band
		Select Case intRange
			Case 1
				'Calculate q1
				Qn = (1 - System.Math.Exp(-GAMMA1 * dblLength)) * dblFraction(1) * dblG0
			Case 2
				'Calculate q2
				Qn = (1 - System.Math.Exp(-gamma2 * dblLength)) * dblFraction(2) * dblG0
			Case 3
				'Calculate q3
				'Q = dblFraction(3) * dblG0
				Qn = glassemm * (dblFraction(3) * dblG0 - SIGMA * surfacetemp ^ 4) 'cw
			Case Else
				Qn = 0 'Error
		End Select
	End Function
End Module