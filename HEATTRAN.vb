Option Strict Off
Option Explicit On
Imports System.Math

Module HEATTRAN

    Dim result1 As Integer = 0
    Dim result2 As Integer = 0

    Sub Areas(ByVal room As Integer, ByVal Z As Double, ByRef A1 As Double, ByRef A2 As Double, ByRef A3 As Double, ByRef A4 As Double)
		'************************************************************
		'*  Procedure to calculate room surface areas.
		'*
		'*  Z = layer interface height above floor
		'*  Called by: MatrixC(), Calc_Configs
		'*
		'*  Revised: 16 November 1995 Colleen Wade
		'************************************************************
		
		'Ceiling Area
		A1 = RoomLength(room) * RoomWidth(room)
		
		'Upper Wall Area
		A2 = (RoomHeight(room) - Z) * 2 * (RoomLength(room) + RoomWidth(room))
		If A2 < 0.0000000001 Then A2 = 0
		'Lower Wall Area
		A3 = Z * 2 * (RoomLength(room) + RoomWidth(room))
		
		'Floor Area
		A4 = A1
		
	End Sub
	
	Function Calc_Angle(ByVal X As Double, ByVal Y As Double, ByVal R As Double) As Double
		'************************************************************
		'*  Procedure calculates the solid angles between a point source
		'*  fire and the room surfaces assuming that the fire is located
		'*  in the centre of the room on the floor.
		'*
		'*  Called by: wAngle
		'*
		'*  Revised by Colleen Wade 5/7/95
		' 1 feb 2008 Amanda
		'************************************************************
		
		Dim D, b, A, C, E As Double
		On Error GoTo anglehandler
		
		A = 1 + R * R / (X * X + Y * Y)
        b = Y * Sqrt(A / (Y * Y + R * R))
        C = X * Sqrt(A / (X * X + R + R))
		
		If (-C * C + 1) < gcd_Machine_Error Then GoTo anglehandler
        If Abs((C - 1) * (b - 1)) < gcd_Machine_Error Then GoTo anglehandler
		
        D = Atan(b / Sqrt(-b * b + 1)) 'result in radians
        E = Atan(C / Sqrt(-C * C + 1)) 'result in radians
		Calc_Angle = (D + E - PI / 2) / (4 * PI)
		Exit Function
		
anglehandler: 
		Calc_Angle = 0
	End Function
	
    Sub Calc_Configs(ByVal room As Integer, ByVal Z As Double, ByVal A1 As Double, ByVal A2 As Double, ByVal A3 As Double, ByVal A4 As Double)
        '************************************************************
        '*  Procedure to calculate the configuration factors for the
        '*  room. Called at each time step.
        '*
        '*  Z = layer interface height
        '*  Called by: MatrixAB
        '*  Calls: ViewFactor_Parallel
        '*
        '*  Revised by Colleen Wade 5/7/95
        '************************************************************

        Dim F4d, F1d As Double

        'get view factor for floor to interface
        F4d = ViewFactor_Parallel(room, Z)

        'get view factor for ceiling to interface
        F1d = ViewFactor_Parallel(room, RoomHeight(room) - Z)

        'View Factors
        'F(1,4) previously calculated
        F(1, 1, room) = 0
        F(4, 4, room) = 0
        F(1, 2, room) = 1 - F1d

        If A2 <> 0 Then
            F(2, 1, room) = A1 / A2 * F(1, 2, room)
        Else
            F(2, 1, room) = 0
        End If

        F(2, 2, room) = 1 - 2 * F(2, 1, room)
        F(4, 3, room) = 1 - F4d

        If A3 <> 0 Then
            F(3, 4, room) = A4 / A3 * F(4, 3, room)
        Else
            F(3, 4, room) = 0
        End If

        F(3, 3, room) = 1 - 2 * F(3, 4, room)
        F(1, 3, room) = 1 - F(1, 4, room) - F(1, 2, room)

        If A3 <> 0 Then
            F(3, 1, room) = A1 / A3 * F(1, 3, room)
        Else
            F(3, 1, room) = 0
        End If

        F(3, 2, room) = 1 - F(3, 1, room) - F(3, 3, room) - F(3, 4, room)

        If A2 <> 0 Then
            F(2, 3, room) = A3 / A2 * F(3, 2, room)
        Else
            F(2, 3, room) = 0
        End If

        F(2, 4, room) = 1 - F(2, 1, room) - F(2, 2, room) - F(2, 3, room)
        F(4, 2, room) = A2 / A4 * F(2, 4, room)
        F(4, 1, room) = F(1, 4, room)

    End Sub
	
	Function co2emm(ByVal temp As Double, ByVal PL As Double) As Double
		'************************************************************
		'*  This function returns the emissivity of carbon dioxide
		'*  in a mixture with non-radiating gases at 1 atm total pressure
		'*  and of hemispherical shape.
		'*
		'*  temp = gas temperature
		'*  PL = partial pressure of co2 x mean beam length
		'*
		'*  Revised: 30 November 1995 Colleen Wade
		'************************************************************
		
		If PL >= 1.2 Then
			If temp <= 350 Then
				co2emm = 0.2 : Exit Function
			ElseIf temp <= 450 Then 
				co2emm = 0.18 : Exit Function
			ElseIf temp <= 550 Then 
				co2emm = 0.17 : Exit Function
			ElseIf temp <= 650 Then 
				co2emm = 0.18 : Exit Function
			ElseIf temp <= 750 Then 
				co2emm = 0.19 : Exit Function
			ElseIf temp <= 850 Then 
				co2emm = 0.2 : Exit Function
			ElseIf temp > 850 Then 
				co2emm = 0.21 : Exit Function
			End If
		ElseIf PL >= 0.61 Then 
			If temp <= 350 Then
				co2emm = 0.17 : Exit Function
			ElseIf temp <= 450 Then 
				co2emm = 0.16 : Exit Function
			ElseIf temp <= 550 Then 
				co2emm = 0.15 : Exit Function
			ElseIf temp <= 750 Then 
				co2emm = 0.16 : Exit Function
			ElseIf temp > 750 Then 
				co2emm = 0.17 : Exit Function
			End If
		ElseIf PL >= 0.3 Then 
			If temp <= 350 Then
				co2emm = 0.15 : Exit Function
			ElseIf temp <= 450 Then 
				co2emm = 0.14 : Exit Function
			ElseIf temp <= 550 Then 
				co2emm = 0.13 : Exit Function
			ElseIf temp <= 850 Then 
				co2emm = 0.14 : Exit Function
			ElseIf temp > 850 Then 
				co2emm = 0.15 : Exit Function
			End If
		ElseIf PL >= 0.24 Then 
			If temp <= 350 Then
				co2emm = 0.14 : Exit Function
			ElseIf temp <= 450 Then 
				co2emm = 0.13 : Exit Function
			ElseIf temp <= 550 Then 
				co2emm = 0.12 : Exit Function
			ElseIf temp <= 650 Then 
				co2emm = 0.13 : Exit Function
			ElseIf temp <= 750 Then 
				co2emm = 0.14 : Exit Function
			ElseIf temp <= 850 Then 
				co2emm = 0.13 : Exit Function
			ElseIf temp > 850 Then 
				co2emm = 0.14 : Exit Function
			End If
		ElseIf PL >= 0.12 Then 
			If temp <= 350 Then
				co2emm = 0.12 : Exit Function
			ElseIf temp <= 450 Then 
				co2emm = 0.11 : Exit Function
			ElseIf temp <= 650 Then 
				co2emm = 0.1 : Exit Function
			ElseIf temp <= 750 Then 
				co2emm = 0.11 : Exit Function
			ElseIf temp <= 850 Then 
				co2emm = 0.12 : Exit Function
			ElseIf temp > 850 Then 
				co2emm = 0.13 : Exit Function
			End If
		ElseIf PL >= 0.061 Then 
			If temp <= 350 Then
				co2emm = 0.095 : Exit Function
			ElseIf temp <= 450 Then 
				co2emm = 0.085 : Exit Function
			ElseIf temp <= 550 Then 
				co2emm = 0.084 : Exit Function
			ElseIf temp <= 650 Then 
				co2emm = 0.085 : Exit Function
			ElseIf temp <= 750 Then 
				co2emm = 0.09 : Exit Function
			ElseIf temp <= 850 Then 
				co2emm = 0.095 : Exit Function
			ElseIf temp > 850 Then 
				co2emm = 0.099 : Exit Function
			End If
		ElseIf PL >= 0.03 Then 
			If temp <= 350 Then
				co2emm = 0.08 : Exit Function
			ElseIf temp <= 450 Then 
				co2emm = 0.07 : Exit Function
			ElseIf temp <= 550 Then 
				co2emm = 0.069 : Exit Function
			ElseIf temp <= 650 Then 
				co2emm = 0.07 : Exit Function
			ElseIf temp <= 750 Then 
				co2emm = 0.075 : Exit Function
			ElseIf temp > 750 Then 
				co2emm = 0.08 : Exit Function
			End If
		ElseIf PL >= 0.018 Then 
			If temp <= 350 Then
				co2emm = 0.066 : Exit Function
			ElseIf temp <= 450 Then 
				co2emm = 0.06 : Exit Function
			ElseIf temp <= 550 Then 
				co2emm = 0.059 : Exit Function
			ElseIf temp <= 650 Then 
				co2emm = 0.06 : Exit Function
			ElseIf temp <= 750 Then 
				co2emm = 0.064 : Exit Function
			ElseIf temp > 750 Then 
				co2emm = 0.069 : Exit Function
			End If
		ElseIf PL >= 0.012 Then 
			If temp <= 350 Then
				co2emm = 0.058 : Exit Function
			ElseIf temp <= 550 Then 
				co2emm = 0.05 : Exit Function
			ElseIf temp <= 650 Then 
				co2emm = 0.051 : Exit Function
			ElseIf temp <= 750 Then 
				co2emm = 0.055 : Exit Function
			ElseIf temp <= 850 Then 
				co2emm = 0.077 : Exit Function
			ElseIf temp > 850 Then 
				co2emm = 0.059 : Exit Function
			End If
		ElseIf PL >= 0.0061 Then 
			If temp <= 350 Then
				co2emm = 0.05 : Exit Function
			ElseIf temp <= 450 Then 
				co2emm = 0.045 : Exit Function
			ElseIf temp <= 550 Then 
				co2emm = 0.04 : Exit Function
			ElseIf temp <= 650 Then 
				co2emm = 0.041 : Exit Function
			ElseIf temp <= 750 Then 
				co2emm = 0.044 : Exit Function
			ElseIf temp > 750 Then 
				co2emm = 0.045 : Exit Function
			End If
		ElseIf PL >= 0.003 Then 
			If temp <= 350 Then
				co2emm = 0.03 : Exit Function
			ElseIf temp <= 550 Then 
				co2emm = 0.028 : Exit Function
			ElseIf temp <= 650 Then 
				co2emm = 0.029 : Exit Function
			ElseIf temp > 650 Then 
				co2emm = 0.03 : Exit Function
			End If
		ElseIf PL >= 0.0018 Then 
			If temp <= 350 Then
				co2emm = 0.026 : Exit Function
			ElseIf temp <= 450 Then 
				co2emm = 0.024 : Exit Function
			ElseIf temp <= 550 Then 
				co2emm = 0.022 : Exit Function
			ElseIf temp > 550 Then 
				co2emm = 0.023 : Exit Function
			End If
		ElseIf PL >= 0.0015 Then 
			If temp <= 550 Then
				co2emm = 0.022 : Exit Function
			ElseIf temp <= 750 Then 
				co2emm = 0.021 : Exit Function
			ElseIf temp <= 850 Then 
				co2emm = 0.02 : Exit Function
			ElseIf temp > 850 Then 
				co2emm = 0.019 : Exit Function
			End If
		ElseIf PL >= 0.0012 Then 
			If temp <= 750 Then
				co2emm = 0.018 : Exit Function
			ElseIf temp <= 850 Then 
				co2emm = 0.017 : Exit Function
			ElseIf temp > 850 Then 
				co2emm = 0.016 : Exit Function
			End If
		ElseIf PL >= 0.0009 Then 
			If temp <= 550 Then
				co2emm = 0.015 : Exit Function
			ElseIf temp <= 850 Then 
				co2emm = 0.014 : Exit Function
			ElseIf temp > 850 Then 
				co2emm = 0.013 : Exit Function
			End If
		ElseIf PL >= 0.0006 Then 
			If temp <= 750 Then
				co2emm = 0.01 : Exit Function
			ElseIf temp <= 850 Then 
				co2emm = 0.095 : Exit Function
			ElseIf temp > 850 Then 
				co2emm = 0.009 : Exit Function
			End If
		Else
			If temp <= 650 Then
				co2emm = 0.006 : Exit Function
			ElseIf temp <= 750 Then 
				co2emm = 0.005 : Exit Function
			ElseIf temp <= 850 Then 
				co2emm = 0.0045 : Exit Function
			ElseIf temp > 850 Then 
				co2emm = 0.0038 : Exit Function
			End If
		End If
		
	End Function
	
    Sub FourWallRad(ByVal room As Integer, ByVal Q As Double, ByVal CeilingTemp As Double, ByVal Upperwalltemp As Double, ByVal LowerWallTemp As Double, ByVal FloorTemp As Double, ByVal uppertemp As Double, ByVal lowertemp As Double, ByVal Z As Double, ByRef prd(,) As Double, ByRef A1 As Double, ByRef A2 As Double, ByRef A3 As Double, ByRef A4 As Double, ByRef matc(,) As Double, ByVal YCO2 As Double, ByVal YH2O As Double, ByVal YCO2Lower As Double, ByVal YH2Olower As Double, ByVal volume As Double)
        '************************************************************
        '*  Procedure to calculate the radiant heat fluxes to the four
        '*  surfaces in the room: ceiling, upper walls, lower walls,
        '*  and floor.
        '*
        '*  Z = position of the layer interface above floor
        '*  ceilingtemp = temperature of ceiling
        '*  upperwalltemp = temperature of the upper wall
        '*  lowerwalltemp = temperature of the lower wall
        '*  floortemp = temperature of the floor
        '*
        '*  Subprocedures and functions called:
        '*      MatrixAB, MatrixC, MatrixE, Matsub, Matprd, Matsol
        '*      MatrixD
        '*
        '*  Revised: 16 November 1995 Colleen Wade
        '************************************************************

        Dim MatA(4, 4) As Double
        Dim MatB(4, 4) As Double
        Dim MatE(4, 1) As Double
        Dim R(4, 1) As Double
        Dim MatD(4, 4) As Double
        'ReDim MatA(4, 4)
        'ReDim MatB(4, 4)
        'ReDim MatE(4, 1)
        'ReDim R(4, 1)
        'ReDim MatD(4, 4)

        'room surface areas (including all vents)
        Call Areas(room, Z, A1, A2, A3, A4)

        'define elements in MatA() and MatB()
        Call MatrixAB(room, MatA, MatB, Z, A1, A2, A3, A4, uppertemp, lowertemp, YCO2, YH2O, YCO2Lower, YH2Olower, volume)

        'define elements in MatC() 4 x 1 - point source fire  + gas layers
        Call MatrixC(room, Q, Z, A1, A2, A3, A4, matc, uppertemp, lowertemp)

        'define elements in MatE() - black body emissive powers
        Call MatrixE(CeilingTemp, Upperwalltemp, LowerWallTemp, FloorTemp, MatE)

        'multiply matrices MatB() x MatE() = prd()
        Call Matprd(MatB, MatE, prd, 4, 4, 1)

        'subtract matrix c  R()=prd()-matC()
        Call Matsub(prd, matc, R, 4, 1)

        'call procedure to solve simultaneous linear equations
        'solution in R()
        Call MatSol(MatA, R, 4)

        'define elements in MatD - contains surface emissivities
        Call MatrixD(MatD, room)

        'get net surface radiant heat fluxes
        'result stored in prd()   4 x 1 matrix
        Call Matprd(MatD, R, prd, 4, 4, 1)
        System.Windows.Forms.Application.DoEvents()
    End Sub
	
	Function gas_emissivity_lower(ByVal room As Integer, ByVal Z As Double, ByVal lowertemp As Double, ByVal YCO2Lower As Double, ByVal YH2Olower As Double, ByVal volume As Double) As Double
		'*  ===================================================================
		'*  This function returns the emissivity of the lower layer. assuming
		'*
		'*  Arguments passed to the function:
		'*      Z = height of the smoke layer interface above the floor (m)
		'*
		'*  Global variables used:
		'*      RoomLength
		'*      Roomwidth
		'*      EmissionCoefficent
		'*
		'*  Revised: 13/1/98 Colleen Wade
		'*  ===================================================================
		
        Dim SurfaceArea, PathLength As Double
		Dim eH20, Pc, Pw, PLW, PLC, eCO2 As Double
		Dim YCO2, GE As Double
		
		'find lower gas layer volume
		'volume = RoomLength(room) * RoomWidth(room) * Z
		volume = RoomVolume(room) - volume
		
		'find lower gas layer surface area
		SurfaceArea = 2 * (RoomFloorArea(room) + Z * (RoomLength(room) + RoomWidth(room)))
		
		'find geometric beam length (correction factor =0.9)
		PathLength = 0.9 * 4 * volume / SurfaceArea
		
		'partial pressure of Water Vapor in atm
		'Pw = H2OMassFraction(1, 1) * ReferenceDensity * ReferenceTemp * Gas_Constant / (MolecularWeightH2O * 101.3)
		'Pw = YH2Olower * ReferenceDensity * ReferenceTemp * Gas_Constant / (MolecularWeightH2O * Atm_Pressure)
		Pw = YH2Olower * (Gas_Constant / Gas_Constant_Air) / (MolecularWeightH2O)
		
		'PLW = Pw * PathLength * lowertemp / InteriorTemp  'atm-m
		PLW = Pw * PathLength 'atm-m
		
		'emissivity of water vapor
		eH20 = H2Oemm(lowertemp, PLW)
		
		'mass fraction of CO2 in ambient air
		'YCO2 = 0.033 * MolecularWeightCO2 / MolecularWeightAir
		
		'partial pressure of Carbon Dioxide in atm
		'Pc = YCO2 * ReferenceDensity * ReferenceTemp * Gas_Constant / (MolecularWeightCO2 * Atm_Pressure)
        'Pc = YCO2 * (Gas_Constant / Gas_Constant_Air) / (MolecularWeightCO2)
        Pc = YCO2Lower * (Gas_Constant / Gas_Constant_Air) / (MolecularWeightCO2)

		PLC = Pc * PathLength 'atm-m
		
		'emissivity of carbon dioxide
		eCO2 = co2emm(lowertemp, PLC)
		
		'find the gas emissivity
		GE = eH20 + 0.5 * eCO2
		
		'find the gas + soot emissivity
        gas_emissivity_lower = (1 - Exp(-2.3 * OD_lower(room, stepcount - 1) * PathLength)) + GE * Exp(-2.3 * OD_lower(room, stepcount - 1) * PathLength)

    End Function
	
	Function gas_emissivity_upper(ByVal room As Integer, ByVal volume As Double, ByVal Z As Double, ByVal uppertemp As Double, ByVal YCO2 As Double, ByVal YH2O As Double, ByVal OD As Double) As Double
		'*  ===================================================================
		'*  This function returns the emissivity of the upper layer assuming
		'*  only soot, H20 and CO2 are emitting. Reference is pg 1-101 of SFPE Handbook.
		'*
		'*  Arguments passed to the function:
		'*      Z = height of the smoke layer interface above the floor (m)
		'*
		'*  Global variables used:
		'*      RoomLength
		'*      Roomwidth
		'*      EmissionCoefficent
		'*
		'*  Revised: 29 November 1996 Colleen Wade
		'*  ===================================================================
		
		Dim Pw, SurfaceArea, PathLength, Pc As Double
		Dim PLC, b, A, GE, PLW As Double
		Dim eH20, eCO2 As Double
		
		On Error GoTo errorhandler
		
		'find upper gas layer surface area
		If volume > 0 Then
			'total surface area of upper layer
			SurfaceArea = 2 * (RoomFloorArea(room) + (RoomHeight(room) - Z) * (RoomLength(room) + RoomWidth(room)))
		Else
			SurfaceArea = 0
			Pw = YH2O * (Gas_Constant / Gas_Constant_Air) / (MolecularWeightH2O)
			PLW = 0
			eH20 = H2Oemm(uppertemp, PLW)
			Pc = YCO2 * (Gas_Constant / Gas_Constant_Air) / (MolecularWeightCO2)
			PLC = 0
			eCO2 = co2emm(uppertemp, PLC)
			GE = eH20 + 0.5 * eCO2
			gas_emissivity_upper = GE
			Exit Function
		End If
		
		'find geometric beam length (correction factor =0.9)
		PathLength = 0.9 * 4 * volume / SurfaceArea 'm
		
		'partial pressure of Water Vapor at ambient in atm
		Pw = YH2O * (Gas_Constant / Gas_Constant_Air) / (MolecularWeightH2O)
		
		PLW = Pw * PathLength 'atm-m
		
		'emissivity of water vapor
		eH20 = H2Oemm(uppertemp, PLW)
		
		'partial pressure of Carbon Dioxide in atm
		Pc = YCO2 * (Gas_Constant / Gas_Constant_Air) / (MolecularWeightCO2)
		PLC = Pc * PathLength 'atm-m
		
		'emissivity of carbon dioxide
		eCO2 = co2emm(uppertemp, PLC)
		
		'find the gas emissivity
		GE = eH20 + 0.5 * eCO2

		'find the gas + soot emissivity
		'gas_emissivity_upper = (1 - Exp(-EmissionCoefficient * PathLength)) + GE * Exp(-EmissionCoefficient * PathLength)
		'If stepcount > 1 And flagstop = 0 Then
		If stepcount > 1 Then
			'gas_emissivity_upper = (1 - Exp(-2.3 * OD_upper(room, stepcount - 1) * PathLength)) + GE * Exp(-2.3 * OD_upper(room, stepcount - 1) * PathLength)
            gas_emissivity_upper = (1 - Exp(-2.3 * OD * PathLength)) + GE * Exp(-2.3 * OD * PathLength)

		End If
		Exit Function
errorhandler: 
		'MsgBox Error(Err) + " in Function gas_emissivity_upper in Module HEATTRAN.BAS"
		Resume Next
	End Function
	
    Sub Get_Vent_Area(ByVal room As Integer, ByVal tim As Double, ByRef VentAreaUpper As Double, ByRef VentAreaLower As Double, ByVal Z As Double)
        '**********************************************************************
        '*  Return the total open vent area.
        '*
        '*  Revised: Colleen Wade 8 May 1997
        '*********************************************************************
        '%%% needs correcting for gradual opening of vents

        Dim i, j, k As Integer
        Dim area As Double

        'total vent area above layer height
        VentAreaUpper = 0

        'total vent area below layer height
        VentAreaLower = 0
        i = room
        'For i = 1 To NumberRooms
        For k = 1 To NumberRooms + 1
            'If i < k Then
            If NumberVents(i, k) <> 0 And k <> i Then
                For j = 1 To NumberVents(i, k)
                    If tim < VentOpenTime(i, k, j) Or (tim > ventclosetime(i, k, j) And ventclosetime(i, k, j) > VentOpenTime(i, k, j)) Then
                        'vent is closed
                    Else
                        'vent is open
                        area = ventarea(tim, i, k, j)

                        If Z >= soffitheight(i, k, j) Then
                            'VentAreaLower = VentAreaLower + VentWidth(i, k, j) * VentHeight(i, k, j)
                            VentAreaLower = VentAreaLower + area
                        ElseIf Z < soffitheight(i, k, j) And Z > VentSillHeight(i, k, j) Then
                            'VentAreaUpper = VentAreaUpper + VentWidth(i, k, j) * (soffitheight(i, k, j) - Z)
                            'VentAreaLower = VentAreaLower + VentWidth(i, k, j) * (Z - VentSillHeight(i, k, j))
                            VentAreaUpper = VentAreaUpper + area / VentHeight(i, k, j) * (soffitheight(i, k, j) - Z)
                            VentAreaLower = VentAreaLower + area / VentHeight(i, k, j) * (Z - VentSillHeight(i, k, j))
                        Else
                            'VentAreaUpper = VentAreaUpper + VentWidth(i, k, j) * VentHeight(i, k, j)
                            VentAreaUpper = VentAreaUpper + area
                        End If

                    End If
                Next j
            End If
            'ElseIf k < i Then
            '    If NumberVents(k, i) <> 0 And k <> i Then
            '        For j = 1 To NumberVents(k, i)
            '            If tim < VentOpenTime(k, i, j) Or (tim > VentCloseTime(k, i, j) And VentCloseTime(k, i, j) > VentOpenTime(k, i, j)) Then
            '            'vent is closed
            '            Else
            '                'vent is open
            '                If Z >= soffitheight(k, i, j) Then
            '                    VentAreaLower = VentAreaLower + VentWidth(k, i, j) * VentHeight(k, i, j)
            '                ElseIf Z < soffitheight(k, i, j) And Z > VentSillHeight(k, i, j) Then
            '                    ventareaupper = ventareaupper + VentWidth(k, i, j) * (soffitheight(k, i, j) - Z)
            '                    VentAreaLower = VentAreaLower + VentWidth(k, i, j) * (Z - VentSillHeight(k, i, j))
            '                Else
            '                    ventareaupper = ventareaupper + VentWidth(k, i, j) * VentHeight(k, i, j)
            '                End If
            '            End If
            '        Next j
            '    End If
            'End If
        Next k
        'Next i
    End Sub
	
	Function H2Oemm(ByVal temp As Double, ByVal PL As Double) As Double
		'************************************************************
		'*  This function returns the emissivity of water vapour
		'*  in a mixture with non-radiating gases at 1 atm total pressure
		'*  and of hemispherical shape.
		'*
		'*  temp = gas temperature
		'*  PL = partial pressure of H2O x mean beam length
		'*
		'*  Revised: 30 November 1995 Colleen Wade
		'************************************************************
		
		On Error GoTo h2ohandler
		
		If PL >= 6.1 Then
			If temp <= 350 Then
				H2Oemm = 0.66 : Exit Function
			ElseIf temp <= 450 Then 
				H2Oemm = 0.62 : Exit Function
			ElseIf temp > 450 Then 
				H2Oemm = 0.6 : Exit Function
			End If
		ElseIf PL >= 3 Then 
			If temp <= 350 Then
				H2Oemm = 0.59 : Exit Function
			ElseIf temp <= 450 Then 
				H2Oemm = 0.55 : Exit Function
			ElseIf temp <= 750 Then 
				H2Oemm = 0.54 : Exit Function
			ElseIf temp <= 850 Then 
				H2Oemm = 0.53 : Exit Function
			ElseIf temp > 850 Then 
				H2Oemm = 0.5 : Exit Function
			End If
		ElseIf PL >= 1.5 Then 
			If temp <= 350 Then
				H2Oemm = 0.52 : Exit Function
			ElseIf temp <= 450 Then 
				H2Oemm = 0.49 : Exit Function
			ElseIf temp <= 550 Then 
				H2Oemm = 0.47 : Exit Function
			ElseIf temp <= 750 Then 
				H2Oemm = 0.46 : Exit Function
			ElseIf temp <= 850 Then 
				H2Oemm = 0.45 : Exit Function
			ElseIf temp > 850 Then 
				H2Oemm = 0.42 : Exit Function
			End If
		ElseIf PL >= 0.9 Then 
			If temp <= 350 Then
				H2Oemm = 0.45 : Exit Function
			ElseIf temp <= 450 Then 
				H2Oemm = 0.43 : Exit Function
			ElseIf temp <= 650 Then 
				H2Oemm = 0.4 : Exit Function
			ElseIf temp <= 750 Then 
				H2Oemm = 0.39 : Exit Function
			ElseIf temp <= 850 Then 
				H2Oemm = 0.37 : Exit Function
			ElseIf temp > 850 Then 
				H2Oemm = 0.36 : Exit Function
			End If
		ElseIf PL >= 0.61 Then 
			If temp <= 350 Then
				H2Oemm = 0.41 : Exit Function
			ElseIf temp <= 450 Then 
				H2Oemm = 0.38 : Exit Function
			ElseIf temp <= 550 Then 
				H2Oemm = 0.36 : Exit Function
			ElseIf temp <= 650 Then 
				H2Oemm = 0.35 : Exit Function
			ElseIf temp <= 750 Then 
				H2Oemm = 0.34 : Exit Function
			ElseIf temp <= 850 Then 
				H2Oemm = 0.32 : Exit Function
			ElseIf temp > 850 Then 
				H2Oemm = 0.31 : Exit Function
			End If
		ElseIf PL >= 0.3 Then 
			If temp <= 350 Then
				H2Oemm = 0.34 : Exit Function
			ElseIf temp <= 450 Then 
				H2Oemm = 0.3 : Exit Function
			ElseIf temp <= 550 Then 
				H2Oemm = 0.28 : Exit Function
			ElseIf temp <= 650 Then 
				H2Oemm = 0.27 : Exit Function
			ElseIf temp <= 750 Then 
				H2Oemm = 0.26 : Exit Function
			ElseIf temp <= 850 Then 
				H2Oemm = 0.25 : Exit Function
			ElseIf temp > 850 Then 
				H2Oemm = 0.24 : Exit Function
			End If
		ElseIf PL >= 0.18 Then 
			If temp <= 350 Then
				H2Oemm = 0.27 : Exit Function
			ElseIf temp <= 450 Then 
				H2Oemm = 0.25 : Exit Function
			ElseIf temp <= 550 Then 
				H2Oemm = 0.24 : Exit Function
			ElseIf temp <= 650 Then 
				H2Oemm = 0.23 : Exit Function
			ElseIf temp <= 750 Then 
				H2Oemm = 0.22 : Exit Function
			ElseIf temp <= 850 Then 
				H2Oemm = 0.2 : Exit Function
			ElseIf temp > 850 Then 
				H2Oemm = 0.18 : Exit Function
			End If
		ElseIf PL >= 0.12 Then 
			If temp <= 350 Then
				H2Oemm = 0.23 : Exit Function
			ElseIf temp <= 450 Then 
				H2Oemm = 0.21 : Exit Function
			ElseIf temp <= 650 Then 
				H2Oemm = 0.19 : Exit Function
			ElseIf temp <= 750 Then 
				H2Oemm = 0.18 : Exit Function
			ElseIf temp <= 850 Then 
				H2Oemm = 0.17 : Exit Function
			ElseIf temp > 750 Then 
				H2Oemm = 0.16 : Exit Function
			End If
		ElseIf PL >= 0.061 Then 
			If temp <= 350 Then
				H2Oemm = 0.18 : Exit Function
			ElseIf temp <= 450 Then 
				H2Oemm = 0.16 : Exit Function
			ElseIf temp <= 650 Then 
				H2Oemm = 0.14 : Exit Function
			ElseIf temp <= 750 Then 
				H2Oemm = 0.13 : Exit Function
			ElseIf temp <= 850 Then 
				H2Oemm = 0.12 : Exit Function
			ElseIf temp > 850 Then 
				H2Oemm = 0.11 : Exit Function
			End If
		ElseIf PL >= 0.03 Then 
			If temp <= 350 Then
				H2Oemm = 0.13 : Exit Function
			ElseIf temp <= 450 Then 
				H2Oemm = 0.11 : Exit Function
			ElseIf temp <= 550 Then 
				H2Oemm = 0.095 : Exit Function
			ElseIf temp <= 650 Then 
				H2Oemm = 0.09 : Exit Function
			ElseIf temp <= 750 Then 
				H2Oemm = 0.085 : Exit Function
			ElseIf temp <= 850 Then 
				H2Oemm = 0.075 : Exit Function
			ElseIf temp > 850 Then 
				H2Oemm = 0.069 : Exit Function
			End If
		ElseIf PL >= 0.018 Then 
			If temp <= 350 Then
				H2Oemm = 0.096 : Exit Function
			ElseIf temp <= 450 Then 
				H2Oemm = 0.082 : Exit Function
			ElseIf temp <= 550 Then 
				H2Oemm = 0.07 : Exit Function
			ElseIf temp <= 650 Then 
				H2Oemm = 0.067 : Exit Function
			ElseIf temp <= 750 Then 
				H2Oemm = 0.06 : Exit Function
			ElseIf temp <= 850 Then 
				H2Oemm = 0.054 : Exit Function
			ElseIf temp > 850 Then 
				H2Oemm = 0.046 : Exit Function
			End If
		ElseIf PL >= 0.0012 Then 
			If temp <= 350 Then
				H2Oemm = 0.079 : Exit Function
			ElseIf temp <= 450 Then 
				H2Oemm = 0.068 : Exit Function
			ElseIf temp <= 550 Then 
				H2Oemm = 0.057 : Exit Function
			ElseIf temp <= 650 Then 
				H2Oemm = 0.051 : Exit Function
			ElseIf temp <= 750 Then 
				H2Oemm = 0.045 : Exit Function
			ElseIf temp <= 850 Then 
				H2Oemm = 0.04 : Exit Function
			ElseIf temp > 850 Then 
				H2Oemm = 0.035 : Exit Function
			End If
		ElseIf PL >= 0.0061 Then 
			If temp <= 350 Then
				H2Oemm = 0.048 : Exit Function
			ElseIf temp <= 450 Then 
				H2Oemm = 0.043 : Exit Function
			ElseIf temp <= 550 Then 
				H2Oemm = 0.036 : Exit Function
			ElseIf temp <= 650 Then 
				H2Oemm = 0.031 : Exit Function
			ElseIf temp <= 750 Then 
				H2Oemm = 0.027 : Exit Function
			ElseIf temp <= 850 Then 
				H2Oemm = 0.023 : Exit Function
			ElseIf temp > 850 Then 
				H2Oemm = 0.02 : Exit Function
			End If
		ElseIf PL >= 0.0045 Then 
			If temp <= 350 Then
				H2Oemm = 0.04 : Exit Function
			ElseIf temp <= 450 Then 
				H2Oemm = 0.035 : Exit Function
			ElseIf temp <= 550 Then 
				H2Oemm = 0.029 : Exit Function
			ElseIf temp <= 650 Then 
				H2Oemm = 0.025 : Exit Function
			ElseIf temp <= 750 Then 
				H2Oemm = 0.022 : Exit Function
			ElseIf temp <= 850 Then 
				H2Oemm = 0.018 : Exit Function
			ElseIf temp > 850 Then 
				H2Oemm = 0.016 : Exit Function
			End If
		ElseIf PL >= 0.003 Then 
			If temp <= 350 Then
				H2Oemm = 0.029 : Exit Function
			ElseIf temp <= 450 Then 
				H2Oemm = 0.025 : Exit Function
			ElseIf temp <= 550 Then 
				H2Oemm = 0.021 : Exit Function
			ElseIf temp <= 650 Then 
				H2Oemm = 0.018 : Exit Function
			ElseIf temp <= 750 Then 
				H2Oemm = 0.016 : Exit Function
			ElseIf temp <= 850 Then 
				H2Oemm = 0.013 : Exit Function
			ElseIf temp > 850 Then 
				H2Oemm = 0.012 : Exit Function
			End If
		ElseIf PL >= 0.0021 Then 
			If temp <= 350 Then
				H2Oemm = 0.023 : Exit Function
			ElseIf temp <= 450 Then 
				H2Oemm = 0.019 : Exit Function
			ElseIf temp <= 550 Then 
				H2Oemm = 0.016 : Exit Function
			ElseIf temp <= 650 Then 
				H2Oemm = 0.013 : Exit Function
			ElseIf temp <= 750 Then 
				H2Oemm = 0.012 : Exit Function
			ElseIf temp <= 850 Then 
				H2Oemm = 0.0095 : Exit Function
			ElseIf temp > 850 Then 
				H2Oemm = 0.008 : Exit Function
			End If
		Else
			If temp <= 350 Then
				H2Oemm = 0.017 : Exit Function
			ElseIf temp <= 450 Then 
				H2Oemm = 0.015 : Exit Function
			ElseIf temp <= 550 Then 
				H2Oemm = 0.013 : Exit Function
			ElseIf temp <= 650 Then 
				H2Oemm = 0.011 : Exit Function
			ElseIf temp <= 750 Then 
				H2Oemm = 0.009 : Exit Function
			ElseIf temp <= 850 Then 
				H2Oemm = 0.0075 : Exit Function
			ElseIf temp > 850 Then 
				H2Oemm = 0.007 : Exit Function
			End If
		End If
		
		Exit Function
		
h2ohandler: 
		MsgBox(ErrorToString(Err.Number) & " in H2Oemm")
		ERRNO = Err.Number
        'System.Windows.Forms.Cursor.Current = Default_Renamed
		Exit Function
		
	End Function
    Sub Implicit_Surface_Temps_furnace(ByVal qnet As Double, ByVal i As Integer, ByRef TempNode(,) As Double, ByVal Fourier As Double, ByVal nodes As Integer, ByVal DeltaX As Double, ByVal conductivity As Double, ByVal outsidebiot As Double, ByVal SurfaceEmissivity As Double)

        Dim k, j As Integer
        Dim UC(nodes, nodes) As Double
        Dim CX(nodes, 1) As Double
        Dim ambienttemp As Double = 293 'K


        'only one surface layer
        UC(1, 1) = 1 + 2 * Fourier
        UC(1, 2) = -2 * Fourier

        k = 2
        For j = 2 To nodes - 1
            UC(j, k - 1) = -Fourier
            UC(j, k) = 1 + 2 * Fourier
            UC(j, k + 1) = -Fourier
            k = k + 1
        Next j

        UC(nodes, nodes - 1) = -2 * Fourier
        UC(nodes, nodes) = 1 + 2 * Fourier

        'interior boundary conditions
        CX(1, 1) = -2 * qnet * 1000 * Fourier * DeltaX / conductivity + TempNode(1, i)

        For k = 2 To nodes - 1
            CX(k, 1) = TempNode(k, i)
        Next k

        'exterior boundary conditions
        CX(nodes, 1) = 2 * Fourier * outsidebiot * ((ambienttemp - TempNode(nodes, i)) - SurfaceEmissivity / OutsideConvCoeff * StefanBoltzmann * (TempNode(nodes, i) ^ 4 - ambienttemp ^ 4)) + TempNode(nodes, i)

        Dim ier As Short
        If frmOptions1.optLUdecom.Checked = True Then
            'find surface temperatures for the next timestep
            'using method of LU decomposition (preferred)
            Call MatSol(UC, CX, nodes)
        Else
            'find surface temperatures for the next timestep
            'using method of Gauss-Jordan elimination
            Call LINEAR2(nodes, UC, CX, ier)
            If ier = 1 Then MsgBox("singular matrix in implicit_surface_temps_furnace")
        End If

        For j = 1 To nodes
            TempNode(j, i + 1) = CX(j, 1)
        Next j

        Erase UC
        Erase CX

    End Sub
    Sub Implicit_Surface_Temps_furnace2(ByVal qnet As Double, ByVal i As Integer, ByRef TempNode(,) As Double, ByVal Fourier As Double, ByVal nodes As Integer, ByVal DeltaX As Double, ByVal conductivity As Double, ByVal subconductivity As Double, ByVal outsidebiot As Double, ByVal SurfaceEmissivity As Double, ByVal SubThickness As Double, ByVal SubSpecificHeat As Double, ByVal SubDensity As Double, ByVal SubEmissivity As Double)

        Dim k, j As Integer
        Dim UC(2 * nodes - 1, 2 * nodes - 1) As Double
        Dim CX(2 * nodes - 1, 1) As Double
        Dim ambienttemp As Double = 293 'K
        Dim OutsideBiot2, Fourier2 As Double

        Fourier2 = subconductivity * Timestep / (((SubThickness / 1000) / (nodes - 1)) ^ 2 * SubDensity * SubSpecificHeat)
        OutsideBiot2 = OutsideConvCoeff * ((SubThickness / 1000) / (nodes - 1)) / subconductivity

        UC(1, 1) = 1 + 2 * Fourier
        UC(1, 2) = -2 * Fourier

        k = 2
        For j = 2 To nodes - 1
            UC(j, k - 1) = -Fourier
            UC(j, k) = 1 + 2 * Fourier
            UC(j, k + 1) = -Fourier
            k = k + 1
        Next j

        j = nodes
        UC(j, k - 1) = -Fourier
        UC(j, k) = 1 + Fourier + Fourier2
        UC(j, k + 1) = -Fourier2
        k = k + 1

        For j = nodes + 1 To 2 * nodes - 2
            UC(j, k - 1) = -Fourier2
            UC(j, k) = 1 + 2 * Fourier2
            UC(j, k + 1) = -Fourier2
            k = k + 1
        Next j

        UC(2 * nodes - 1, 2 * nodes - 2) = -2 * Fourier2
        UC(2 * nodes - 1, 2 * nodes - 1) = 1 + 2 * Fourier2

        'interior boundary conditions
        CX(1, 1) = -2 * qnet * 1000 * Fourier * DeltaX / conductivity + TempNode(1, i)

        For k = 2 To 2 * nodes - 2
            CX(k, 1) = TempNode(k, i)
        Next k

        'exterior boundary conditions
        CX(2 * nodes - 1, 1) = 2 * Fourier2 * OutsideBiot2 * ((ExteriorTemp - TempNode(2 * nodes - 1, i)) - SubEmissivity / OutsideConvCoeff * StefanBoltzmann * (TempNode(2 * nodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + TempNode(2 * nodes - 1, i)

        Dim ier As Short
        If frmOptions1.optLUdecom.Checked = True Then
            'find surface temperatures for the next timestep
            'using method of LU decomposition (preferred)
            Call MatSol(UC, CX, 2 * nodes - 1)
        Else
            'find surface temperatures for the next timestep
            'using method of Gauss-Jordan elimination
            Call LINEAR2(2 * nodes - 1, UC, CX, ier)
            If ier = 1 Then MsgBox("singular matrix in implicit_surface_temps_furnace2")
        End If

        For j = 1 To 2 * nodes - 1
            TempNode(j, i + 1) = CX(j, 1)
        Next j

        Erase UC
        Erase CX

    End Sub
    Sub surfacetempcalc(ByVal index As Integer, ByVal room As Integer, ByVal submaterial As Boolean)
        'furnace test
        Dim TempNode(,) As Double
        Dim qnet As Double
        Dim convectco As Double = 0
        Dim Fourier As Double
        Dim Area As Double
        Dim DeltaX As Double
        Dim Conductivity As Double, SubConductivity As Double, SubThickness As Double, SubSpecificHeat As Double, SubDensity As Double
        Dim nodes As Integer
        Dim outsidebiot As Double
        Dim SurfaceEmissivity As Double = Surface_Emissivity(index, room)
        Dim SubEmissivity As Double = 0.8
        Dim thermalinertia As Double
        'Dim maxtimesteps As Integer = CInt(SimTime / Timestep)
        Dim maxtimesteps As Integer = 6 * 3600 / Timestep 'sec

        Dim ORIENT As String
        Dim i As Integer

        Derived_Variables()

        If index = 1 Then
            'ceiling
            Area = RoomFloorArea(room)
            Fourier = CeilingFourier(room)
            'thickness = CeilingThickness(room)
            DeltaX = CeilingDeltaX(room)
            Conductivity = CeilingConductivity(room)
            'density = CeilingDensity(room)
            'specificheat = CeilingSpecificHeat(room)
            nodes = Ceilingnodes
            outsidebiot = CeilingOutsideBiot(room)
            'outsidebiot = OutsideConvCoeff * ((thickness / 1000) / (nodes - 1)) / Conductivity
            thermalinertia = ThermalInertiaCeiling(room)
            ORIENT = 1
            'Fourier = Conductivity * Timestep / (((thickness / 1000) / (nodes - 1)) ^ 2 * density * specificheat)
            SubConductivity = CeilingSubConductivity(room)
            SubThickness = CeilingSubThickness(room)
            SubSpecificHeat = CeilingSubSpecificHeat(room)
            SubDensity = CeilingSubDensity(room)

        ElseIf index = 2 Or index = 3 Then
            Area = RoomHeight(room) * (RoomWidth(room) + RoomLength(room)) * 2
            Fourier = WallFourier(room)
            DeltaX = WallDeltaX(room)
            Conductivity = WallConductivity(room)
            nodes = Wallnodes
            outsidebiot = WallOutsideBiot(room)
            thermalinertia = ThermalInertiaWall(room)
            ORIENT = 2

            SubConductivity = WallSubConductivity(room)
            SubThickness = WallSubThickness(room)
            SubSpecificHeat = WallSubSpecificHeat(room)
            SubDensity = WallSubDensity(room)

            'thickness = WallThickness(room)
            'density = WallDensity(room)
            'specificheat = WallSpecificHeat(room)
            'outsidebiot = OutsideConvCoeff * ((thickness / 1000) / (nodes - 1)) / Conductivity
            'Fourier = Conductivity * Timestep / (((thickness / 1000) / (nodes - 1)) ^ 2 * density * specificheat)
        Else
            'floor
            Area = RoomFloorArea(room)
            Fourier = FloorFourier(room)
            DeltaX = FloorDeltaX(room)
            Conductivity = FloorConductivity(room)
            nodes = Floornodes
            outsidebiot = FloorOutsideBiot(room)
            thermalinertia = ThermalInertiaFloor(room)
            ORIENT = 1

            SubConductivity = FloorSubConductivity(room)
            SubThickness = FloorSubThickness(room)
            SubSpecificHeat = FloorSubSpecificHeat(room)
            SubDensity = FloorSubDensity(room)

            'thickness = WallThickness(room)
            'density = WallDensity(room)
            'specificheat = WallSpecificHeat(room)
            'outsidebiot = OutsideConvCoeff * ((thickness / 1000) / (nodes - 1)) / Conductivity
            'Fourier = Conductivity * Timestep / (((thickness / 1000) / (nodes - 1)) ^ 2 * density * specificheat)

        End If

        If submaterial = False Then
            'only have one surface layer
            ReDim TempNode(0 To nodes, 0 To maxtimesteps + 1)
            For m = 1 To nodes
                For j = 0 To maxtimesteps + 1
                    TempNode(m, j) = 293
                Next
            Next
        Else
            'have two layers
            ReDim TempNode(0 To 2 * nodes - 1, 0 To maxtimesteps + 1)
            For m = 1 To 2 * nodes - 1
                For j = 0 To maxtimesteps + 1
                    TempNode(m, j) = 293
                Next
            Next
        End If

        ReDim Preserve furnaceNHL(0 To 2, 0 To maxtimesteps + 1)
        ReDim Preserve furnaceST(0 To 2, 0 To maxtimesteps + 1)
        ReDim Preserve furnaceqnet(0 To 2, 0 To maxtimesteps + 1)
        Dim furnacetemp As Double = 293
        Dim surfacetemp As Double = 293
        Dim qsum As Double = 0
        Dim t As Double

        For i = 0 To maxtimesteps
            t = i * Timestep / 60
            furnacetemp = 345 * Log10(8 * t + 1) + 293 'K
            'convectco = Get_Convection_Coefficient2(Area, furnacetemp, surfacetemp, ORIENT)
            convectco = 25
            qnet = SurfaceEmissivity * StefanBoltzmann / 1000 * (furnacetemp ^ 4 - surfacetemp ^ 4) + convectco / 1000 * (furnacetemp - surfacetemp)

            If submaterial = False Then
                'only have one surface layer
                Implicit_Surface_Temps_furnace(-qnet, i, TempNode, Fourier, nodes, DeltaX, Conductivity, outsidebiot, SurfaceEmissivity)
            Else
                'have two layers
                Implicit_Surface_Temps_furnace2(-qnet, i, TempNode, Fourier, nodes, DeltaX, Conductivity, SubConductivity, outsidebiot, SurfaceEmissivity, SubThickness, SubSpecificHeat, SubDensity, SubEmissivity)
            End If

            surfacetemp = TempNode(1, i)

            If qnet > 0 Then
                qsum = qsum + qnet / Sqrt(thermalinertia) * Timestep 'normalised using surface layer properties?
            End If

            If index = 1 Then 'ceiling
                furnaceNHL(0, i) = qsum
                furnaceST(0, i) = surfacetemp
                furnaceqnet(0, i) = qnet
            ElseIf index = 2 Or index = 3 Then 'wall
                furnaceNHL(1, i) = qsum
                furnaceST(1, i) = surfacetemp
                furnaceqnet(1, i) = qnet
            Else 'floor
                furnaceNHL(2, i) = qsum
                furnaceST(2, i) = surfacetemp
                furnaceqnet(2, i) = qnet
            End If

        Next


    End Sub
	
    Sub Implicit_Surface_Temps(ByVal room As Integer, ByVal i As Integer, ByRef UWallNode(,,) As Double, ByRef CeilingNode(,,) As Double, ByRef LWallNode(,,) As Double, ByRef FloorNode(,,) As Double)
        '*  ================================================================
        '*      This function updates the surface temperatures, using an
        '*      implicit finite difference method.
        '*
        '*      Revised: Colleen Wade 14 July 1995
        '*  ================================================================

        Dim k, j As Integer
        Dim FloorFourier2, FloorOutsideBiot2 As Double
        Dim Lf(Floornodes, Floornodes) As Double
        Dim FX(Floornodes, 1) As Double
        Dim UW(Wallnodes, Wallnodes) As Double
        Dim UC(Ceilingnodes, Ceilingnodes) As Double
        Dim wx(Wallnodes, 1) As Double
        Dim cx(Ceilingnodes, 1) As Double
        Dim LX(Wallnodes, 1) As Double
        Dim LW(Wallnodes, Wallnodes) As Double
        Dim emissivity2 As Double = 0.8 '???

        If HaveFloorSubstrate(room) = False Then
            ReDim Lf(Floornodes, Floornodes)
            ReDim FX(Floornodes, 1)
        Else
            ReDim Lf(2 * Floornodes - 1, 2 * Floornodes - 1)
            ReDim FX(2 * Floornodes - 1, 1)
            FloorFourier2 = FloorSubConductivity(room) * Timestep / (((FloorSubThickness(room) / 1000) / (Floornodes - 1)) ^ 2 * FloorSubDensity(room) * FloorSubSpecificHeat(room))
            FloorOutsideBiot2 = OutsideConvCoeff * ((FloorSubThickness(room) / 1000) / (Floornodes - 1)) / FloorSubConductivity(room)
        End If

        UW(1, 1) = 1 + 2 * WallFourier(room)
        UW(1, 2) = -2 * WallFourier(room)
        LW(1, 1) = 1 + 2 * WallFourier(room)
        LW(1, 2) = -2 * WallFourier(room)
        UC(1, 1) = 1 + 2 * CeilingFourier(room)
        UC(1, 2) = -2 * CeilingFourier(room)
        Lf(1, 1) = 1 + 2 * FloorFourier(room)
        Lf(1, 2) = -2 * FloorFourier(room)

        k = 2
        For j = 2 To Wallnodes - 1
            UW(j, k - 1) = -WallFourier(room)
            UW(j, k) = 1 + 2 * WallFourier(room)
            UW(j, k + 1) = -WallFourier(room)
            LW(j, k - 1) = -WallFourier(room)
            LW(j, k) = 1 + 2 * WallFourier(room)
            LW(j, k + 1) = -WallFourier(room)
            k = k + 1
        Next j

        k = 2
        For j = 2 To Ceilingnodes - 1
            UC(j, k - 1) = -CeilingFourier(room)
            UC(j, k) = 1 + 2 * CeilingFourier(room)
            UC(j, k + 1) = -CeilingFourier(room)
            k = k + 1
        Next j

        If HaveFloorSubstrate(room) = False Then
            k = 2
            For j = 2 To Floornodes - 1
                Lf(j, k - 1) = -FloorFourier(room)
                Lf(j, k) = 1 + 2 * FloorFourier(room)
                Lf(j, k + 1) = -FloorFourier(room)
                k = k + 1
            Next j
        Else
            k = 2
            For j = 2 To Floornodes - 1
                Lf(j, k - 1) = -FloorFourier(room)
                Lf(j, k) = 1 + 2 * FloorFourier(room)
                Lf(j, k + 1) = -FloorFourier(room)
                k = k + 1
            Next j

            j = Floornodes
            Lf(j, k - 1) = -FloorFourier(room)
            Lf(j, k) = 1 + FloorFourier(room) + FloorFourier2
            Lf(j, k + 1) = -FloorFourier2
            k = k + 1

            For j = Floornodes + 1 To 2 * Floornodes - 2
                Lf(j, k - 1) = -FloorFourier2
                Lf(j, k) = 1 + 2 * FloorFourier2
                Lf(j, k + 1) = -FloorFourier2
                k = k + 1
            Next j
        End If

        UW(Wallnodes, Wallnodes - 1) = -2 * WallFourier(room)
        UW(Wallnodes, Wallnodes) = 1 + 2 * WallFourier(room)
        LW(Wallnodes, Wallnodes - 1) = -2 * WallFourier(room)
        LW(Wallnodes, Wallnodes) = 1 + 2 * WallFourier(room)
        UC(Ceilingnodes, Ceilingnodes - 1) = -2 * CeilingFourier(room)
        UC(Ceilingnodes, Ceilingnodes) = 1 + 2 * CeilingFourier(room)
        If HaveFloorSubstrate(room) = False Then
            Lf(Floornodes, Floornodes - 1) = -2 * FloorFourier(room)
            Lf(Floornodes, Floornodes) = 1 + 2 * FloorFourier(room)
        Else
            Lf(2 * Floornodes - 1, 2 * Floornodes - 2) = -2 * FloorFourier2
            Lf(2 * Floornodes - 1, 2 * Floornodes - 1) = 1 + 2 * FloorFourier2
        End If

        'interior boundary conditions
        wx(1, 1) = -2 * QUpperWall(room, i) * 1000 * WallFourier(room) * WallDeltaX(room) / WallConductivity(room) + UWallNode(room, 1, i)
        cx(1, 1) = -2 * QCeiling(room, i) * 1000 * CeilingFourier(room) * CeilingDeltaX(room) / CeilingConductivity(room) + CeilingNode(room, 1, i)
        FX(1, 1) = -2 * QFloor(room, i) * 1000 * FloorFourier(room) * FloorDeltaX(room) / FloorConductivity(room) + FloorNode(room, 1, i)
        LX(1, 1) = -2 * QLowerWall(room, i) * 1000 * WallFourier(room) * WallDeltaX(room) / WallConductivity(room) + LWallNode(room, 1, i)

        For k = 2 To Wallnodes - 1
            wx(k, 1) = UWallNode(room, k, i)
            LX(k, 1) = LWallNode(room, k, i)
        Next k
        For k = 2 To Ceilingnodes - 1
            cx(k, 1) = CeilingNode(room, k, i)
        Next k
        If HaveFloorSubstrate(room) = False Then
            For k = 2 To Floornodes - 1
                FX(k, 1) = FloorNode(room, k, i)
            Next k
        Else
            For k = 2 To 2 * Floornodes - 2
                FX(k, 1) = FloorNode(room, k, i)
            Next k
        End If
        'exterior boundary conditions
        wx(Wallnodes, 1) = 2 * WallFourier(room) * WallOutsideBiot(room) * ((ExteriorTemp - UWallNode(room, Wallnodes, i)) - Surface_Emissivity(2, room) / OutsideConvCoeff * StefanBoltzmann * (UWallNode(room, Wallnodes, i) ^ 4 - ExteriorTemp ^ 4)) + UWallNode(room, Wallnodes, i)
        cx(Ceilingnodes, 1) = 2 * CeilingFourier(room) * CeilingOutsideBiot(room) * ((ExteriorTemp - CeilingNode(room, Ceilingnodes, i)) - Surface_Emissivity(1, room) / OutsideConvCoeff * StefanBoltzmann * (CeilingNode(room, Ceilingnodes, i) ^ 4 - ExteriorTemp ^ 4)) + CeilingNode(room, Ceilingnodes, i)
        LX(Wallnodes, 1) = 2 * WallFourier(room) * WallOutsideBiot(room) * ((ExteriorTemp - LWallNode(room, Wallnodes, i)) - Surface_Emissivity(3, room) / OutsideConvCoeff * StefanBoltzmann * (LWallNode(room, Wallnodes, i) ^ 4 - ExteriorTemp ^ 4)) + LWallNode(room, Wallnodes, i)
        If HaveFloorSubstrate(room) = False Then
            FX(Floornodes, 1) = 2 * FloorFourier(room) * FloorOutsideBiot(room) * ((ExteriorTemp - FloorNode(room, Floornodes, i)) - Surface_Emissivity(4, room) / OutsideConvCoeff * StefanBoltzmann * (FloorNode(room, Floornodes, i) ^ 4 - ExteriorTemp ^ 4)) + FloorNode(room, Floornodes, i)
        Else
            'FX(2 * Floornodes - 1, 1) = 2 * FloorFourier2 * FloorOutsideBiot2 * ((ExteriorTemp - FloorNode(room, 2 * Floornodes - 1, i)) - Surface_Emissivity(4, room) / OutsideConvCoeff * StefanBoltzmann * (FloorNode(room, 2 * Floornodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + FloorNode(room, 2 * Floornodes - 1, i)
            FX(2 * Floornodes - 1, 1) = 2 * FloorFourier2 * FloorOutsideBiot2 * ((ExteriorTemp - FloorNode(room, 2 * Floornodes - 1, i)) - emissivity2 / OutsideConvCoeff * StefanBoltzmann * (FloorNode(room, 2 * Floornodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + FloorNode(room, 2 * Floornodes - 1, i)
        End If

        Dim ier As Short
        If frmOptions1.optLUdecom.Checked = True Then
            'find surface temperatures for the next timestep
            'using method of LU decomposition (preferred)
            Call MatSol(UW, wx, Wallnodes) 'upper wall
            Call MatSol(UC, cx, Ceilingnodes) 'ceiling
            Call MatSol(LW, LX, Wallnodes) 'lower wall
            If HaveFloorSubstrate(room) = False Then
                Call MatSol(Lf, FX, Floornodes) 'floor
            Else
                Call MatSol(Lf, FX, 2 * Floornodes - 1)
            End If
        Else
            'find surface temperatures for the next timestep
            'using method of Gauss-Jordan elimination
            Call LINEAR2(Wallnodes, UW, wx, ier)
            Call LINEAR2(Ceilingnodes, UC, cx, ier)
            Call LINEAR2(Wallnodes, LW, LX, ier)
            If HaveFloorSubstrate(room) = False Then
                Call LINEAR2(Floornodes, Lf, FX, ier)
            Else
                Call LINEAR2(2 * Floornodes - 1, Lf, FX, ier)
            End If
            If ier = 1 Then MsgBox("singular matrix in implicit_surface_temps")
        End If

        For j = 1 To Wallnodes
            LWallNode(room, j, i + 1) = LX(j, 1)
            UWallNode(room, j, i + 1) = wx(j, 1)
        Next j
        For j = 1 To Ceilingnodes
            CeilingNode(room, j, i + 1) = cx(j, 1)
        Next j
        If HaveFloorSubstrate(room) = False Then
            For j = 1 To Floornodes
                FloorNode(room, j, i + 1) = FX(j, 1)
            Next j
        Else
            For j = 1 To 2 * Floornodes - 1
                FloorNode(room, j, i + 1) = FX(j, 1)
            Next j
        End If

        'store surface temps in another array
        Upperwalltemp(room, i + 1) = UWallNode(room, 1, i + 1)
        CeilingTemp(room, i + 1) = CeilingNode(room, 1, i + 1)
        LowerWallTemp(room, i + 1) = LWallNode(room, 1, i + 1)
        FloorTemp(room, i + 1) = FloorNode(room, 1, i + 1)

        UnexposedUpperwalltemp(room, i + 1) = UWallNode(room, Wallnodes, i + 1)
        UnexposedLowerwalltemp(room, i + 1) = LWallNode(room, Wallnodes, i + 1)
        UnexposedCeilingtemp(room, i + 1) = CeilingNode(room, Ceilingnodes, i + 1)
        If HaveFloorSubstrate(room) = True Then
            UnexposedFloortemp(room, i + 1) = FloorNode(room, 2 * Floornodes - 1, i + 1)
        Else
            UnexposedFloortemp(room, i + 1) = FloorNode(room, Floornodes, i + 1)
        End If
        Erase UW
        Erase UC
        Erase wx
        Erase cx
        Erase Lf
        Erase FX
        Erase LX
        Erase LW
    End Sub
    Sub Implicit_Surface_Temps_CLT(ByVal room As Integer, ByVal i As Integer, ByRef UWallNode(,,) As Double, ByRef CeilingNode(,,) As Double, ByRef LWallNode(,,) As Double, ByRef FloorNode(,,) As Double)
        'not used
        '*  ================================================================
        '*      This function updates the surface temperatures, using an
        '*      implicit finite difference method.
        '*
        '*      Revised: Colleen Wade 14 July 1995
        '*  ================================================================

        Try

            Dim NLw As Integer = WallThickness(room) / 1000 / Lamella 'number of lamella - in two places also in main_program2
            Dim NLc As Integer = CeilingThickness(room) / 1000 / Lamella 'number of lamella - in two places also in main_program2

            Dim walllayersremaining As Integer = NLw - Lamella2 / Lamella + 1
            Dim ceillayersremaining As Integer = NLc - Lamella1 / Lamella + 1

            If walllayersremaining = 0 Or ceillayersremaining = 0 Then
                If walllayersremaining = 0 Then Dim Message As String = CStr(tim(stepcount, 1)) & " sec. Wall layer at " & Format(Lamella2, "0.000") & " m delaminates. Simulation terminated. "
                If ceillayersremaining = 0 Then Dim Message As String = CStr(tim(stepcount, 1)) & " sec. Ceiling layer at " & Format(Lamella1, "0.000") & " m delaminates. Simulation terminated. "

                flagstop = 1 'do not continue
                Exit Sub
            End If

            Dim k, j As Integer

            Dim Wallnodestemp As Integer = (Wallnodes - 1) * (walllayersremaining) + 1 'remove one layer and recalc number of nodes
            Dim UWallNodeTemp(Wallnodestemp) As Double
            Dim LWallNodeTemp(Wallnodestemp) As Double
            Dim wallnodeadjust As Integer = (Wallnodes - 1) * NLw + 1 - Wallnodestemp

            For k = 1 To Wallnodestemp
                UWallNodeTemp(k) = UWallNode(room, k + wallnodeadjust, i)
                LWallNodeTemp(k) = LWallNode(room, k + wallnodeadjust, i)
            Next
            Dim UWallNodeExposed As Double = UWallNodeTemp(1)
            Dim UWallNodeUnExposed As Double = UWallNodeTemp(Wallnodestemp)
            Dim LWallNodeUnExposed As Double = LWallNodeTemp(Wallnodestemp)
            Dim LWallNodeExposed As Double = LWallNodeTemp(1)

            Dim Ceilingnodestemp As Integer = (Ceilingnodes - 1) * (ceillayersremaining) + 1 'remove one layer and recalc number of nodes
            Dim CeilingNodeTemp(Ceilingnodestemp) As Double
            Dim ceilingnodeadjust As Integer = (Ceilingnodes - 1) * NLc + 1 - Ceilingnodestemp

            For k = 1 To Ceilingnodestemp
                CeilingNodeTemp(k) = CeilingNode(room, k + ceilingnodeadjust, i)
            Next
            Dim CeilingNodeUnExposed As Double = CeilingNodeTemp(Ceilingnodestemp)
            Dim CeilingNodeExposed As Double = CeilingNodeTemp(1)

            Dim Lf(Floornodes, Floornodes) As Double
            Dim FX(Floornodes, 1) As Double
            Dim FloorFourier2, FloorOutsideBiot2 As Double

            If HaveFloorSubstrate(room) = False Then
                ReDim Lf(Floornodes, Floornodes)
                ReDim FX(Floornodes, 1)
            Else
                ReDim Lf(2 * Floornodes - 1, 2 * Floornodes - 1)
                ReDim FX(2 * Floornodes - 1, 1)
                FloorFourier2 = FloorSubConductivity(room) * Timestep / (((FloorSubThickness(room) / 1000) / (Floornodes - 1)) ^ 2 * FloorSubDensity(room) * FloorSubSpecificHeat(room))
                FloorOutsideBiot2 = OutsideConvCoeff * ((FloorSubThickness(room) / 1000) / (Floornodes - 1)) / FloorSubConductivity(room)
            End If

            Lf(1, 1) = 1 + 2 * FloorFourier(room)
            Lf(1, 2) = -2 * FloorFourier(room)

            If HaveFloorSubstrate(room) = False Then
                k = 2
                For j = 2 To Floornodes - 1
                    Lf(j, k - 1) = -FloorFourier(room)
                    Lf(j, k) = 1 + 2 * FloorFourier(room)
                    Lf(j, k + 1) = -FloorFourier(room)
                    k = k + 1
                Next j
            Else
                k = 2
                For j = 2 To Floornodes - 1
                    Lf(j, k - 1) = -FloorFourier(room)
                    Lf(j, k) = 1 + 2 * FloorFourier(room)
                    Lf(j, k + 1) = -FloorFourier(room)
                    k = k + 1
                Next j

                j = Floornodes
                Lf(j, k - 1) = -FloorFourier(room)
                Lf(j, k) = 1 + FloorFourier(room) + FloorFourier2
                Lf(j, k + 1) = -FloorFourier2
                k = k + 1

                For j = Floornodes + 1 To 2 * Floornodes - 2
                    Lf(j, k - 1) = -FloorFourier2
                    Lf(j, k) = 1 + 2 * FloorFourier2
                    Lf(j, k + 1) = -FloorFourier2
                    k = k + 1
                Next j
            End If

            'Find DeltaX
            'WallDeltaX(fireroom) = (WallThickness(room) / 1000) / (Wallnodestemp - 1)
            'CeilingDeltaX(fireroom) = (CeilingThickness(room) / 1000) / (Ceilingnodestemp - 1)
            WallDeltaX(fireroom) = walllayersremaining * Lamella / (Wallnodestemp - 1)
            CeilingDeltaX(fireroom) = ceillayersremaining * Lamella / (Ceilingnodestemp - 1)

            'Find Biot Numbers
            WallOutsideBiot(room) = OutsideConvCoeff * WallDeltaX(room) / WallConductivity(room)
            CeilingOutsideBiot(room) = OutsideConvCoeff * CeilingDeltaX(room) / CeilingConductivity(room)

            'Find Fourier Numbers
            WallFourier(fireroom) = AlphaWall(room) * Timestep / (WallDeltaX(room)) ^ 2
            CeilingFourier(fireroom) = AlphaCeiling(room) * Timestep / (CeilingDeltaX(room) ^ 2)

            Dim LW(Wallnodestemp, Wallnodestemp) As Double
            Dim UW(Wallnodestemp, Wallnodestemp) As Double
            Dim UC(Ceilingnodestemp, Ceilingnodestemp) As Double
            Dim wx(Wallnodestemp, 1) As Double
            Dim cx(Ceilingnodestemp, 1) As Double
            Dim LX(Wallnodestemp, 1) As Double

            LW(1, 1) = 1 + 2 * WallFourier(room)
            LW(1, 2) = -2 * WallFourier(room)

            UW(1, 1) = 1 + 2 * WallFourier(room)
            UW(1, 2) = -2 * WallFourier(room)

            UC(1, 1) = 1 + 2 * CeilingFourier(room)
            UC(1, 2) = -2 * CeilingFourier(room)

            k = 2
            For j = 2 To Wallnodestemp - 1
                UW(j, k - 1) = -WallFourier(room)
                UW(j, k) = 1 + 2 * WallFourier(room)
                UW(j, k + 1) = -WallFourier(room)
                LW(j, k - 1) = -WallFourier(room)
                LW(j, k) = 1 + 2 * WallFourier(room)
                LW(j, k + 1) = -WallFourier(room)
                k = k + 1
            Next j

            k = 2
            For j = 2 To Ceilingnodestemp - 1
                UC(j, k - 1) = -CeilingFourier(room)
                UC(j, k) = 1 + 2 * CeilingFourier(room)
                UC(j, k + 1) = -CeilingFourier(room)
                k = k + 1
            Next j

            k = 2

            LW(Wallnodestemp, Wallnodestemp - 1) = -2 * WallFourier(room)
            LW(Wallnodestemp, Wallnodestemp) = 1 + 2 * WallFourier(room)

            UW(Wallnodestemp, Wallnodestemp - 1) = -2 * WallFourier(room)
            UW(Wallnodestemp, Wallnodestemp) = 1 + 2 * WallFourier(room)

            UC(Ceilingnodestemp, Ceilingnodestemp - 1) = -2 * CeilingFourier(room)
            UC(Ceilingnodestemp, Ceilingnodestemp) = 1 + 2 * CeilingFourier(room)

            If HaveFloorSubstrate(room) = False Then
                Lf(Floornodes, Floornodes - 1) = -2 * FloorFourier(room)
                Lf(Floornodes, Floornodes) = 1 + 2 * FloorFourier(room)
            Else
                Lf(2 * Floornodes - 1, 2 * Floornodes - 2) = -2 * FloorFourier2
                Lf(2 * Floornodes - 1, 2 * Floornodes - 1) = 1 + 2 * FloorFourier2
            End If

            'interior boundary conditions
            wx(1, 1) = -2 * QUpperWall(room, i) * 1000 * WallFourier(room) * WallDeltaX(room) / WallConductivity(room) + UWallNodeExposed
            For k = 2 To Wallnodestemp - 1
                wx(k, 1) = UWallNodeTemp(k)
            Next k

            LX(1, 1) = -2 * QLowerWall(room, i) * 1000 * WallFourier(room) * WallDeltaX(room) / WallConductivity(room) + LWallNodeExposed
            For k = 2 To Wallnodestemp - 1
                LX(k, 1) = LWallNodeTemp(k)
            Next k

            cx(1, 1) = -2 * QCeiling(room, i) * 1000 * CeilingFourier(room) * CeilingDeltaX(room) / CeilingConductivity(room) + CeilingNodeExposed
            For k = 2 To Ceilingnodestemp - 1
                cx(k, 1) = CeilingNodeTemp(k)
            Next k

            FX(1, 1) = -2 * QFloor(room, i) * 1000 * FloorFourier(room) * FloorDeltaX(room) / FloorConductivity(room) + FloorNode(room, 1, i)
            If HaveFloorSubstrate(room) = False Then
                For k = 2 To Floornodes - 1
                    FX(k, 1) = FloorNode(room, k, i)
                Next k
            Else
                For k = 2 To 2 * Floornodes - 2
                    FX(k, 1) = FloorNode(room, k, i)
                Next k
            End If

            If HaveFloorSubstrate(room) = False Then
                FX(Floornodes, 1) = 2 * FloorFourier(room) * FloorOutsideBiot(room) * ((ExteriorTemp - FloorNode(room, Floornodes, i)) - Surface_Emissivity(4, room) / OutsideConvCoeff * StefanBoltzmann * (FloorNode(room, Floornodes, i) ^ 4 - ExteriorTemp ^ 4)) + FloorNode(room, Floornodes, i)
            Else
                FX(2 * Floornodes - 1, 1) = 2 * FloorFourier2 * FloorOutsideBiot2 * ((ExteriorTemp - FloorNode(room, 2 * Floornodes - 1, i)) - Surface_Emissivity(4, room) / OutsideConvCoeff * StefanBoltzmann * (FloorNode(room, 2 * Floornodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + FloorNode(room, 2 * Floornodes - 1, i)
            End If

            'exterior boundary conditions
            wx(Wallnodestemp, 1) = 2 * WallFourier(room) * WallOutsideBiot(room) * ((ExteriorTemp - UWallNodeUnExposed) - Surface_Emissivity(2, room) / OutsideConvCoeff * StefanBoltzmann * (UWallNodeUnExposed ^ 4 - ExteriorTemp ^ 4)) + UWallNodeUnExposed
            cx(Ceilingnodestemp, 1) = 2 * CeilingFourier(room) * CeilingOutsideBiot(room) * ((ExteriorTemp - CeilingNodeUnExposed) - Surface_Emissivity(1, room) / OutsideConvCoeff * StefanBoltzmann * (CeilingNodeUnExposed ^ 4 - ExteriorTemp ^ 4)) + CeilingNodeUnExposed
            LX(Wallnodestemp, 1) = 2 * WallFourier(room) * WallOutsideBiot(room) * ((ExteriorTemp - LWallNodeUnExposed) - Surface_Emissivity(2, room) / OutsideConvCoeff * StefanBoltzmann * (LWallNodeUnExposed ^ 4 - ExteriorTemp ^ 4)) + LWallNodeUnExposed

            Dim ier As Short
            If frmOptions1.optLUdecom.Checked = True Then
                'find surface temperatures for the next timestep
                'using method of LU decomposition (preferred)
                Call MatSol(UW, wx, Wallnodestemp) 'upper wall
                Call MatSol(UC, cx, Ceilingnodestemp) 'ceiling
                Call MatSol(LW, LX, Wallnodestemp) 'lower wall
                If HaveFloorSubstrate(room) = False Then
                    Call MatSol(Lf, FX, Floornodes) 'floor
                Else
                    Call MatSol(Lf, FX, 2 * Floornodes - 1)
                End If
            Else
                'find surface temperatures for the next timestep
                'using method of Gauss-Jordan elimination
                Call LINEAR2(Wallnodestemp, UW, wx, ier)
                Call LINEAR2(Ceilingnodestemp, UC, cx, ier)
                Call LINEAR2(Wallnodestemp, LW, LX, ier)
                If HaveFloorSubstrate(room) = False Then
                    Call LINEAR2(Floornodes, Lf, FX, ier)
                Else
                    Call LINEAR2(2 * Floornodes - 1, Lf, FX, ier)
                End If
                If ier = 1 Then MsgBox("singular matrix in implicit_surface_temps")
            End If

            For j = 1 To Wallnodestemp
                UWallNode(room, j + wallnodeadjust, i + 1) = wx(j, 1)
                LWallNode(room, j + wallnodeadjust, i + 1) = LX(j, 1)
            Next j
            For j = 1 To wallnodeadjust
                UWallNode(room, j, i + 1) = chartemp + 273 + 1
                LWallNode(room, j, i + 1) = chartemp + 273 + 1
            Next

            For j = 1 To Ceilingnodestemp
                CeilingNode(room, j + ceilingnodeadjust, i + 1) = cx(j, 1)
            Next j
            For j = 1 To ceilingnodeadjust
                CeilingNode(room, j, i + 1) = chartemp + 273 + 1
            Next

            If HaveFloorSubstrate(room) = False Then
                For j = 1 To Floornodes
                    FloorNode(room, j, i + 1) = FX(j, 1)
                Next j
            Else
                For j = 1 To 2 * Floornodes - 1
                    FloorNode(room, j, i + 1) = FX(j, 1)
                Next j
            End If

            'store surface temps in another array
            Upperwalltemp(room, i + 1) = UWallNode(room, 1 + wallnodeadjust, i + 1)
            CeilingTemp(room, i + 1) = CeilingNode(room, 1 + ceilingnodeadjust, i + 1)
            LowerWallTemp(room, i + 1) = LWallNode(room, 1 + wallnodeadjust, i + 1)
            FloorTemp(room, i + 1) = FloorNode(room, 1, i + 1)

            UnexposedUpperwalltemp(room, i + 1) = UWallNode(room, Wallnodestemp, i + 1)
            UnexposedCeilingtemp(room, i + 1) = CeilingNode(room, Ceilingnodestemp, i + 1)
            UnexposedLowerwalltemp(room, i + 1) = LWallNode(room, Wallnodestemp, i + 1)
            If HaveFloorSubstrate(room) = True Then
                UnexposedFloortemp(room, i + 1) = FloorNode(room, 2 * Floornodes - 1, i + 1)
            Else
                UnexposedFloortemp(room, i + 1) = FloorNode(room, Floornodes, i + 1)
            End If

            Erase UW
            Erase UC
            Erase wx
            Erase cx
            Erase Lf
            Erase FX
            Erase LX
            Erase LW

        Catch ex As Exception
            MsgBox(Err.Description & " Line " & Err.Erl, MsgBoxStyle.Exclamation, "Exception in Implicit_Surface_TEmp_CLT ")
        End Try
    End Sub
    Sub Implicit_Surface_Temps_CLTC(ByVal room As Integer, ByVal i As Integer, ByRef CeilingNode(,,) As Double)
        'not used
        '*  ================================================================
        '*      This function updates the surface temperatures, using an
        '*      implicit finite difference method.
        '*
        '*  ================================================================

        Try
            Dim ceilinglayersremaining As Integer
            Dim k, j As Integer
            Dim ceilingnodestemp As Integer
            Dim NLC As Integer
            Dim ier As Short
            Dim ceilingnodeadjust As Integer
            Dim chardensity As Double = 150 'kg/m3


            If CLTceilingpercent > 0 Then
                NLC = CeilingThickness(room) / 1000 / Lamella 'number of lamella - in two places also in main_program2

                ceilinglayersremaining = NLC - Lamella1 / Lamella + 1

                If ceilinglayersremaining = 0 Then
                    If ceilinglayersremaining = 0 Then Dim Message As String = CStr(tim(stepcount, 1)) & " sec. Ceiling layer at " & Format(Lamella1, "0.000") & " m delaminates. Simulation terminated. "

                    flagstop = 1 'do not continue
                    Exit Sub
                End If

                ceilingnodestemp = (Ceilingnodes - 1) * (ceilinglayersremaining) + 1 'remove one layer and recalc number of nodes
                ceilingnodeadjust = (Ceilingnodes - 1) * NLC + 1 - ceilingnodestemp

                'Find DeltaX
                CeilingDeltaX(room) = ceilinglayersremaining * Lamella / (ceilingnodestemp - 1)
            Else
                'clt model is on but not for ceiling
                NLC = 1
                ceilinglayersremaining = 1
                ceilingnodestemp = Ceilingnodes
                ceilingnodeadjust = 0
            End If

            Dim CeilingNodeTemp(ceilingnodestemp) As Double
            Dim CeilingNodeStatus(ceilingnodestemp) As Integer  'array to identify charred element

            For k = 1 To ceilingnodestemp
                CeilingNodeTemp(k) = CeilingNode(room, k + ceilingnodeadjust, i)

                'array to identify charred element
                If CeilingNodeTemp(k) > 300 + 273 Then CeilingNodeStatus(ceilingnodestemp) = 1

            Next

            Dim CeilingNodeUnExposed As Double = CeilingNodeTemp(ceilingnodestemp)
            Dim CeilingNodeExposed As Double = CeilingNodeTemp(1)

            'Find Biot Numbers
            CeilingOutsideBiot(room) = OutsideConvCoeff * CeilingDeltaX(room) / CeilingConductivity(room)

            'Find Fourier Numbers
            CeilingFourier(room) = AlphaCeiling(room) * Timestep / (CeilingDeltaX(room)) ^ 2

            '  AlphaCeiling(room) = CeilingConductivity(room) / (CeilingSpecificHeat(room) * CeilingDensity(room))

            Dim UC(ceilingnodestemp, ceilingnodestemp) As Double
            Dim CX(ceilingnodestemp, 1) As Double

            UC(1, 1) = 1 + 2 * CeilingFourier(room)
            UC(1, 2) = -2 * CeilingFourier(room)

            k = 2
            For j = 2 To ceilingnodestemp - 1
                UC(j, k - 1) = -CeilingFourier(room)
                UC(j, k) = 1 + 2 * CeilingFourier(room)
                UC(j, k + 1) = -CeilingFourier(room)

                k = k + 1
            Next j

            UC(ceilingnodestemp, ceilingnodestemp - 1) = -2 * CeilingFourier(room)
            UC(ceilingnodestemp, ceilingnodestemp) = 1 + 2 * CeilingFourier(room)

            'exterior boundary conditions
            CX(ceilingnodestemp, 1) = 2 * CeilingFourier(room) * CeilingOutsideBiot(room) * ((ExteriorTemp - CeilingNodeUnExposed) - Surface_Emissivity(1, room) / OutsideConvCoeff * StefanBoltzmann * (CeilingNodeUnExposed ^ 4 - ExteriorTemp ^ 4)) + CeilingNodeUnExposed

            'interior boundary conditions
            CX(1, 1) = -2 * QCeiling(room, i) * 1000 * CeilingFourier(room) * CeilingDeltaX(room) / CeilingConductivity(room) + CeilingNodeExposed
            For k = 2 To ceilingnodestemp - 1
                CX(k, 1) = CeilingNodeTemp(k)
            Next k

            If frmOptions1.optLUdecom.Checked = True Then
                'find surface temperatures for the next timestep
                'using method of LU decomposition (preferred)
                Call MatSol(UC, CX, ceilingnodestemp) 'ceiling

            Else
                'find surface temperatures for the next timestep
                'using method of Gauss-Jordan elimination
                Call LINEAR2(ceilingnodestemp, UC, CX, ier)
                If ier = 1 Then MsgBox("singular matrix in implicit_surface_temps_CLTC")
            End If

            For j = 1 To ceilingnodestemp
                CeilingNode(room, j + ceilingnodeadjust, i + 1) = CX(j, 1)
            Next j
            For j = 1 To ceilingnodeadjust
                CeilingNode(room, j, i + 1) = chartemp + 273 + 1
            Next

            'store surface temps at next timestep in another array
            CeilingTemp(room, i + 1) = CeilingNode(room, 1 + ceilingnodeadjust, i + 1)
            UnexposedCeilingtemp(room, i + 1) = CeilingNode(room, ceilingnodestemp, i + 1)

            Erase UC
            Erase CX

        Catch ex As Exception
            MsgBox(Err.Description & " Line " & Err.Erl, MsgBoxStyle.Exclamation, "Exception in Implicit_Surface_Temp_CLTC")
        End Try
    End Sub

    Sub Implicit_Surface_Temps_CLTW(ByVal room As Integer, ByVal i As Integer, ByRef UWallNode(,,) As Double, ByRef LWallNode(,,) As Double)
        'not used
        '*  ================================================================
        '*      This function updates the surface temperatures, using an
        '*      implicit finite difference method.
        '*
        '*  ================================================================

        Try
            Dim walllayersremaining As Integer
            Dim k, j As Integer
            Dim Wallnodestemp As Integer
            Dim NLw As Integer
            Dim ier As Short
            Dim wallnodeadjust As Integer

            If CLTwallpercent > 0 Then
                NLw = WallThickness(room) / 1000 / Lamella 'number of lamella - in two places also in main_program2

                walllayersremaining = NLw - Lamella2 / Lamella + 1

                If walllayersremaining = 0 Then
                    If walllayersremaining = 0 Then Dim Message As String = CStr(tim(stepcount, 1)) & " sec. Wall layer at " & Format(Lamella2, "0.000") & " m delaminates. Simulation terminated. "

                    flagstop = 1 'do not continue
                    Exit Sub
                End If

                Wallnodestemp = (Wallnodes - 1) * (walllayersremaining) + 1 'remove one layer and recalc number of nodes
                wallnodeadjust = (Wallnodes - 1) * NLw + 1 - Wallnodestemp

                'Find DeltaX
                WallDeltaX(room) = walllayersremaining * Lamella / (Wallnodestemp - 1)
            Else
                'clt model is on but not for wall
                NLw = 1
                walllayersremaining = 1
                Wallnodestemp = Wallnodes
                wallnodeadjust = 0
            End If

            Dim UWallNodeTemp(Wallnodestemp) As Double
            Dim LWallNodeTemp(Wallnodestemp) As Double

            For k = 1 To Wallnodestemp
                UWallNodeTemp(k) = UWallNode(room, k + wallnodeadjust, i)
                LWallNodeTemp(k) = LWallNode(room, k + wallnodeadjust, i)
            Next

            Dim UWallNodeUnExposed As Double = UWallNodeTemp(Wallnodestemp)
            Dim LWallNodeUnExposed As Double = LWallNodeTemp(Wallnodestemp)
            Dim LWallNodeExposed As Double = LWallNodeTemp(1)
            Dim UWallNodeExposed As Double = UWallNodeTemp(1)

            'Find Biot Numbers
            WallOutsideBiot(room) = OutsideConvCoeff * WallDeltaX(room) / WallConductivity(room)

            'Find Fourier Numbers
            WallFourier(room) = AlphaWall(room) * Timestep / (WallDeltaX(room)) ^ 2

            Dim LW(Wallnodestemp, Wallnodestemp) As Double
            Dim UW(Wallnodestemp, Wallnodestemp) As Double
            Dim WX(Wallnodestemp, 1) As Double
            Dim LX(Wallnodestemp, 1) As Double

            LW(1, 1) = 1 + 2 * WallFourier(room)
            LW(1, 2) = -2 * WallFourier(room)

            UW(1, 1) = 1 + 2 * WallFourier(room)
            UW(1, 2) = -2 * WallFourier(room)

            k = 2
            For j = 2 To Wallnodestemp - 1
                UW(j, k - 1) = -WallFourier(room)
                UW(j, k) = 1 + 2 * WallFourier(room)
                UW(j, k + 1) = -WallFourier(room)
                LW(j, k - 1) = -WallFourier(room)
                LW(j, k) = 1 + 2 * WallFourier(room)
                LW(j, k + 1) = -WallFourier(room)
                k = k + 1
            Next j

            LW(Wallnodestemp, Wallnodestemp - 1) = -2 * WallFourier(room)
            LW(Wallnodestemp, Wallnodestemp) = 1 + 2 * WallFourier(room)

            UW(Wallnodestemp, Wallnodestemp - 1) = -2 * WallFourier(room)
            UW(Wallnodestemp, Wallnodestemp) = 1 + 2 * WallFourier(room)

            'exterior boundary conditions
            WX(Wallnodestemp, 1) = 2 * WallFourier(room) * WallOutsideBiot(room) * ((ExteriorTemp - UWallNodeUnExposed) - Surface_Emissivity(2, room) / OutsideConvCoeff * StefanBoltzmann * (UWallNodeUnExposed ^ 4 - ExteriorTemp ^ 4)) + UWallNodeUnExposed
            LX(Wallnodestemp, 1) = 2 * WallFourier(room) * WallOutsideBiot(room) * ((ExteriorTemp - LWallNodeUnExposed) - Surface_Emissivity(2, room) / OutsideConvCoeff * StefanBoltzmann * (LWallNodeUnExposed ^ 4 - ExteriorTemp ^ 4)) + LWallNodeUnExposed

            'interior boundary conditions
            WX(1, 1) = -2 * QUpperWall(room, i) * 1000 * WallFourier(room) * WallDeltaX(room) / WallConductivity(room) + UWallNodeExposed
            For k = 2 To Wallnodestemp - 1
                WX(k, 1) = UWallNodeTemp(k)
            Next k

            LX(1, 1) = -2 * QLowerWall(room, i) * 1000 * WallFourier(room) * WallDeltaX(room) / WallConductivity(room) + LWallNodeExposed
            For k = 2 To Wallnodestemp - 1
                LX(k, 1) = LWallNodeTemp(k)
            Next k

            If frmOptions1.optLUdecom.Checked = True Then
                'find surface temperatures for the next timestep
                'using method of LU decomposition (preferred)
                Call MatSol(UW, WX, Wallnodestemp) 'upper wall
                Call MatSol(LW, LX, Wallnodestemp) 'lower wall
            Else
                'find surface temperatures for the next timestep
                'using method of Gauss-Jordan elimination
                Call LINEAR2(Wallnodestemp, UW, WX, ier)
                Call LINEAR2(Wallnodestemp, LW, LX, ier)
                If ier = 1 Then MsgBox("singular matrix in implicit_surface_temps_CLTW")
            End If

            For j = 1 To Wallnodestemp
                UWallNode(room, j + wallnodeadjust, i + 1) = WX(j, 1)
                LWallNode(room, j + wallnodeadjust, i + 1) = LX(j, 1)
            Next j
            For j = 1 To wallnodeadjust
                UWallNode(room, j, i + 1) = chartemp + 273 + 1
                LWallNode(room, j, i + 1) = chartemp + 273 + 1
            Next

            'store surface temps at next timestep in another array
            Upperwalltemp(room, i + 1) = UWallNode(room, 1 + wallnodeadjust, i + 1)
            LowerWallTemp(room, i + 1) = LWallNode(room, 1 + wallnodeadjust, i + 1)
            UnexposedUpperwalltemp(room, i + 1) = UWallNode(room, Wallnodestemp, i + 1)
            UnexposedLowerwalltemp(room, i + 1) = LWallNode(room, Wallnodestemp, i + 1)

            Erase LX
            Erase UW
            Erase WX
            Erase LW

        Catch ex As Exception
            MsgBox(Err.Description & " Line " & Err.Erl, MsgBoxStyle.Exclamation, "Exception in Implicit_Surface_Temp_CLTW ")
        End Try
    End Sub
    Sub Implicit_Surface_Temps_CLTW_Char(ByVal room As Integer, ByVal i As Integer, ByRef UWallNode(,,) As Double, ByRef LWallNode(,,) As Double)
        '*  ================================================================
        '*      This function updates the surface temperatures, using an
        '*      implicit finite difference method.
        '*
        '*  ================================================================

        Try
            Dim walllayersremaining As Integer
            Dim k, j As Integer
            Dim Wallnodestemp As Integer
            Dim NLw As Integer
            Dim ier As Short
            Dim wallnodeadjust As Integer
            Dim chardensity, moisturecontent, temp As Double
            Dim prop_ku, prop_kl As Double 'W/mK
            Dim char_alpha, char_c, char_fourier As Double
            Dim wood_alpha, wood_c, wood_fourier As Double
            Dim UWoutbiot, LWoutbiot As Double

            moisturecontent = init_moisturecontent

            'moisturecontent = 0.12 'assume 15%
            chardensity = 0.63 * CeilingDensity(room) / (1 + moisturecontent)

            If CLTwallpercent >= 0 Then
                NLw = WallThickness(room) / 1000 / Lamella 'number of lamella - in two places also in main_program2

                walllayersremaining = NLw - Lamella2 / Lamella + 1

                If walllayersremaining = 0 Then
                    If walllayersremaining = 0 Then Dim Message As String = CStr(tim(stepcount, 1)) & " sec. Wall layer at " & Format(Lamella2, "0.000") & " m delaminates. Simulation terminated. "

                    flagstop = 1 'do not continue
                    Exit Sub
                End If

                Wallnodestemp = (Wallnodes - 1) * (walllayersremaining) + 1 'remove one layer and recalc number of nodes
                wallnodeadjust = (Wallnodes - 1) * NLw + 1 - Wallnodestemp

                'Find DeltaX
                WallDeltaX(room) = walllayersremaining * Lamella / (Wallnodestemp - 1)
            Else
                'clt model is on but not for wall
                NLw = 1
                walllayersremaining = 1
                Wallnodestemp = Wallnodes
                wallnodeadjust = 0
            End If

            Dim UWallNodeTemp(Wallnodestemp) As Double
            Dim LWallNodeTemp(Wallnodestemp) As Double
            'Dim UWallNodeStatus(Wallnodestemp) As Integer  'array to identify charred element
            'Dim LWallNodeStatus(Wallnodestemp) As Integer  'array to identify charred element

            For k = 1 To Wallnodestemp
                UWallNodeTemp(k) = UWallNode(room, k + wallnodeadjust, i)
                LWallNodeTemp(k) = LWallNode(room, k + wallnodeadjust, i)

                'array to identify charred element
                If UWallNodeTemp(k) > 300 + 273 Then UWallNodeStatus(k) = 1
                'array to identify charred element
                If LWallNodeTemp(k) > 300 + 273 Then LWallNodeStatus(k) = 1
            Next

            Dim UWallNodeUnExposed As Double = UWallNodeTemp(Wallnodestemp)
            Dim LWallNodeUnExposed As Double = LWallNodeTemp(Wallnodestemp)
            Dim LWallNodeExposed As Double = LWallNodeTemp(1)
            Dim UWallNodeExposed As Double = UWallNodeTemp(1)

            prop_ku = wood_props_k(UWallNodeTemp(Wallnodestemp), UWallNodeStatus(Wallnodestemp), UWallNodeMaxTemp(Wallnodestemp))
            'If UWallNodeTemp(Wallnodestemp) <= 473 Then
            '    prop_ku = 0.285
            '    'prop_ku = WallConductivity(room)
            'ElseIf UWallNodeTemp(Wallnodestemp) <= 663 Then
            '    prop_ku = -0.617 + 0.0038 * UWallNodeTemp(Wallnodestemp) - 0.000004 * UWallNodeTemp(Wallnodestemp) ^ 2
            'Else
            '    prop_ku = 0.04429 + 0.0001477 * UWallNodeTemp(Wallnodestemp)
            'End If

            'Find Biot Numbers -exterior side
            UWoutbiot = OutsideConvCoeff * WallDeltaX(room) / prop_ku

            prop_kl = wood_props_k(LWallNodeTemp(Wallnodestemp), LWallNodeStatus(Wallnodestemp), LWallNodeMaxTemp(Wallnodestemp))
            'If LWallNodeTemp(Wallnodestemp) <= 473 Then
            '    prop_kl = 0.285
            '    'prop_kl = WallConductivity(room)
            'ElseIf LWallNodeTemp(Wallnodestemp) <= 663 Then
            '    prop_kl = -0.617 + 0.0038 * LWallNodeTemp(Wallnodestemp) - 0.000004 * LWallNodeTemp(Wallnodestemp) ^ 2
            'Else
            '    prop_kl = 0.04429 + 0.0001477 * LWallNodeTemp(Wallnodestemp)
            'End If
            LWoutbiot = OutsideConvCoeff * WallDeltaX(room) / prop_kl


            Dim LW(Wallnodestemp, Wallnodestemp) As Double
            Dim UW(Wallnodestemp, Wallnodestemp) As Double
            Dim WX(Wallnodestemp, 1) As Double
            Dim LX(Wallnodestemp, 1) As Double

            'Find Fourier Numbers -exterior side
            If UWallNodeStatus(Wallnodestemp) = 1 Then
                'char
                temp = UWallNodeTemp(Wallnodestemp) - 273
                char_c = 714 + 2.3 * temp - 0.0008 * temp ^ 2 - 0.00000037 * temp ^ 3
                char_alpha = prop_ku / (char_c * chardensity)
                char_fourier = char_alpha * Timestep / (WallDeltaX(room)) ^ 2

                'exterior boundary conditions
                WX(Wallnodestemp, 1) = 2 * char_fourier * UWoutbiot * ((ExteriorTemp - UWallNodeUnExposed) - Surface_Emissivity(2, room) / OutsideConvCoeff * StefanBoltzmann * (UWallNodeUnExposed ^ 4 - ExteriorTemp ^ 4)) + UWallNodeUnExposed
                UW(Wallnodestemp, Wallnodestemp - 1) = -2 * char_fourier
                UW(Wallnodestemp, Wallnodestemp) = 1 + 2 * char_fourier
                UW(1, 1) = 1 + 2 * char_fourier
                UW(1, 2) = -2 * char_fourier
            Else
                'wood
                wood_c = 101.3 + 3.867 * UWallNodeTemp(Wallnodestemp) 'dry wood
                wood_c = (wood_c + 4187 * moisturecontent) / (1 + moisturecontent) + (23.55 * (UWallNodeTemp(Wallnodestemp) - 273) - 1326 * moisturecontent + 2417) * moisturecontent
                wood_alpha = prop_ku / (wood_c * WallDensity(room))
                wood_fourier = wood_alpha * Timestep / (WallDeltaX(room)) ^ 2

                'exterior boundary conditions
                WX(Wallnodestemp, 1) = 2 * wood_fourier * UWoutbiot * ((ExteriorTemp - UWallNodeUnExposed) - Surface_Emissivity(2, room) / OutsideConvCoeff * StefanBoltzmann * (UWallNodeUnExposed ^ 4 - ExteriorTemp ^ 4)) + UWallNodeUnExposed
                UW(Wallnodestemp, Wallnodestemp - 1) = -2 * wood_fourier
                UW(Wallnodestemp, Wallnodestemp) = 1 + 2 * wood_fourier
                UW(1, 1) = 1 + 2 * wood_fourier
                UW(1, 2) = -2 * wood_fourier
            End If

            If LWallNodeStatus(Wallnodestemp) = 1 Then
                'char
                temp = LWallNodeTemp(Wallnodestemp) - 273
                char_c = 714 + 2.3 * temp - 0.0008 * temp ^ 2 - 0.00000037 * temp ^ 3
                char_alpha = prop_kl / (char_c * chardensity)
                char_fourier = char_alpha * Timestep / (WallDeltaX(room)) ^ 2

                'exterior boundary conditions
                LX(Wallnodestemp, 1) = 2 * char_fourier * LWoutbiot * ((ExteriorTemp - LWallNodeUnExposed) - Surface_Emissivity(3, room) / OutsideConvCoeff * StefanBoltzmann * (LWallNodeUnExposed ^ 4 - ExteriorTemp ^ 4)) + LWallNodeUnExposed
                LW(Wallnodestemp, Wallnodestemp - 1) = -2 * char_fourier
                LW(Wallnodestemp, Wallnodestemp) = 1 + 2 * char_fourier
                LW(1, 1) = 1 + 2 * char_fourier
                LW(1, 2) = -2 * char_fourier
            Else
                'wood
                wood_c = 101.3 + 3.867 * LWallNodeTemp(Wallnodestemp) 'dry
                wood_c = (wood_c + 4187 * moisturecontent) / (1 + moisturecontent) + (23.55 * (LWallNodeTemp(Wallnodestemp) - 273) - 1326 * moisturecontent + 2417) * moisturecontent 'wet
                wood_alpha = prop_kl / (wood_c * WallDensity(room))
                wood_fourier = wood_alpha * Timestep / (WallDeltaX(room)) ^ 2

                'exterior boundary conditions
                LX(Wallnodestemp, 1) = 2 * wood_fourier * LWoutbiot * ((ExteriorTemp - LWallNodeUnExposed) - Surface_Emissivity(3, room) / OutsideConvCoeff * StefanBoltzmann * (LWallNodeUnExposed ^ 4 - ExteriorTemp ^ 4)) + LWallNodeUnExposed
                LW(Wallnodestemp, Wallnodestemp - 1) = -2 * wood_fourier
                LW(Wallnodestemp, Wallnodestemp) = 1 + 2 * wood_fourier
                LW(1, 1) = 1 + 2 * wood_fourier
                LW(1, 2) = -2 * wood_fourier
            End If

            prop_ku = wood_props_k(UWallNodeTemp(2), UWallNodeStatus(2), UWallNodeMaxTemp(2))
            'If UWallNodeTemp(2) <= 473 Then
            '    prop_ku = 0.285
            '    'prop_ku = WallConductivity(room)
            'ElseIf UWallNodeTemp(2) <= 663 Then
            '    prop_ku = -0.617 + 0.0038 * UWallNodeTemp(2) - 0.000004 * UWallNodeTemp(2) ^ 2
            'Else
            '    prop_ku = 0.04429 + 0.0001477 * UWallNodeTemp(2)
            'End If

            If UWallNodeStatus(2) = 1 Then 'char
                temp = UWallNodeTemp(2) - 273
                char_c = 714 + 2.3 * temp - 0.0008 * temp ^ 2 - 0.00000037 * temp ^ 3
                char_alpha = prop_ku / (char_c * chardensity)
                char_fourier = char_alpha * Timestep / (WallDeltaX(room)) ^ 2
                UW(1, 1) = 1 + 2 * char_fourier
                UW(1, 2) = -2 * char_fourier
                'interior boundary conditions
                WX(1, 1) = -2 * QUpperWall(room, i) * 1000 * char_fourier * WallDeltaX(room) / prop_ku + UWallNodeExposed
                For k = 2 To Wallnodestemp - 1
                    WX(k, 1) = UWallNodeTemp(k)
                Next k
            Else
                'wood
                wood_c = 101.3 + 3.867 * UWallNodeTemp(2)
                wood_c = (wood_c + 4187 * moisturecontent) / (1 + moisturecontent) + (23.55 * (UWallNodeTemp(2) - 273) - 1326 * moisturecontent + 2417) * moisturecontent
                wood_alpha = prop_ku / (wood_c * WallDensity(room))
                wood_fourier = wood_alpha * Timestep / (WallDeltaX(room)) ^ 2
                UW(1, 1) = 1 + 2 * wood_fourier
                UW(1, 2) = -2 * wood_fourier
                'interior boundary conditions
                WX(1, 1) = -2 * QUpperWall(room, i) * 1000 * wood_fourier * WallDeltaX(room) / prop_ku + UWallNodeExposed
                For k = 2 To Wallnodestemp - 1
                    WX(k, 1) = UWallNodeTemp(k)
                Next k
            End If

            prop_kl = wood_props_k(LWallNodeTemp(Wallnodestemp), LWallNodeStatus(Wallnodestemp), LWallNodeMaxTemp(Wallnodestemp))
            'If LWallNodeTemp(2) <= 473 Then
            '    prop_kl = 0.285
            '    'prop_kl = WallConductivity(room)
            'ElseIf LWallNodeTemp(2) <= 663 Then
            '    prop_kl = -0.617 + 0.0038 * LWallNodeTemp(2) - 0.000004 * LWallNodeTemp(2) ^ 2
            'Else
            '    prop_kl = 0.04429 + 0.0001477 * LWallNodeTemp(2)
            'End If

            If LWallNodeStatus(2) = 1 Then 'char
                temp = LWallNodeTemp(2) - 273
                char_c = 714 + 2.3 * temp - 0.0008 * temp ^ 2 - 0.00000037 * temp ^ 3
                char_alpha = prop_kl / (char_c * chardensity)
                char_fourier = char_alpha * Timestep / (WallDeltaX(room)) ^ 2
                LW(1, 1) = 1 + 2 * char_fourier
                LW(1, 2) = -2 * char_fourier
                'interior boundary conditions
                LX(1, 1) = -2 * QLowerWall(room, i) * 1000 * char_fourier * WallDeltaX(room) / prop_kl + LWallNodeExposed
                For k = 2 To Wallnodestemp - 1
                    LX(k, 1) = LWallNodeTemp(k)
                Next k
            Else
                'wood
                wood_c = 101.3 + 3.867 * LWallNodeTemp(2)
                wood_c = (wood_c + 4187 * moisturecontent) / (1 + moisturecontent) + (23.55 * (LWallNodeTemp(2) - 273) - 1326 * moisturecontent + 2417) * moisturecontent
                wood_alpha = prop_kl / (wood_c * WallDensity(room))
                wood_fourier = wood_alpha * Timestep / (WallDeltaX(room)) ^ 2
                LW(1, 1) = 1 + 2 * wood_fourier
                LW(1, 2) = -2 * wood_fourier
                'interior boundary conditions
                LX(1, 1) = -2 * QLowerWall(room, i) * 1000 * wood_fourier * WallDeltaX(room) / prop_kl + LWallNodeExposed
                For k = 2 To Wallnodestemp - 1
                    LX(k, 1) = LWallNodeTemp(k)
                Next k
            End If

            'inside nodes
            k = 2
            For j = 2 To Wallnodestemp - 1

                prop_ku = wood_props_k(UWallNodeTemp(j + 1), UWallNodeStatus(j + 1), UWallNodeMaxTemp(j + 1))
                'If UWallNodeTemp(j + 1) <= 473 Then
                '    prop_ku = 0.285
                '    'prop_ku = WallConductivity(room)
                'ElseIf UWallNodeTemp(j + 1) <= 663 Then
                '    prop_ku = -0.617 + 0.0038 * UWallNodeTemp(j + 1) - 0.000004 * UWallNodeTemp(j + 1) ^ 2
                'Else
                '    prop_ku = 0.04429 + 0.0001477 * UWallNodeTemp(j + 1)
                'End If

                If UWallNodeStatus(j + 1) = 1 Then
                    temp = UWallNodeTemp(j + 1) - 273
                    char_c = 714 + 2.3 * temp - 0.0008 * temp ^ 2 - 0.00000037 * temp ^ 3
                    char_alpha = prop_ku / (char_c * chardensity)
                    char_fourier = char_alpha * Timestep / (WallDeltaX(room)) ^ 2
                    UW(j, k - 1) = -char_fourier
                    UW(j, k) = 1 + 2 * char_fourier
                    UW(j, k + 1) = -char_fourier
                Else
                    wood_c = 101.3 + 3.867 * UWallNodeTemp(j + 1)
                    wood_c = (wood_c + 4187 * moisturecontent) / (1 + moisturecontent) + (23.55 * (UWallNodeTemp(j + 1) - 273) - 1326 * moisturecontent + 2417) * moisturecontent
                    wood_alpha = prop_ku / (wood_c * WallDensity(room))
                    wood_fourier = wood_alpha * Timestep / (WallDeltaX(room)) ^ 2
                    UW(j, k - 1) = -wood_fourier
                    UW(j, k) = 1 + 2 * wood_fourier
                    UW(j, k + 1) = -wood_fourier
                End If

                k = k + 1
            Next j

            k = 2
            For j = 2 To Wallnodestemp - 1
                prop_kl = wood_props_k(LWallNodeTemp(j + 1), LWallNodeStatus(j + 1), LWallNodeMaxTemp(j + 1))
                'If LWallNodeTemp(j + 1) <= 473 Then
                '    prop_kl = 0.285
                '    'prop_kl = WallConductivity(room)
                'ElseIf LWallNodeTemp(j + 1) <= 663 Then
                '    prop_kl = -0.617 + 0.0038 * LWallNodeTemp(j + 1) - 0.000004 * LWallNodeTemp(j + 1) ^ 2
                'Else
                '    prop_kl = 0.04429 + 0.0001477 * LWallNodeTemp(j + 1)
                'End If

                If LWallNodeStatus(j + 1) = 1 Then
                    temp = LWallNodeTemp(j + 1) - 273
                    char_c = 714 + 2.3 * temp - 0.0008 * temp ^ 2 - 0.00000037 * temp ^ 3
                    char_alpha = prop_kl / (char_c * chardensity)
                    char_fourier = char_alpha * Timestep / (WallDeltaX(room)) ^ 2
                    LW(j, k - 1) = -char_fourier
                    LW(j, k) = 1 + 2 * char_fourier
                    LW(j, k + 1) = -char_fourier
                Else
                    wood_c = 101.3 + 3.867 * LWallNodeTemp(j + 1)
                    wood_c = (wood_c + 4187 * moisturecontent) / (1 + moisturecontent) + (23.55 * (LWallNodeTemp(j + 1) - 273) - 1326 * moisturecontent + 2417) * moisturecontent
                    wood_alpha = prop_kl / (wood_c * WallDensity(room))
                    wood_fourier = wood_alpha * Timestep / (WallDeltaX(room)) ^ 2
                    LW(j, k - 1) = -wood_fourier
                    LW(j, k) = 1 + 2 * wood_fourier
                    LW(j, k + 1) = -wood_fourier
                End If

                k = k + 1
            Next j

            If frmOptions1.optLUdecom.Checked = True Then
                'find surface temperatures for the next timestep
                'using method of LU decomposition (preferred)
                Call MatSol(UW, WX, Wallnodestemp) 'upper wall
                Call MatSol(LW, LX, Wallnodestemp) 'lower wall
            Else
                'find surface temperatures for the next timestep
                'using method of Gauss-Jordan elimination
                Call LINEAR2(Wallnodestemp, UW, WX, ier)
                Call LINEAR2(Wallnodestemp, LW, LX, ier)
                If ier = 1 Then MsgBox("singular matrix in implicit_surface_temps_CLTW")
            End If

            For j = 1 To Wallnodestemp
                UWallNode(room, j + wallnodeadjust, i + 1) = WX(j, 1)
                LWallNode(room, j + wallnodeadjust, i + 1) = LX(j, 1)
                If WX(j, 1) > UWallNodeMaxTemp(j + wallnodeadjust) Then UWallNodeMaxTemp(j + wallnodeadjust) = WX(j, 1)
                If LX(j, 1) > LWallNodeMaxTemp(j + wallnodeadjust) Then LWallNodeMaxTemp(j + wallnodeadjust) = LX(j, 1)
            Next j
            For j = 1 To wallnodeadjust
                UWallNode(room, j, i + 1) = chartemp + 273 + 1
                LWallNode(room, j, i + 1) = chartemp + 273 + 1
            Next

            'store surface temps at next timestep in another array
            Upperwalltemp(room, i + 1) = UWallNode(room, 1 + wallnodeadjust, i + 1)
            LowerWallTemp(room, i + 1) = LWallNode(room, 1 + wallnodeadjust, i + 1)
            UnexposedUpperwalltemp(room, i + 1) = UWallNode(room, Wallnodestemp, i + 1)
            UnexposedLowerwalltemp(room, i + 1) = LWallNode(room, Wallnodestemp, i + 1)

            Erase LX
            Erase UW
            Erase WX
            Erase LW

        Catch ex As Exception
            MsgBox(Err.Description & " Line " & Err.Erl, MsgBoxStyle.Exclamation, "Exception in Implicit_Surface_Temps_CLTW_Char ")
        End Try
    End Sub

    Sub Implicit_Surface_Temps_CLTC_kinetic(ByVal room As Integer, ByVal i As Integer, ByRef CeilingNode(,,) As Double)
        '*  ================================================================
        '*      This function updates the ceiling nodal, using an
        '*      implicit finite difference method.
        '*      in conjunction with a kinetic pyrolysis model 
        '*  ================================================================

        Try
            Dim ceilinglayersremaining As Integer
            Dim k, j As Integer
            Dim ceilingnodestemp As Integer
            Dim NLC As Integer
            Dim ier As Short
            Dim ceilingnodeadjust As Integer
            Dim chardensity, moisturecontent, temp As Double
            Dim prop_k As Double 'W/mK
            Dim char_alpha, char_c, char_fourier As Double
            Dim wood_alpha, wood_c, wood_fourier As Double
            Dim Coutbiot As Double
            Dim kwood As Double = 0.285
            kwood = CeilingConductivity(room)

            ' moisturecontent = 0.1 'assume 10% mass fraction
            moisturecontent = init_moisturecontent

            chardensity = 0.63 * CeilingDensity(room) / (1 + moisturecontent)


        Catch ex As Exception
            MsgBox(Err.Description & " Line " & Err.Erl, MsgBoxStyle.Exclamation, "Exception in " & Err.Source)
        End Try
    End Sub
    Sub Implicit_Surface_Temps_CLTC_Char(ByVal room As Integer, ByVal i As Integer, ByRef CeilingNode(,,) As Double)
        '*  ================================================================
        '*      This function updates the surface temperatures, using an
        '*      implicit finite difference method.
        '*
        '*  ================================================================

        Try
            Dim ceilinglayersremaining As Integer
            Dim k, j As Integer
            Dim ceilingnodestemp As Integer
            Dim NLC As Integer
            Dim ier As Short
            Dim ceilingnodeadjust As Integer
            Dim chardensity, moisturecontent, temp As Double
            Dim prop_k As Double 'W/mK
            Dim char_alpha, char_c, char_fourier As Double
            Dim wood_alpha, wood_c, wood_fourier As Double
            Dim Coutbiot As Double

            Dim kwood As Double = 0.285

            'kwood = CeilingConductivity(room)

            'moisturecontent = 0.08 'was 12%
            moisturecontent = init_moisturecontent

            chardensity = 0.63 * CeilingDensity(room) / (1 + moisturecontent)

            If CLTceilingpercent >= 0 Then
                NLC = CeilingThickness(room) / 1000 / Lamella 'number of lamella - in two places also in main_program2

                ceilinglayersremaining = NLC - Lamella1 / Lamella + 1

                If ceilinglayersremaining = 0 Then
                    If ceilinglayersremaining = 0 Then Dim Message As String = CStr(tim(stepcount, 1)) & " sec. Ceiling layer at " & Format(Lamella1, "0.000") & " m delaminates. Simulation terminated. "

                    flagstop = 1 'do not continue
                    Exit Sub
                End If

                ceilingnodestemp = (Ceilingnodes - 1) * (ceilinglayersremaining) + 1 'remove one layer and recalc number of nodes
                ceilingnodeadjust = (Ceilingnodes - 1) * NLC + 1 - ceilingnodestemp

                'Find DeltaX
                CeilingDeltaX(room) = ceilinglayersremaining * Lamella / (ceilingnodestemp - 1)
            Else
                'clt model is on but not for ceiling
                NLC = 1
                ceilinglayersremaining = 1
                ceilingnodestemp = Ceilingnodes
                ceilingnodeadjust = 0
            End If

            Dim CeilingNodeTemp(ceilingnodestemp) As Double
            'Dim CeilingNodeStatus(ceilingnodestemp) As Integer  'array to identify charred element

            For k = 1 To ceilingnodestemp
                CeilingNodeTemp(k) = CeilingNode(room, k + ceilingnodeadjust, i)

                'array to identify charred element
                If CeilingNodeTemp(k) > 300 + 273 Then CeilingNodeStatus(k) = 1 'this is reversible?

            Next

            Dim CeilingNodeUnExposed As Double = CeilingNodeTemp(ceilingnodestemp)
            Dim CeilingNodeExposed As Double = CeilingNodeTemp(1)

            prop_k = wood_props_k(CeilingNodeTemp(ceilingnodestemp), CeilingNodeStatus(ceilingnodestemp), CeilingNodeMaxTemp(ceilingnodestemp))

            'If CeilingNodeTemp(ceilingnodestemp) <= 473 Then 'uses node temp not max temp reached by that node?
            '    prop_k = kwood
            '    prop_k = CeilingConductivity(room)

            'ElseIf CeilingNodeTemp(ceilingnodestemp) <= 663 Then
            '    prop_k = -0.617 + 0.0038 * CeilingNodeTemp(ceilingnodestemp) - 0.000004 * CeilingNodeTemp(ceilingnodestemp) ^ 2
            'Else
            '    prop_k = 0.04429 + 0.0001477 * CeilingNodeTemp(ceilingnodestemp)
            'End If

            'Find Biot Numbers -exterior side
            Coutbiot = OutsideConvCoeff * CeilingDeltaX(room) / prop_k

            Dim UC(ceilingnodestemp, ceilingnodestemp) As Double
            Dim CX(ceilingnodestemp, 1) As Double

            'Find Fourier Numbers -exterior side
            If CeilingNodeStatus(ceilingnodestemp) = 1 Then
                'char
                temp = CeilingNodeTemp(ceilingnodestemp) - 273 'node temp in deg C
                char_c = 714 + 2.3 * temp - 0.0008 * temp ^ 2 - 0.00000037 * temp ^ 3
                char_alpha = prop_k / (char_c * chardensity)
                char_fourier = char_alpha * Timestep / (CeilingDeltaX(room)) ^ 2

                'exterior boundary conditions
                CX(ceilingnodestemp, 1) = 2 * char_fourier * Coutbiot * ((ExteriorTemp - CeilingNodeUnExposed) - Surface_Emissivity(1, room) / OutsideConvCoeff * StefanBoltzmann * (CeilingNodeUnExposed ^ 4 - ExteriorTemp ^ 4)) + CeilingNodeUnExposed
                UC(ceilingnodestemp, ceilingnodestemp - 1) = -2 * char_fourier
                UC(ceilingnodestemp, ceilingnodestemp) = 1 + 2 * char_fourier
                UC(1, 1) = 1 + 2 * char_fourier
                UC(1, 2) = -2 * char_fourier
            Else
                'wood
                wood_c = 101.3 + 3.867 * CeilingNodeTemp(ceilingnodestemp)
                wood_c = (wood_c + 4187 * moisturecontent) / (1 + moisturecontent) + (23.55 * (CeilingNodeTemp(ceilingnodestemp) - 273) - 1326 * moisturecontent + 2417) * moisturecontent

                wood_alpha = prop_k / (wood_c * CeilingDensity(room))
                wood_fourier = wood_alpha * Timestep / (CeilingDeltaX(room)) ^ 2

                'exterior boundary conditions
                CX(ceilingnodestemp, 1) = 2 * wood_fourier * Coutbiot * ((ExteriorTemp - CeilingNodeUnExposed) - Surface_Emissivity(1, room) / OutsideConvCoeff * StefanBoltzmann * (CeilingNodeUnExposed ^ 4 - ExteriorTemp ^ 4)) + CeilingNodeUnExposed
                UC(ceilingnodestemp, ceilingnodestemp - 1) = -2 * wood_fourier
                UC(ceilingnodestemp, ceilingnodestemp) = 1 + 2 * wood_fourier
                UC(1, 1) = 1 + 2 * wood_fourier
                UC(1, 2) = -2 * wood_fourier
            End If

            'exposed side
            prop_k = wood_props_k(CeilingNodeTemp(2), CeilingNodeStatus(2), CeilingNodeMaxTemp(2))

            'If CeilingNodeTemp(2) <= 473 Then
            '    prop_k = kwood
            'ElseIf CeilingNodeTemp(2) <= 663 Then
            '    prop_k = -0.617 + 0.0038 * CeilingNodeTemp(2) - 0.000004 * CeilingNodeTemp(2) ^ 2
            'Else
            '    prop_k = 0.04429 + 0.0001477 * CeilingNodeTemp(2)
            'End If

            If CeilingNodeStatus(2) = 1 Then 'char
                temp = CeilingNodeTemp(2) - 273 'node temp in deg C
                char_c = 714 + 2.3 * temp - 0.0008 * temp ^ 2 - 0.00000037 * temp ^ 3
                char_alpha = prop_k / (char_c * chardensity)
                char_fourier = char_alpha * Timestep / (CeilingDeltaX(room)) ^ 2
                UC(1, 1) = 1 + 2 * char_fourier
                UC(1, 2) = -2 * char_fourier
                'interior boundary conditions
                CX(1, 1) = -2 * QCeiling(room, i) * 1000 * char_fourier * CeilingDeltaX(room) / prop_k + CeilingNodeExposed
                For k = 2 To ceilingnodestemp - 1
                    CX(k, 1) = CeilingNodeTemp(k)
                Next k
            Else
                'wood
                wood_c = 101.3 + 3.867 * CeilingNodeTemp(2)
                wood_c = (wood_c + 4187 * moisturecontent) / (1 + moisturecontent) + (23.55 * (CeilingNodeTemp(2) - 273) - 1326 * moisturecontent + 2417) * moisturecontent

                wood_alpha = prop_k / (wood_c * CeilingDensity(room))
                wood_fourier = wood_alpha * Timestep / (CeilingDeltaX(room)) ^ 2
                UC(1, 1) = 1 + 2 * wood_fourier
                UC(1, 2) = -2 * wood_fourier
                'interior boundary conditions
                CX(1, 1) = -2 * QCeiling(room, i) * 1000 * wood_fourier * CeilingDeltaX(room) / prop_k + CeilingNodeExposed
                For k = 2 To ceilingnodestemp - 1
                    CX(k, 1) = CeilingNodeTemp(k)
                Next k
            End If


            'inside nodes
            k = 2
            For j = 2 To ceilingnodestemp - 1

                prop_k = wood_props_k(CeilingNodeTemp(j + 1), CeilingNodeStatus(j + 1), CeilingNodeMaxTemp(j + 1))

                'If CeilingNodeTemp(j + 1) <= 473 Then
                '    prop_k = kwood
                'ElseIf CeilingNodeTemp(j + 1) <= 663 Then
                '    prop_k = -0.617 + 0.0038 * CeilingNodeTemp(j + 1) - 0.000004 * CeilingNodeTemp(j + 1) ^ 2
                'Else
                '    prop_k = 0.04429 + 0.0001477 * CeilingNodeTemp(j + 1)
                'End If

                If CeilingNodeStatus(j + 1) = 1 Then
                    temp = CeilingNodeTemp(j + 1) - 273
                    char_c = 714 + 2.3 * temp - 0.0008 * temp ^ 2 - 0.00000037 * temp ^ 3
                    char_alpha = prop_k / (char_c * chardensity)
                    char_fourier = char_alpha * Timestep / (CeilingDeltaX(room)) ^ 2
                    UC(j, k - 1) = -char_fourier
                    UC(j, k) = 1 + 2 * char_fourier
                    UC(j, k + 1) = -char_fourier
                Else
                    wood_c = 101.3 + 3.867 * CeilingNodeTemp(j + 1)
                    wood_c = (wood_c + 4187 * moisturecontent) / (1 + moisturecontent) + (23.55 * (CeilingNodeTemp(j + 1) - 273) - 1326 * moisturecontent + 2417) * moisturecontent

                    wood_alpha = prop_k / (wood_c * CeilingDensity(room))
                    wood_fourier = wood_alpha * Timestep / (CeilingDeltaX(room)) ^ 2
                    UC(j, k - 1) = -wood_fourier
                    UC(j, k) = 1 + 2 * wood_fourier
                    UC(j, k + 1) = -wood_fourier
                End If

                k = k + 1
            Next j

            If frmOptions1.optLUdecom.Checked = True Then
                'find surface temperatures for the next timestep
                'using method of LU decomposition (preferred)
                Call MatSol(UC, CX, ceilingnodestemp) 'ceiling

            Else
                'find surface temperatures for the next timestep
                'using method of Gauss-Jordan elimination
                Call LINEAR2(ceilingnodestemp, UC, CX, ier)

                If ier = 1 Then MsgBox("singular matrix in implicit_surface_temps_CLTC_char")
            End If

            For j = 1 To ceilingnodestemp
                CeilingNode(room, j + ceilingnodeadjust, i + 1) = CX(j, 1)
                If CX(j, 1) > CeilingNodeMaxTemp(j + ceilingnodeadjust) Then CeilingNodeMaxTemp(j + ceilingnodeadjust) = CX(j, 1)
            Next j
            For j = 1 To ceilingnodeadjust
                CeilingNode(room, j, i + 1) = chartemp + 273 + 1
            Next

            'store surface temps at next timestep in another array
            CeilingTemp(room, i + 1) = CeilingNode(room, 1 + ceilingnodeadjust, i + 1)

            UnexposedCeilingtemp(room, i + 1) = CeilingNode(room, ceilingnodestemp, i + 1)

            Erase CX
            Erase UC

        Catch ex As Exception
            MsgBox(Err.Description & " Line " & Err.Erl, MsgBoxStyle.Exclamation, "Exception in Implicit_Surface_Temps_CLTW_Char")
        End Try
    End Sub

    Sub Implicit_Surface_Temps_floor(ByVal room As Integer, ByVal i As Integer, ByRef FloorNode(,,) As Double)
        '*  ================================================================
        '*      This function updates the surface temperatures, using an
        '*      implicit finite difference method.
        '*
        '*  ================================================================

        Try

            Dim k, j As Integer
            Dim ier As Short
            Dim LF(Floornodes, Floornodes) As Double
            Dim FX(Floornodes, 1) As Double
            Dim FloorFourier2, FloorOutsideBiot2 As Double

            If HaveFloorSubstrate(room) = False Then
                ReDim LF(Floornodes, Floornodes)
                ReDim FX(Floornodes, 1)
            Else
                ReDim LF(2 * Floornodes - 1, 2 * Floornodes - 1)
                ReDim FX(2 * Floornodes - 1, 1)

                FloorFourier2 = FloorSubConductivity(room) * Timestep / (((FloorSubThickness(room) / 1000) / (Floornodes - 1)) ^ 2 * FloorSubDensity(room) * FloorSubSpecificHeat(room))
                FloorOutsideBiot2 = OutsideConvCoeff * ((FloorSubThickness(room) / 1000) / (Floornodes - 1)) / FloorSubConductivity(room)
            End If

            LF(1, 1) = 1 + 2 * FloorFourier(room)
            LF(1, 2) = -2 * FloorFourier(room)

            If HaveFloorSubstrate(room) = False Then
                k = 2
                For j = 2 To Floornodes - 1
                    LF(j, k - 1) = -FloorFourier(room)
                    LF(j, k) = 1 + 2 * FloorFourier(room)
                    LF(j, k + 1) = -FloorFourier(room)
                    k = k + 1
                Next j
            Else
                k = 2
                For j = 2 To Floornodes - 1
                    LF(j, k - 1) = -FloorFourier(room)
                    LF(j, k) = 1 + 2 * FloorFourier(room)
                    LF(j, k + 1) = -FloorFourier(room)
                    k = k + 1
                Next j

                j = Floornodes
                LF(j, k - 1) = -FloorFourier(room)
                LF(j, k) = 1 + FloorFourier(room) + FloorFourier2
                LF(j, k + 1) = -FloorFourier2
                k = k + 1

                For j = Floornodes + 1 To 2 * Floornodes - 2
                    LF(j, k - 1) = -FloorFourier2
                    LF(j, k) = 1 + 2 * FloorFourier2
                    LF(j, k + 1) = -FloorFourier2
                    k = k + 1
                Next j
            End If

            If HaveFloorSubstrate(room) = False Then
                LF(Floornodes, Floornodes - 1) = -2 * FloorFourier(room)
                LF(Floornodes, Floornodes) = 1 + 2 * FloorFourier(room)
            Else
                LF(2 * Floornodes - 1, 2 * Floornodes - 2) = -2 * FloorFourier2
                LF(2 * Floornodes - 1, 2 * Floornodes - 1) = 1 + 2 * FloorFourier2
            End If

            FX(1, 1) = -2 * QFloor(room, i) * 1000 * FloorFourier(room) * FloorDeltaX(room) / FloorConductivity(room) + GLOBAL_Renamed.FloorNode(room, 1, i)
            If HaveFloorSubstrate(room) = False Then
                For k = 2 To Floornodes - 1
                    FX(k, 1) = GLOBAL_Renamed.FloorNode(room, k, i)
                Next k
            Else
                For k = 2 To 2 * Floornodes - 2
                    FX(k, 1) = GLOBAL_Renamed.FloorNode(room, k, i)
                Next k
            End If

            If HaveFloorSubstrate(room) = False Then
                FX(Floornodes, 1) = 2 * FloorFourier(room) * FloorOutsideBiot(room) * ((ExteriorTemp - GLOBAL_Renamed.FloorNode(room, Floornodes, i)) - Surface_Emissivity(4, room) / OutsideConvCoeff * StefanBoltzmann * (GLOBAL_Renamed.FloorNode(room, Floornodes, i) ^ 4 - ExteriorTemp ^ 4)) + GLOBAL_Renamed.FloorNode(room, Floornodes, i)
            Else
                FX(2 * Floornodes - 1, 1) = 2 * FloorFourier2 * FloorOutsideBiot2 * ((ExteriorTemp - GLOBAL_Renamed.FloorNode(room, 2 * Floornodes - 1, i)) - Surface_Emissivity(4, room) / OutsideConvCoeff * StefanBoltzmann * (GLOBAL_Renamed.FloorNode(room, 2 * Floornodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + GLOBAL_Renamed.FloorNode(room, 2 * Floornodes - 1, i)
            End If

            If frmOptions1.optLUdecom.Checked = True Then
                'find surface temperatures for the next timestep
                'using method of LU decomposition (preferred)
                If HaveFloorSubstrate(room) = False Then
                    Call MatSol(LF, FX, Floornodes) 'floor
                Else
                    Call MatSol(LF, FX, 2 * Floornodes - 1)
                End If
            Else
                'find surface temperatures for the next timestep
                'using method of Gauss-Jordan elimination
                If HaveFloorSubstrate(room) = False Then
                    Call LINEAR2(Floornodes, LF, FX, ier)
                Else
                    Call LINEAR2(2 * Floornodes - 1, LF, FX, ier)
                End If
                If ier = 1 Then MsgBox("singular matrix in implicit_surface_temps_floor")
            End If

            If HaveFloorSubstrate(room) = False Then
                For j = 1 To Floornodes
                    GLOBAL_Renamed.FloorNode(room, j, i + 1) = FX(j, 1)
                Next j
            Else
                For j = 1 To 2 * Floornodes - 1
                    GLOBAL_Renamed.FloorNode(room, j, i + 1) = FX(j, 1)
                Next j
            End If

            FloorTemp(room, i + 1) = GLOBAL_Renamed.FloorNode(room, 1, i + 1)

            If HaveFloorSubstrate(room) = True Then
                UnexposedFloortemp(room, i + 1) = GLOBAL_Renamed.FloorNode(room, 2 * Floornodes - 1, i + 1)
            Else
                UnexposedFloortemp(room, i + 1) = GLOBAL_Renamed.FloorNode(room, Floornodes, i + 1)
            End If

            Erase LF
            Erase FX

        Catch ex As Exception
            MsgBox(Err.Description & " Line " & Err.Erl, MsgBoxStyle.Exclamation, "Exception in Implicit_Surface_Temp_floor ")
        End Try
    End Sub
    Sub Implicit_Surface_Temps2(ByVal room As Integer, ByVal wall As Integer, ByVal ceiling As Integer, ByVal i As Integer, ByRef UWallNode(,,) As Double, ByRef CeilingNode(,,) As Double, ByRef LWallNode(,,) As Double, ByRef FloorNode(,,) As Double)
        '*  ================================================================
        '*      This function updates the surface temperatures, using an
        '*      implicit finite difference method.
        '*
        '*      Two layer composite boundary
        '*
        '*      Revised: Colleen Wade 3 June 1997
        '*  ================================================================

        Dim j, k, ier As Short
        Dim CeilingOutsideBiot2, FloorFourier2, WallFourier2, CeilingFourier2, WallOutsideBiot2, FloorOutsideBiot2 As Double
        Dim UW(2 * Wallnodes - 1, 2 * Wallnodes - 1) As Double
        Dim UC(2 * Ceilingnodes - 1, 2 * Ceilingnodes - 1) As Double
        Dim wx(2 * Wallnodes - 1, 1) As Double
        Dim cx(2 * Ceilingnodes - 1, 1) As Double
        Dim Lf(Floornodes, Floornodes) As Double
        Dim FX(Floornodes, 1) As Double
        Dim LX(2 * Wallnodes - 1, 1) As Double
        Dim LW(2 * Wallnodes - 1, 2 * Wallnodes - 1) As Double
        Dim emmissivity2 As Double = 0.8

        If HaveFloorSubstrate(room) = False Then
            ReDim Lf(Floornodes, Floornodes)
            ReDim FX(Floornodes, 1)
        Else
            ReDim Lf(2 * Floornodes - 1, 2 * Floornodes - 1)
            ReDim FX(2 * Floornodes - 1, 1)
            FloorFourier2 = FloorSubConductivity(room) * Timestep / (((FloorSubThickness(room) / 1000) / (Floornodes - 1)) ^ 2 * FloorSubDensity(room) * FloorSubSpecificHeat(room))
            FloorOutsideBiot2 = OutsideConvCoeff * ((FloorSubThickness(room) / 1000) / (Floornodes - 1)) / FloorSubConductivity(room)
        End If

        On Error GoTo temp_error_handler

        WallFourier2 = WallSubConductivity(room) * Timestep / (((WallSubThickness(room) / 1000) / (Wallnodes - 1)) ^ 2 * WallSubDensity(room) * WallSubSpecificHeat(room))
        CeilingFourier2 = CeilingSubConductivity(room) * Timestep / (((CeilingSubThickness(room) / 1000) / (Ceilingnodes - 1)) ^ 2 * CeilingSubDensity(room) * CeilingSubSpecificHeat(room))

        WallOutsideBiot2 = OutsideConvCoeff * ((WallSubThickness(room) / 1000) / (Wallnodes - 1)) / WallSubConductivity(room)
        CeilingOutsideBiot2 = OutsideConvCoeff * ((CeilingSubThickness(room) / 1000) / (Ceilingnodes - 1)) / CeilingSubConductivity(room)

        UW(1, 1) = 1 + 2 * WallFourier(room)
        UW(1, 2) = -2 * WallFourier(room)
        LW(1, 1) = 1 + 2 * WallFourier(room)
        LW(1, 2) = -2 * WallFourier(room)
        UC(1, 1) = 1 + 2 * CeilingFourier(room)
        UC(1, 2) = -2 * CeilingFourier(room)
        Lf(1, 1) = 1 + 2 * FloorFourier(room)
        Lf(1, 2) = -2 * FloorFourier(room)

        If HaveFloorSubstrate(room) = False Then
            k = 2
            For j = 2 To Floornodes - 1
                Lf(j, k - 1) = -FloorFourier(room)
                Lf(j, k) = 1 + 2 * FloorFourier(room)
                Lf(j, k + 1) = -FloorFourier(room)
                k = k + 1
            Next j
        Else
            k = 2
            For j = 2 To Floornodes - 1
                Lf(j, k - 1) = -FloorFourier(room)
                Lf(j, k) = 1 + 2 * FloorFourier(room)
                Lf(j, k + 1) = -FloorFourier(room)
                k = k + 1
            Next j

            j = Floornodes
            Lf(j, k - 1) = -FloorFourier(room)
            Lf(j, k) = 1 + FloorFourier(room) + FloorFourier2
            Lf(j, k + 1) = -FloorFourier2
            k = k + 1

            For j = Floornodes + 1 To 2 * Floornodes - 2
                Lf(j, k - 1) = -FloorFourier2
                Lf(j, k) = 1 + 2 * FloorFourier2
                Lf(j, k + 1) = -FloorFourier2
                k = k + 1
            Next j
        End If


        k = 2
        For j = 2 To Wallnodes - 1
            UW(j, k - 1) = -WallFourier(room)
            UW(j, k) = 1 + 2 * WallFourier(room)
            UW(j, k + 1) = -WallFourier(room)
            LW(j, k - 1) = -WallFourier(room)
            LW(j, k) = 1 + 2 * WallFourier(room)
            LW(j, k + 1) = -WallFourier(room)
            k = k + 1
        Next j

        j = Wallnodes
        UW(j, k - 1) = -WallFourier(room)
        UW(j, k) = 1 + WallFourier(room) + WallFourier2
        UW(j, k + 1) = -WallFourier2
        LW(j, k - 1) = -WallFourier(room)
        LW(j, k) = 1 + WallFourier(room) + WallFourier2
        LW(j, k + 1) = -WallFourier2
        k = k + 1

        For j = Wallnodes + 1 To 2 * Wallnodes - 2
            UW(j, k - 1) = -WallFourier2
            UW(j, k) = 1 + 2 * WallFourier2
            UW(j, k + 1) = -WallFourier2
            LW(j, k - 1) = -WallFourier2
            LW(j, k) = 1 + 2 * WallFourier2
            LW(j, k + 1) = -WallFourier2
            k = k + 1
        Next j

        k = 2
        For j = 2 To Ceilingnodes - 1
            UC(j, k - 1) = -CeilingFourier(room)
            UC(j, k) = 1 + 2 * CeilingFourier(room)
            UC(j, k + 1) = -CeilingFourier(room)
            k = k + 1
        Next j

        j = Ceilingnodes
        UC(j, k - 1) = -CeilingFourier(room)
        UC(j, k) = 1 + CeilingFourier(room) + CeilingFourier2
        UC(j, k + 1) = -CeilingFourier2
        k = k + 1

        For j = Ceilingnodes + 1 To 2 * Ceilingnodes - 2
            UC(j, k - 1) = -CeilingFourier2
            UC(j, k) = 1 + 2 * CeilingFourier2
            UC(j, k + 1) = -CeilingFourier2
            k = k + 1
        Next j

        UW(2 * Wallnodes - 1, 2 * Wallnodes - 2) = -2 * WallFourier2
        UW(2 * Wallnodes - 1, 2 * Wallnodes - 1) = 1 + 2 * WallFourier2
        LW(2 * Wallnodes - 1, 2 * Wallnodes - 2) = -2 * WallFourier2
        LW(2 * Wallnodes - 1, 2 * Wallnodes - 1) = 1 + 2 * WallFourier2
        UC(2 * Ceilingnodes - 1, 2 * Ceilingnodes - 2) = -2 * CeilingFourier2
        UC(2 * Ceilingnodes - 1, 2 * Ceilingnodes - 1) = 1 + 2 * CeilingFourier2
        If HaveFloorSubstrate(room) = False Then
            Lf(Floornodes, Floornodes - 1) = -2 * FloorFourier(room)
            Lf(Floornodes, Floornodes) = 1 + 2 * FloorFourier(room)
        Else
            Lf(2 * Floornodes - 1, 2 * Floornodes - 2) = -2 * FloorFourier2
            Lf(2 * Floornodes - 1, 2 * Floornodes - 1) = 1 + 2 * FloorFourier2
        End If

        'interior boundary conditions
        'InsideUpperConvCoeff = Get_Convection_Coefficient(uppertemp(i), uppertemp(1))
        wx(1, 1) = -2 * QUpperWall(room, i) * 1000 * WallFourier(room) * WallDeltaX(room) / WallConductivity(room) + UWallNode(room, 1, i)
        cx(1, 1) = -2 * QCeiling(room, i) * 1000 * CeilingFourier(room) * CeilingDeltaX(room) / CeilingConductivity(room) + CeilingNode(room, 1, i)
        'InsideLowerConvCoeff = Get_Convection_Coefficient(lowertemp(i), lowertemp(1))
        FX(1, 1) = -2 * QFloor(room, i) * 1000 * FloorFourier(room) * FloorDeltaX(room) / FloorConductivity(room) + FloorNode(room, 1, i)
        LX(1, 1) = -2 * QLowerWall(room, i) * 1000 * WallFourier(room) * WallDeltaX(room) / WallConductivity(room) + LWallNode(room, 1, i)

        For k = 2 To 2 * Wallnodes - 2
            wx(k, 1) = UWallNode(room, k, i)
            LX(k, 1) = LWallNode(room, k, i)
        Next k

        For k = 2 To 2 * Ceilingnodes - 2
            cx(k, 1) = CeilingNode(room, k, i)
        Next k

        If HaveFloorSubstrate(room) = False Then
            For k = 2 To Floornodes - 1
                FX(k, 1) = FloorNode(room, k, i)
            Next k
        Else
            For k = 2 To 2 * Floornodes - 2
                FX(k, 1) = FloorNode(room, k, i)
            Next k
        End If

        'exterior boundary conditions
        'wx(2 * Wallnodes - 1, 1) = 2 * WallFourier2 * WallOutsideBiot2 * ((ExteriorTemp - UWallNode(room, 2 * Wallnodes - 1, i)) - Surface_Emissivity(2, room) / OutsideConvCoeff * StefanBoltzmann * (UWallNode(room, 2 * Wallnodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + UWallNode(room, 2 * Wallnodes - 1, i)
        wx(2 * Wallnodes - 1, 1) = 2 * WallFourier2 * WallOutsideBiot2 * ((ExteriorTemp - UWallNode(room, 2 * Wallnodes - 1, i)) - emmissivity2 / OutsideConvCoeff * StefanBoltzmann * (UWallNode(room, 2 * Wallnodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + UWallNode(room, 2 * Wallnodes - 1, i)
        'cx(2 * Ceilingnodes - 1, 1) = 2 * CeilingFourier2 * CeilingOutsideBiot2 * ((ExteriorTemp - CeilingNode(room, 2 * Ceilingnodes - 1, i)) - Surface_Emissivity(1, room) / OutsideConvCoeff * StefanBoltzmann * (CeilingNode(room, 2 * Ceilingnodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + CeilingNode(room, 2 * Ceilingnodes - 1, i)
        cx(2 * Ceilingnodes - 1, 1) = 2 * CeilingFourier2 * CeilingOutsideBiot2 * ((ExteriorTemp - CeilingNode(room, 2 * Ceilingnodes - 1, i)) - emmissivity2 / OutsideConvCoeff * StefanBoltzmann * (CeilingNode(room, 2 * Ceilingnodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + CeilingNode(room, 2 * Ceilingnodes - 1, i)
        'LX(2 * Wallnodes - 1, 1) = 2 * WallFourier2 * WallOutsideBiot2 * ((ExteriorTemp - LWallNode(room, 2 * Wallnodes - 1, i)) - Surface_Emissivity(3, room) / OutsideConvCoeff * StefanBoltzmann * (LWallNode(room, 2 * Wallnodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + LWallNode(room, 2 * Wallnodes - 1, i)
        LX(2 * Wallnodes - 1, 1) = 2 * WallFourier2 * WallOutsideBiot2 * ((ExteriorTemp - LWallNode(room, 2 * Wallnodes - 1, i)) - emmissivity2 / OutsideConvCoeff * StefanBoltzmann * (LWallNode(room, 2 * Wallnodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + LWallNode(room, 2 * Wallnodes - 1, i)
        If HaveFloorSubstrate(room) = False Then
            FX(Floornodes, 1) = 2 * FloorFourier(room) * FloorOutsideBiot(room) * ((ExteriorTemp - FloorNode(room, Floornodes, i)) - Surface_Emissivity(4, room) / OutsideConvCoeff * StefanBoltzmann * (FloorNode(room, Floornodes, i) ^ 4 - ExteriorTemp ^ 4)) + FloorNode(room, Floornodes, i)
        Else
            'FX(2 * Floornodes - 1, 1) = 2 * FloorFourier2 * FloorOutsideBiot2 * ((ExteriorTemp - FloorNode(room, 2 * Floornodes - 1, i)) - Surface_Emissivity(4, room) / OutsideConvCoeff * StefanBoltzmann * (FloorNode(room, 2 * Floornodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + FloorNode(room, 2 * Floornodes - 1, i)
            FX(2 * Floornodes - 1, 1) = 2 * FloorFourier2 * FloorOutsideBiot2 * ((ExteriorTemp - FloorNode(room, 2 * Floornodes - 1, i)) - emmissivity2 / OutsideConvCoeff * StefanBoltzmann * (FloorNode(room, 2 * Floornodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + FloorNode(room, 2 * Floornodes - 1, i)
        End If

        If frmOptions1.optLUdecom.Checked = True Then
            'find surface temperatures for the next timestep
            'using method of LU decomposition
            Call MatSol(UW, wx, 2 * Wallnodes - 1) 'upper wall
            Call MatSol(UC, cx, 2 * Ceilingnodes - 1) 'ceiling
            Call MatSol(LW, LX, 2 * Wallnodes - 1) 'lower wall
            If HaveFloorSubstrate(room) = False Then
                Call MatSol(Lf, FX, Floornodes) 'floor
            Else
                Call MatSol(Lf, FX, 2 * Floornodes - 1)
            End If
        Else
            'find surface temperatures for the next timestep
            'using method of Gauss-Jordan elimination
            Call LINEAR2(2 * Wallnodes - 1, UW, wx, ier)
            Call LINEAR2(2 * Ceilingnodes - 1, UC, cx, ier)
            Call LINEAR2(2 * Wallnodes - 1, LW, LX, ier)
            If HaveFloorSubstrate(room) = False Then
                Call LINEAR2(Floornodes, Lf, FX, ier)
            Else
                Call LINEAR2(2 * Floornodes - 1, Lf, FX, ier)
            End If
            If ier = 1 Then MsgBox("singular matrix in implicit_surface_temps2")
        End If

        For j = 1 To 2 * Wallnodes - 1
            UWallNode(room, j, i + 1) = wx(j, 1)
            LWallNode(room, j, i + 1) = LX(j, 1)
        Next j
        For j = 1 To 2 * Ceilingnodes - 1
            CeilingNode(room, j, i + 1) = cx(j, 1)
        Next j
        If HaveFloorSubstrate(room) = False Then
            For j = 1 To Floornodes
                FloorNode(room, j, i + 1) = FX(j, 1)
            Next j
        Else
            For j = 1 To 2 * Floornodes - 1
                FloorNode(room, j, i + 1) = FX(j, 1)
            Next j
        End If
        'store surface temps in another array
        Upperwalltemp(room, i + 1) = UWallNode(room, 1, i + 1)
        CeilingTemp(room, i + 1) = CeilingNode(room, 1, i + 1)
        LowerWallTemp(room, i + 1) = LWallNode(room, 1, i + 1)
        FloorTemp(room, i + 1) = FloorNode(room, 1, i + 1)

        UnexposedUpperwalltemp(room, i + 1) = UWallNode(room, 2 * Wallnodes - 1, i + 1)
        UnexposedLowerwalltemp(room, i + 1) = LWallNode(room, 2 * Wallnodes - 1, i + 1)
        UnexposedCeilingtemp(room, i + 1) = CeilingNode(room, 2 * Ceilingnodes - 1, i + 1)
        If HaveFloorSubstrate(room) = True Then
            UnexposedFloortemp(room, i + 1) = FloorNode(room, 2 * Floornodes - 1, i + 1)
        Else
            UnexposedFloortemp(room, i + 1) = FloorNode(room, Floornodes, i + 1)
        End If
        Erase UW
        Erase UC
        Erase wx
        Erase cx
        Erase Lf
        Erase FX
        Erase LX
        Erase LW
        Exit Sub

temp_error_handler:
        Exit Sub
    End Sub
	
    Sub Implicit_Surface_Temps3(ByVal room As Integer, ByVal wall As Integer, ByVal i As Integer, ByRef UWallNode(,,) As Double, ByRef CeilingNode(,,) As Double, ByRef LWallNode(,,) As Double, ByRef FloorNode(,,) As Double)
        '*  ================================================================
        '*      This function updates the surface temperatures, using an
        '*      implicit finite difference method.
        '*
        '*      Two layer composite wall, single layer ceiling
        '*
        '*      Revised: Colleen Wade 3 June 1997
        '*  ================================================================

        Dim k, j As Integer
        Dim FloorFourier2, WallFourier2, WallOutsideBiot2, FloorOutsideBiot2 As Double
        Dim UW(2 * Wallnodes - 1, 2 * Wallnodes - 1) As Double
        Dim UC(Ceilingnodes, Ceilingnodes) As Double
        Dim wx(2 * Wallnodes - 1, 1) As Double
        Dim cx(Ceilingnodes, 1) As Double
        Dim Lf(Floornodes, Floornodes) As Double
        Dim FX(Floornodes, 1) As Double
        Dim LX(2 * Wallnodes - 1, 1) As Double
        Dim LW(2 * Wallnodes - 1, 2 * Wallnodes - 1) As Double
        Dim emissivity2 As Double = 0.8

        If HaveFloorSubstrate(room) = False Then
            ReDim Lf(Floornodes, Floornodes)
            ReDim FX(Floornodes, 1)
        Else
            ReDim Lf(2 * Floornodes - 1, 2 * Floornodes - 1)
            ReDim FX(2 * Floornodes - 1, 1)
            FloorFourier2 = FloorSubConductivity(room) * Timestep / (((FloorSubThickness(room) / 1000) / (Floornodes - 1)) ^ 2 * FloorSubDensity(room) * FloorSubSpecificHeat(room))
            FloorOutsideBiot2 = OutsideConvCoeff * ((FloorSubThickness(room) / 1000) / (Floornodes - 1)) / FloorSubConductivity(room)
        End If
        On Error GoTo temp_error_handler2

        WallFourier2 = WallSubConductivity(room) * Timestep / (((WallSubThickness(room) / 1000) / (Wallnodes - 1)) ^ 2 * WallSubDensity(room) * WallSubSpecificHeat(room))
        WallOutsideBiot2 = OutsideConvCoeff * ((WallSubThickness(room) / 1000) / (Wallnodes - 1)) / WallSubConductivity(room)

        UW(1, 1) = 1 + 2 * WallFourier(room)
        UW(1, 2) = -2 * WallFourier(room)
        LW(1, 1) = 1 + 2 * WallFourier(room)
        LW(1, 2) = -2 * WallFourier(room)
        UC(1, 1) = 1 + 2 * CeilingFourier(room)
        UC(1, 2) = -2 * CeilingFourier(room)
        Lf(1, 1) = 1 + 2 * FloorFourier(room)
        Lf(1, 2) = -2 * FloorFourier(room)

        If HaveFloorSubstrate(room) = False Then
            k = 2
            For j = 2 To Floornodes - 1
                Lf(j, k - 1) = -FloorFourier(room)
                Lf(j, k) = 1 + 2 * FloorFourier(room)
                Lf(j, k + 1) = -FloorFourier(room)
                k = k + 1
            Next j
        Else
            k = 2
            For j = 2 To Floornodes - 1
                Lf(j, k - 1) = -FloorFourier(room)
                Lf(j, k) = 1 + 2 * FloorFourier(room)
                Lf(j, k + 1) = -FloorFourier(room)
                k = k + 1
            Next j

            j = Floornodes
            Lf(j, k - 1) = -FloorFourier(room)
            Lf(j, k) = 1 + FloorFourier(room) + FloorFourier2
            Lf(j, k + 1) = -FloorFourier2
            k = k + 1

            For j = Floornodes + 1 To 2 * Floornodes - 2
                Lf(j, k - 1) = -FloorFourier2
                Lf(j, k) = 1 + 2 * FloorFourier2
                Lf(j, k + 1) = -FloorFourier2
                k = k + 1
            Next j
        End If

        k = 2
        For j = 2 To Wallnodes - 1
            UW(j, k - 1) = -WallFourier(room)
            UW(j, k) = 1 + 2 * WallFourier(room)
            UW(j, k + 1) = -WallFourier(room)
            LW(j, k - 1) = -WallFourier(room)
            LW(j, k) = 1 + 2 * WallFourier(room)
            LW(j, k + 1) = -WallFourier(room)
            k = k + 1
        Next j

        j = Wallnodes
        UW(j, k - 1) = -WallFourier(room)
        UW(j, k) = 1 + WallFourier(room) + WallFourier2
        UW(j, k + 1) = -WallFourier2
        LW(j, k - 1) = -WallFourier(room)
        LW(j, k) = 1 + WallFourier(room) + WallFourier2
        LW(j, k + 1) = -WallFourier2
        k = k + 1

        For j = Wallnodes + 1 To 2 * Wallnodes - 2
            UW(j, k - 1) = -WallFourier2
            UW(j, k) = 1 + 2 * WallFourier2
            UW(j, k + 1) = -WallFourier2
            LW(j, k - 1) = -WallFourier2
            LW(j, k) = 1 + 2 * WallFourier2
            LW(j, k + 1) = -WallFourier2
            k = k + 1
        Next j

        k = 2
        For j = 2 To Ceilingnodes - 1
            UC(j, k - 1) = -CeilingFourier(room)
            UC(j, k) = 1 + 2 * CeilingFourier(room)
            UC(j, k + 1) = -CeilingFourier(room)
            k = k + 1
        Next j

        UW(2 * Wallnodes - 1, 2 * Wallnodes - 2) = -2 * WallFourier2
        UW(2 * Wallnodes - 1, 2 * Wallnodes - 1) = 1 + 2 * WallFourier2
        LW(2 * Wallnodes - 1, 2 * Wallnodes - 2) = -2 * WallFourier2
        LW(2 * Wallnodes - 1, 2 * Wallnodes - 1) = 1 + 2 * WallFourier2
        UC(Ceilingnodes, Ceilingnodes - 1) = -2 * CeilingFourier(room)
        UC(Ceilingnodes, Ceilingnodes) = 1 + 2 * CeilingFourier(room)
        If HaveFloorSubstrate(room) = False Then
            Lf(Floornodes, Floornodes - 1) = -2 * FloorFourier(room)
            Lf(Floornodes, Floornodes) = 1 + 2 * FloorFourier(room)
        Else
            Lf(2 * Floornodes - 1, 2 * Floornodes - 2) = -2 * FloorFourier2
            Lf(2 * Floornodes - 1, 2 * Floornodes - 1) = 1 + 2 * FloorFourier2
        End If

        'interior boundary conditions
        wx(1, 1) = -2 * QUpperWall(room, i) * 1000 * WallFourier(room) * WallDeltaX(room) / WallConductivity(room) + UWallNode(room, 1, i)

        cx(1, 1) = -2 * QCeiling(room, i) * 1000 * CeilingFourier(room) * CeilingDeltaX(room) / CeilingConductivity(room) + CeilingNode(room, 1, i)
        FX(1, 1) = -2 * QFloor(room, i) * 1000 * FloorFourier(room) * FloorDeltaX(room) / FloorConductivity(room) + FloorNode(room, 1, i)
        LX(1, 1) = -2 * QLowerWall(room, i) * 1000 * WallFourier(room) * WallDeltaX(room) / WallConductivity(room) + LWallNode(room, 1, i)

        For k = 2 To 2 * Wallnodes - 2
            wx(k, 1) = UWallNode(room, k, i)
            LX(k, 1) = LWallNode(room, k, i)
        Next k

        For k = 2 To Ceilingnodes - 1
            cx(k, 1) = CeilingNode(room, k, i)
        Next k

        If HaveFloorSubstrate(room) = False Then
            For k = 2 To Floornodes - 1
                FX(k, 1) = FloorNode(room, k, i)
            Next k
        Else
            For k = 2 To 2 * Floornodes - 2
                FX(k, 1) = FloorNode(room, k, i)
            Next k
        End If

        'exterior boundary conditions
        'wx(2 * Wallnodes - 1, 1) = 2 * WallFourier2 * WallOutsideBiot2 * ((ExteriorTemp - UWallNode(room, 2 * Wallnodes - 1, i)) - Surface_Emissivity(2, room) / OutsideConvCoeff * StefanBoltzmann * (UWallNode(room, 2 * Wallnodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + UWallNode(room, 2 * Wallnodes - 1, i)
        wx(2 * Wallnodes - 1, 1) = 2 * WallFourier2 * WallOutsideBiot2 * ((ExteriorTemp - UWallNode(room, 2 * Wallnodes - 1, i)) - emissivity2 / OutsideConvCoeff * StefanBoltzmann * (UWallNode(room, 2 * Wallnodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + UWallNode(room, 2 * Wallnodes - 1, i)
        cx(Ceilingnodes, 1) = 2 * CeilingFourier(room) * CeilingOutsideBiot(room) * ((ExteriorTemp - CeilingNode(room, Ceilingnodes, i)) - Surface_Emissivity(1, room) / OutsideConvCoeff * StefanBoltzmann * (CeilingNode(room, Ceilingnodes, i) ^ 4 - ExteriorTemp ^ 4)) + CeilingNode(room, Ceilingnodes, i)
        'LX(2 * Wallnodes - 1, 1) = 2 * WallFourier2 * WallOutsideBiot2 * ((ExteriorTemp - LWallNode(room, 2 * Wallnodes - 1, i)) - Surface_Emissivity(3, room) / OutsideConvCoeff * StefanBoltzmann * (LWallNode(room, 2 * Wallnodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + LWallNode(room, 2 * Wallnodes - 1, i)
        LX(2 * Wallnodes - 1, 1) = 2 * WallFourier2 * WallOutsideBiot2 * ((ExteriorTemp - LWallNode(room, 2 * Wallnodes - 1, i)) - emissivity2 / OutsideConvCoeff * StefanBoltzmann * (LWallNode(room, 2 * Wallnodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + LWallNode(room, 2 * Wallnodes - 1, i)
        If HaveFloorSubstrate(room) = False Then
            FX(Floornodes, 1) = 2 * FloorFourier(room) * FloorOutsideBiot(room) * ((ExteriorTemp - FloorNode(room, Floornodes, i)) - Surface_Emissivity(4, room) / OutsideConvCoeff * StefanBoltzmann * (FloorNode(room, Floornodes, i) ^ 4 - ExteriorTemp ^ 4)) + FloorNode(room, Floornodes, i)
        Else
            'FX(2 * Floornodes - 1, 1) = 2 * FloorFourier2 * FloorOutsideBiot2 * ((ExteriorTemp - FloorNode(room, 2 * Floornodes - 1, i)) - Surface_Emissivity(4, room) / OutsideConvCoeff * StefanBoltzmann * (FloorNode(room, 2 * Floornodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + FloorNode(room, 2 * Floornodes - 1, i)
            FX(2 * Floornodes - 1, 1) = 2 * FloorFourier2 * FloorOutsideBiot2 * ((ExteriorTemp - FloorNode(room, 2 * Floornodes - 1, i)) - emissivity2 / OutsideConvCoeff * StefanBoltzmann * (FloorNode(room, 2 * Floornodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + FloorNode(room, 2 * Floornodes - 1, i)
        End If

        Dim ier As Short
        If frmOptions1.optLUdecom.Checked = True Then
            'find surface temperatures for the next timestep
            'using method of LU decomposition
            Call MatSol(UW, wx, 2 * Wallnodes - 1) 'upper wall
            Call MatSol(UC, cx, Ceilingnodes) 'ceiling
            Call MatSol(LW, LX, 2 * Wallnodes - 1) 'lower wall
            If HaveFloorSubstrate(room) = False Then
                Call MatSol(Lf, FX, Floornodes) 'floor
            Else
                Call MatSol(Lf, FX, 2 * Floornodes - 1)
            End If
        Else
            'find surface temperatures for the next timestep
            'using method of Gauss-Jordan elimination
            Call LINEAR2(2 * Wallnodes - 1, UW, wx, ier)
            Call LINEAR2(Ceilingnodes, UC, cx, ier)
            Call LINEAR2(2 * Wallnodes - 1, LW, LX, ier)
            If HaveFloorSubstrate(room) = False Then
                Call LINEAR2(Floornodes, Lf, FX, ier)
            Else
                Call LINEAR2(2 * Floornodes - 1, Lf, FX, ier)
            End If
            If ier = 1 Then MsgBox("singular matrix in implicit_surface_temps3")
        End If

        For j = 1 To 2 * Wallnodes - 1
            UWallNode(room, j, i + 1) = wx(j, 1)
            LWallNode(room, j, i + 1) = LX(j, 1)
        Next j
        For j = 1 To Ceilingnodes
            CeilingNode(room, j, i + 1) = cx(j, 1)
        Next j
        If HaveFloorSubstrate(room) = False Then
            For j = 1 To Floornodes
                FloorNode(room, j, i + 1) = FX(j, 1)
            Next j
        Else
            For j = 1 To 2 * Floornodes - 1
                FloorNode(room, j, i + 1) = FX(j, 1)
            Next j
        End If

        'store surface temps in another array
        Upperwalltemp(room, i + 1) = UWallNode(room, 1, i + 1)
        CeilingTemp(room, i + 1) = CeilingNode(room, 1, i + 1)
        LowerWallTemp(room, i + 1) = LWallNode(room, 1, i + 1)
        FloorTemp(room, i + 1) = FloorNode(room, 1, i + 1)

        UnexposedUpperwalltemp(room, i + 1) = UWallNode(room, 2 * Wallnodes - 1, i + 1)
        UnexposedLowerwalltemp(room, i + 1) = LWallNode(room, 2 * Wallnodes - 1, i + 1)
        UnexposedCeilingtemp(room, i + 1) = CeilingNode(room, Ceilingnodes, i + 1)
        If HaveFloorSubstrate(room) = True Then
            UnexposedFloortemp(room, i + 1) = FloorNode(room, 2 * Floornodes - 1, i + 1)
        Else
            UnexposedFloortemp(room, i + 1) = FloorNode(room, Floornodes, i + 1)
        End If
        Erase UW
        Erase UC
        Erase wx
        Erase cx
        Erase Lf
        Erase FX
        Erase LX
        Erase LW
        Exit Sub

temp_error_handler2:
        Exit Sub

    End Sub
	
    Sub Implicit_Surface_Temps4(ByVal room As Short, ByVal ceiling As Short, ByVal i As Short, ByRef UWallNode(,,) As Double, ByRef CeilingNode(,,) As Double, ByRef LWallNode(,,) As Double, ByRef FloorNode(,,) As Double)
        '*  ================================================================
        '*      This function updates the surface temperatures, using an
        '*      implicit finite difference method.
        '*
        '*      single layer wall, ceiling with substrate
        '*
        '*      Revised: Colleen Wade 3 June 1997
        '*  ================================================================

        Dim k, j As Integer
        Dim FloorFourier2, CeilingFourier2, CeilingOutsideBiot2, FloorOutsideBiot2 As Double
        Dim UW(Wallnodes, Wallnodes) As Double
        Dim UC(2 * Ceilingnodes - 1, 2 * Ceilingnodes - 1) As Double
        Dim wx(Wallnodes, 1) As Double
        Dim cx(2 * Ceilingnodes - 1, 1) As Double
        Dim Lf(Floornodes, Floornodes) As Double
        Dim FX(Floornodes, 1) As Double
        Dim LX(Wallnodes, 1) As Double
        Dim LW(Wallnodes, Wallnodes) As Double
        Dim emissivity2 As Double = 0.8

        If HaveFloorSubstrate(room) = False Then
            ReDim Lf(Floornodes, Floornodes)
            ReDim FX(Floornodes, 1)
        Else
            ReDim Lf(2 * Floornodes - 1, 2 * Floornodes - 1)
            ReDim FX(2 * Floornodes - 1, 1)
            FloorFourier2 = FloorSubConductivity(room) * Timestep / (((FloorSubThickness(room) / 1000) / (Floornodes - 1)) ^ 2 * FloorSubDensity(room) * FloorSubSpecificHeat(room))
            FloorOutsideBiot2 = OutsideConvCoeff * ((FloorSubThickness(room) / 1000) / (Floornodes - 1)) / FloorSubConductivity(room)
        End If

        On Error GoTo temp_error_handler3

        CeilingFourier2 = CeilingSubConductivity(room) * Timestep / (((CeilingSubThickness(room) / 1000) / (Ceilingnodes - 1)) ^ 2 * CeilingSubDensity(room) * CeilingSubSpecificHeat(room))
        CeilingOutsideBiot2 = OutsideConvCoeff * ((CeilingSubThickness(room) / 1000) / (Ceilingnodes - 1)) / CeilingSubConductivity(room)

        UW(1, 1) = 1 + 2 * WallFourier(room)
        UW(1, 2) = -2 * WallFourier(room)
        LW(1, 1) = 1 + 2 * WallFourier(room)
        LW(1, 2) = -2 * WallFourier(room)
        UC(1, 1) = 1 + 2 * CeilingFourier(room)
        UC(1, 2) = -2 * CeilingFourier(room)
        Lf(1, 1) = 1 + 2 * FloorFourier(room)
        Lf(1, 2) = -2 * FloorFourier(room)

        If HaveFloorSubstrate(room) = False Then
            k = 2
            For j = 2 To Floornodes - 1
                Lf(j, k - 1) = -FloorFourier(room)
                Lf(j, k) = 1 + 2 * FloorFourier(room)
                Lf(j, k + 1) = -FloorFourier(room)
                k = k + 1
            Next j
        Else
            k = 2
            For j = 2 To Floornodes - 1
                Lf(j, k - 1) = -FloorFourier(room)
                Lf(j, k) = 1 + 2 * FloorFourier(room)
                Lf(j, k + 1) = -FloorFourier(room)
                k = k + 1
            Next j

            j = Floornodes
            Lf(j, k - 1) = -FloorFourier(room)
            Lf(j, k) = 1 + FloorFourier(room) + FloorFourier2
            Lf(j, k + 1) = -FloorFourier2
            k = k + 1

            For j = Floornodes + 1 To 2 * Floornodes - 2
                Lf(j, k - 1) = -FloorFourier2
                Lf(j, k) = 1 + 2 * FloorFourier2
                Lf(j, k + 1) = -FloorFourier2
                k = k + 1
            Next j
        End If

        k = 2
        For j = 2 To Wallnodes - 1
            LW(j, k - 1) = -WallFourier(room)
            LW(j, k) = 1 + 2 * WallFourier(room)
            LW(j, k + 1) = -WallFourier(room)
            UW(j, k - 1) = -WallFourier(room)
            UW(j, k) = 1 + 2 * WallFourier(room)
            UW(j, k + 1) = -WallFourier(room)
            k = k + 1
        Next j

        k = 2
        For j = 2 To Ceilingnodes - 1
            UC(j, k - 1) = -CeilingFourier(room)
            UC(j, k) = 1 + 2 * CeilingFourier(room)
            UC(j, k + 1) = -CeilingFourier(room)
            k = k + 1
        Next j

        j = Ceilingnodes
        UC(j, k - 1) = -CeilingFourier(room)
        UC(j, k) = 1 + CeilingFourier(room) + CeilingFourier2
        UC(j, k + 1) = -CeilingFourier2
        k = k + 1

        For j = Ceilingnodes + 1 To 2 * Ceilingnodes - 2
            UC(j, k - 1) = -CeilingFourier2
            UC(j, k) = 1 + 2 * CeilingFourier2
            UC(j, k + 1) = -CeilingFourier2
            k = k + 1
        Next j

        UW(Wallnodes, Wallnodes - 1) = -2 * WallFourier(room)
        UW(Wallnodes, Wallnodes) = 1 + 2 * WallFourier(room)
        LW(Wallnodes, Wallnodes - 1) = -2 * WallFourier(room)
        LW(Wallnodes, Wallnodes) = 1 + 2 * WallFourier(room)
        UC(2 * Ceilingnodes - 1, 2 * Ceilingnodes - 2) = -2 * CeilingFourier2
        UC(2 * Ceilingnodes - 1, 2 * Ceilingnodes - 1) = 1 + 2 * CeilingFourier2
        If HaveFloorSubstrate(room) = False Then
            Lf(Floornodes, Floornodes - 1) = -2 * FloorFourier(room)
            Lf(Floornodes, Floornodes) = 1 + 2 * FloorFourier(room)
        Else
            Lf(2 * Floornodes - 1, 2 * Floornodes - 2) = -2 * FloorFourier2
            Lf(2 * Floornodes - 1, 2 * Floornodes - 1) = 1 + 2 * FloorFourier2
        End If

        'interior boundary conditions
        wx(1, 1) = -2 * QUpperWall(room, i) * 1000 * WallFourier(room) * WallDeltaX(room) / WallConductivity(room) + UWallNode(room, 1, i)
        cx(1, 1) = -2 * QCeiling(room, i) * 1000 * CeilingFourier(room) * CeilingDeltaX(room) / CeilingConductivity(room) + CeilingNode(room, 1, i)
        FX(1, 1) = -2 * QFloor(room, i) * 1000 * FloorFourier(room) * FloorDeltaX(1) / FloorConductivity(room) + FloorNode(room, 1, i)
        LX(1, 1) = -2 * QLowerWall(room, i) * 1000 * WallFourier(room) * WallDeltaX(room) / WallConductivity(room) + LWallNode(room, 1, i)

        For k = 2 To Wallnodes - 1
            wx(k, 1) = UWallNode(room, k, i)
            LX(k, 1) = LWallNode(room, k, i)
        Next k

        For k = 2 To 2 * Ceilingnodes - 2
            cx(k, 1) = CeilingNode(room, k, i)
        Next k

        If HaveFloorSubstrate(room) = False Then
            For k = 2 To Floornodes - 1
                FX(k, 1) = FloorNode(room, k, i)
            Next k
        Else
            For k = 2 To 2 * Floornodes - 2
                FX(k, 1) = FloorNode(room, k, i)
            Next k
        End If

        'exterior boundary conditions
        wx(Wallnodes, 1) = 2 * WallFourier(room) * WallOutsideBiot(room) * ((ExteriorTemp - UWallNode(room, Wallnodes, i)) - Surface_Emissivity(2, room) / OutsideConvCoeff * StefanBoltzmann * (UWallNode(room, Wallnodes, i) ^ 4 - ExteriorTemp ^ 4)) + UWallNode(room, Wallnodes, i)
        'cx(2 * Ceilingnodes - 1, 1) = 2 * CeilingFourier2 * CeilingOutsideBiot2 * ((ExteriorTemp - CeilingNode(room, 2 * Ceilingnodes - 1, i)) - Surface_Emissivity(1, room) / OutsideConvCoeff * StefanBoltzmann * (CeilingNode(room, 2 * Ceilingnodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + CeilingNode(room, 2 * Ceilingnodes - 1, i)
        cx(2 * Ceilingnodes - 1, 1) = 2 * CeilingFourier2 * CeilingOutsideBiot2 * ((ExteriorTemp - CeilingNode(room, 2 * Ceilingnodes - 1, i)) - emissivity2 / OutsideConvCoeff * StefanBoltzmann * (CeilingNode(room, 2 * Ceilingnodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + CeilingNode(room, 2 * Ceilingnodes - 1, i)
        LX(Wallnodes, 1) = 2 * WallFourier(room) * WallOutsideBiot(room) * ((ExteriorTemp - LWallNode(room, Wallnodes, i)) - Surface_Emissivity(3, room) / OutsideConvCoeff * StefanBoltzmann * (LWallNode(room, Wallnodes, i) ^ 4 - ExteriorTemp ^ 4)) + LWallNode(room, Wallnodes, i)
        If HaveFloorSubstrate(room) = False Then
            FX(Floornodes, 1) = 2 * FloorFourier(room) * FloorOutsideBiot(room) * ((ExteriorTemp - FloorNode(room, Floornodes, i)) - Surface_Emissivity(4, room) / OutsideConvCoeff * StefanBoltzmann * (FloorNode(room, Floornodes, i) ^ 4 - ExteriorTemp ^ 4)) + FloorNode(room, Floornodes, i)
        Else
            'FX(2 * Floornodes - 1, 1) = 2 * FloorFourier2 * FloorOutsideBiot2 * ((ExteriorTemp - FloorNode(room, 2 * Floornodes - 1, i)) - Surface_Emissivity(4, room) / OutsideConvCoeff * StefanBoltzmann * (FloorNode(room, 2 * Floornodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + FloorNode(room, 2 * Floornodes - 1, i)
            FX(2 * Floornodes - 1, 1) = 2 * FloorFourier2 * FloorOutsideBiot2 * ((ExteriorTemp - FloorNode(room, 2 * Floornodes - 1, i)) - emissivity2 / OutsideConvCoeff * StefanBoltzmann * (FloorNode(room, 2 * Floornodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + FloorNode(room, 2 * Floornodes - 1, i)
        End If

        Dim ier As Short
        If frmOptions1.optLUdecom.Checked = True Then
            'find surface temperatures for the next timestep
            'using method of LU decomposition
            Call MatSol(UW, wx, Wallnodes) 'upper wall
            Call MatSol(UC, cx, 2 * Ceilingnodes - 1) 'ceiling
            Call MatSol(LW, LX, Wallnodes) 'lower wall
            If HaveFloorSubstrate(room) = False Then
                Call MatSol(Lf, FX, Floornodes) 'floor
            Else
                Call MatSol(Lf, FX, 2 * Floornodes - 1)
            End If
        Else
            'find surface temperatures for the next timestep
            'using method of Gauss-Jordan elimination
            Call LINEAR2(Wallnodes, UW, wx, ier)
            Call LINEAR2(2 * Ceilingnodes - 1, UC, cx, ier)
            Call LINEAR2(Wallnodes, LW, LX, ier)
            If HaveFloorSubstrate(room) = False Then
                Call LINEAR2(Floornodes, Lf, FX, ier)
            Else
                Call LINEAR2(2 * Floornodes - 1, Lf, FX, ier)
            End If
            If ier = 1 Then MsgBox("singular matrix in implicit_surface_temps4")
        End If

        For j = 1 To Wallnodes
            UWallNode(room, j, i + 1) = wx(j, 1)
            LWallNode(room, j, i + 1) = LX(j, 1)
        Next j
        For j = 1 To 2 * Ceilingnodes - 1
            CeilingNode(room, j, i + 1) = cx(j, 1)
        Next j
        If HaveFloorSubstrate(room) = False Then
            For j = 1 To Floornodes
                FloorNode(room, j, i + 1) = FX(j, 1)
            Next j
        Else
            For j = 1 To 2 * Floornodes - 1
                FloorNode(room, j, i + 1) = FX(j, 1)
            Next j
        End If

        'store surface temps in another array
        Upperwalltemp(room, i + 1) = UWallNode(room, 1, i + 1)
        CeilingTemp(room, i + 1) = CeilingNode(room, 1, i + 1)
        LowerWallTemp(room, i + 1) = LWallNode(room, 1, i + 1)
        FloorTemp(room, i + 1) = FloorNode(room, 1, i + 1)

        UnexposedUpperwalltemp(room, i + 1) = UWallNode(room, Wallnodes, i + 1)
        UnexposedLowerwalltemp(room, i + 1) = LWallNode(room, Wallnodes, i + 1)
        UnexposedCeilingtemp(room, i + 1) = CeilingNode(room, 2 * Ceilingnodes - 1, i + 1)
        If HaveFloorSubstrate(room) = True Then
            UnexposedFloortemp(room, i + 1) = FloorNode(room, 2 * Floornodes - 1, i + 1)
        Else
            UnexposedFloortemp(room, i + 1) = FloorNode(room, Floornodes, i + 1)
        End If
        Erase UW
        Erase UC
        Erase wx
        Erase cx
        Erase Lf
        Erase FX
        Erase LX
        Erase LW

        Exit Sub

temp_error_handler3:
        Exit Sub

    End Sub
	
	Function Initial_H2O() As Double
		'************************************************************
		'*
		'*  This function returns the initial mass fraction of water
		'*  vapor in the room.
		'*
		'*  Revised 1 December 1995 Colleen Wade
		'************************************************************
		
		Dim n As Integer
		Dim PWS, WS As Double
		Dim Temps() As Double
		Dim Psat() As Double
		n = 11
		
		ReDim Temps(n)
		ReDim Psat(n)
		
		'determine the saturation pressure for the water
		'vapor corresponding to the initial ambient temp
		'in the room
		
		'read in steam table values
		Call steam_table(Temps, Psat, n)
		
		'interpolate steam table to determine the sat pressure (PWS)
		'corresponding to the interior ambient temperature
		Call Interpolate_S(Temps, Psat, n, InteriorTemp, PWS)
		
		'determine the saturation humidity ratio
		WS = 0.622 * PWS / (Atm_Pressure - PWS)
		
		Initial_H2O = 0.622 * RelativeHumidity * WS / (WS + 0.622 - RelativeHumidity * WS)
		
	End Function
	
    Function Lower_Layer_Absorb(ByVal room As Integer, ByVal VentAreaUpper As Double, ByVal VentAreaLower As Double, ByVal Z As Double, ByRef prd(,) As Double, ByVal A1 As Double, ByVal A2 As Double, ByVal A3 As Double, ByVal A4 As Double, ByVal uppertemp As Double, ByVal lowertemp As Double, ByVal CeilingTemp As Double, ByVal Upperwalltemp As Double, ByVal LowerWallTemp As Double, ByVal FloorTemp As Double) As Double
        '************************************************************
        '*  This function returns the radiant heat absorbed by the
        '*  lower layer.
        '*
        '*  Revised Colleen Wade 18 August 1995
        '************************************************************

        Dim j, k As Integer
        Dim G1, qout, gemm, h As Double
        Dim ST As Double
        Dim A() As Double
        ReDim A(4)

        On Error GoTo absorbhandler

        G1 = TransmissionFactor(2, 2, room) 'upper layer transmission
        h = TransmissionFactor(3, 3, room) 'lower layer transmission

        A(1) = A1
        A(2) = A2
        A(3) = A3
        A(4) = A4

        qout = 0
        gemm = 0

        'emitting surface
        For j = 1 To 4
            If j = 1 Then ST = CeilingTemp 'assume no vents in ceiling
            If j = 2 And A(2) <> 0 Then ST = (Upperwalltemp ^ 4 - VentAreaUpper / A(2) * (Upperwalltemp ^ 4 - ExteriorTemp ^ 4)) ^ 0.25
            If j = 3 And A(3) <> 0 Then ST = (LowerWallTemp ^ 4 - VentAreaLower / A(3) * (LowerWallTemp ^ 4 - ExteriorTemp ^ 4)) ^ 0.25
            If j = 4 Then ST = FloorTemp 'assume no vents in floor

            'receiving surface
            For k = 1 To 4
                If j < 3 And k > 2 Then

                    'upper to  lower
                    'absorbtion due to heat emitting wall surface
                    qout = qout + G1 * (1 - h) * A(j) * F(j, k, room) * (StefanBoltzmann / 1000 * (ST ^ 4) - (1 - Surface_Emissivity(j, room)) / Surface_Emissivity(j, room) * prd(j, 1))
                    'due to gas layer emission
                    gemm = gemm + (1 - h) * (1 - G1) * StefanBoltzmann / 1000 * (uppertemp ^ 4) * A(j) * F(j, k, room) - A(j) * F(j, k, room) * (1 - h) * StefanBoltzmann / 1000 * (lowertemp ^ 4)
                ElseIf j > 2 Then
                    'lower to upper or lower
                    'absorbtion due to heat emitting wall surface
                    qout = qout + (1 - h) * A(j) * F(j, k, room) * (StefanBoltzmann / 1000 * (ST ^ 4) - (1 - Surface_Emissivity(j, room)) / Surface_Emissivity(j, room) * prd(j, 1))
                    'due to gas layer emission
                    gemm = gemm - A(j) * F(j, k, room) * (1 - h) * (StefanBoltzmann / 1000 * lowertemp ^ 4)
                End If
            Next k
        Next j

        'lets add in the radiation loss though the openings
        If room = fireroom Then
            Lower_Layer_Absorb = radfireLower + qout + gemm - (1 - h) * VentAreaLower * StefanBoltzmann / 1000 * (lowertemp ^ 4 - ExteriorTemp ^ 4) 'kW
        Else
            Lower_Layer_Absorb = radfireLower + qout + gemm
        End If
        Exit Function

absorbhandler:
        MsgBox(ErrorToString(Err.Number) & " Lower_Layer_Absorb")
        ERRNO = Err.Number
        Exit Function

    End Function
	
    Sub Matlub(ByRef A(,) As Double, ByVal n As Integer, ByRef index() As Double, ByRef b(,) As Double)
        '************************************************************
        '*  Procedure performs forward substitution and also back
        '*  substitution on a LU decomposed matrix. The net result
        '*  is the solution of the equation A X = B.
        '*
        '*  From ProMath 2.0
        '*  Revised by Colleen Wade 4/7/95
        '************************************************************

        Dim Ii, i, j As Integer
        Dim Ll As Double
        Dim Accum As Double

        On Error GoTo erroverflow

        Ii = 0

        For i = 1 To n
            Ll = index(i)
            Accum = b(Ll, 1)
            b(Ll, 1) = b(i, 1)

            If Ii <> 0 Then
                For j = Ii To i - 1
                    Accum = Accum - A(i, j) * b(j, 1)
                Next j
            ElseIf Accum <> 0 Then
                Ii = i
            End If

            b(i, 1) = Accum
        Next i

        For i = n To 1 Step -1

            Accum = b(i, 1)
            If i < n Then
                For j = i + 1 To n
                    Accum = Accum - A(i, j) * b(j, 1)
                Next j
            End If

            b(i, 1) = Accum / A(i, i)
        Next i
        Exit Sub

MatlubErrorHandler:
        MsgBox(ErrorToString(Err.Number) & " in Matlub")
        ERRNO = Err.Number
        'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'System.Windows.Forms.Cursor.Current = Default_Renamed
        Exit Sub

erroverflow:
        b(i, 1) = 0
        A(i, j) = 0
        Resume Next

    End Sub
	
	
    Sub Matlud(ByRef A(,) As Double, ByVal n As Integer, ByRef index() As Double, ByRef D As Double)
        '************************************************************
        '*
        '*  Procedure replaces a square matrix A with the LU decomposition
        '*  of a rowwise permutation of itself.
        '*
        '*  From ProMath 2.0
        '*  Revised by Colleen Wade 4/7/95
        '************************************************************

        Dim tiny, Aamax As Double
        Dim i, j As Integer
        Dim Accum, Dum As Double
        Dim k, imax As Integer
        Dim VV() As Double
        ReDim VV(n)

        tiny = 1.0E-20
        D = 1

        For i = 1 To n
            Aamax = 0
            For j = 1 To n
                If (Abs(A(i, j)) > Aamax) Then Aamax = Abs(A(i, j))
            Next j
            If (Aamax = 0) Then

                'Dim message As String = "Singular Matrix in Matlud (HEATRAN.VB)"
                'If ProjectDirectory = RiskDataDirectory Then
                '    frmInputs.rtb_log.Text = tim(stepcount, 1) & " sec: " & message.ToString & Chr(13) & frmInputs.rtb_log.Text
                'Else
                '    MsgBox("Singular Matrix in Matlud")
                'End If

                'flagstop = 1
                Exit Sub
            End If

            VV(i) = 1 / Aamax
        Next i

        For j = 1 To n
            If (j > 1) Then
                For i = 1 To j - 1
                    Accum = A(i, j)
                    If (i > 1) Then
                        For k = 1 To i - 1
                            Accum = Accum - A(i, k) * A(k, j)
                        Next k
                        A(i, j) = Accum
                    End If
                Next i
            End If

            Aamax = 0
            For i = j To n
                Accum = A(i, j)
                If (j > 1) Then
                    For k = 1 To j - 1
                        Accum = Accum - A(i, k) * A(k, j)
                    Next k
                    A(i, j) = Accum
                End If

                Dum = VV(i) * Abs(Accum)
                If Dum >= Aamax Then imax = i : Aamax = Dum
            Next i

            If j <> imax Then
                For k = 1 To n
                    Dum = A(imax, k)
                    A(imax, k) = A(j, k)
                    A(j, k) = Dum
                Next k
                D = -D
                VV(imax) = VV(j)
            End If

            index(j) = imax
            If (j <> n) Then
                If (A(j, j) = 0.0#) Then A(j, j) = tiny
                Dum = 1.0# / A(j, j)
                For i = j + 1 To n : A(i, j) = A(i, j) * Dum : Next i
            End If
        Next j

        If A(n, n) = 0 Then A(n, n) = tiny
    End Sub
	
    Sub Matprd(ByRef A(,) As Double, ByRef b(,) As Double, ByRef R(,) As Double, ByVal n As Integer, ByVal m As Integer, ByVal L As Integer)
        '************************************************************
        '*  Procedure multiplies two matrices B() and A().
        '*  Solution held in matrix R() = A() x B().
        '*
        '*  A() first N x M matrix
        '*  B() second M x L matrix
        '*  R() output N x L matrix
        '*  N number of rows in A() and R()
        '*  M number of columns in A() and rows in B()
        '*  L number of columns in B() and R()
        '*
        '*  From ProMath 2.0
        '*  Revised by Colleen Wade 4 July 1995
        '************************************************************

        Dim j, i, Im As Integer

        On Error GoTo MatprdErrorHandler

        For i = 1 To n
            For j = 1 To L
                R(i, j) = 0
                For Im = 1 To m
                    R(i, j) = R(i, j) + A(i, Im) * b(Im, j)
                Next Im
            Next j
        Next i
        Exit Sub

MatprdErrorHandler:
        MsgBox(ErrorToString(Err.Number) & " in Matprd")
        ERRNO = Err.Number

        Exit Sub

    End Sub
	
    Sub MatrixAB(ByVal room As Integer, ByRef MatA(,) As Double, ByRef MatB(,) As Double, ByVal Z As Double, ByVal A1 As Double, ByVal A2 As Double, ByVal A3 As Double, ByVal A4 As Double, ByVal uppertemp As Double, ByVal lowertemp As Double, ByVal YCO2 As Double, ByVal YH2O As Double, ByVal YCO2Lower As Double, ByVal YH2Olower As Double, ByVal volume As Double)
        '************************************************************
        '*  Procedure defines the elements in the Matrices A and B
        '*
        '*  Z = position of the layer interface above floor
        '*  A1, A2, A3, A4 = areas of ceiling, upper wall, lower wall, floor
        '*
        '*  Called by: FourWallRad
        '*  Calls : Calc_Configs, Transmission
        '*
        '*  Revised: 16 November 1995 Colleen Wade
        '************************************************************

        Dim k, j, delta As Integer

        'get configuration factors to define F(k,j)
        Call Calc_Configs(room, Z, A1, A2, A3, A4)

        'get transmission factors
        Call Transmission(room, volume, Z, uppertemp, lowertemp, YCO2, YH2O, YCO2Lower, YH2Olower)

        For k = 1 To 4
            For j = 1 To 4

                'define kronecker delta function
                If j = k Then
                    delta = 1
                Else
                    delta = 0
                End If

                'define elements of matrices A and B
                MatA(k, j) = delta - F(k, j, room) * TransmissionFactor(j, k, room) * (1 - Surface_Emissivity(j, room))
                MatB(k, j) = delta - F(k, j, room) * TransmissionFactor(j, k, room)

            Next j
        Next k

    End Sub

    Sub MatrixC(ByVal room As Integer, ByVal Q As Double, ByVal Z As Double, ByVal A1 As Double, ByVal A2 As Double, ByVal A3 As Double, ByVal A4 As Double, ByRef matc(,) As Double, ByVal UT As Double, ByVal LT As Double)
        '************************************************************
        '*  Procedure defines the elements in the Matrix C. This includes
        '*  the radiant heat flux striking each wall due to a point source
        '*  fire and that due to an emitting gas layer.
        '*
        '*  Q = total heat release (kW)
        '*  Z = position of the layer interface above floor
        '*  A1 = ceiling area
        '*  A2 = upper wall area
        '*  A3 = lower wall area
        '*  A4 = floor area
        '*  MatC() = the returned matrix 4 rows x 1 col
        '*  UT = upper layer temp
        '*  LT = lower layer temp
        '*
        '*  Called by: FourWallRad
        '*  Calls: wAngle
        '*
        '*  Revised: 24 November 1995 Colleen Wade
        '************************************************************

        Dim k As Integer
        Dim b, A, Area As Double
        Dim qgaslower, qfire, qgasupper As Double
        Dim j As Integer
        Dim PointSource, firehi As Double

        A = TransmissionFactor(2, 2, room) 'upper layer transmission fire to surface k
        b = TransmissionFactor(3, 3, room) 'lower layer transmission fire to surface k

        radfireUpper = 0
        radfireLower = 0

        On Error Resume Next
        'assume point source located at approx 1/2 flame height
        Dim QF As Double = HeatRelease(fireroom, stepcount, 2)
        Dim HRRUA As Double = ObjectMLUA(2, 1)
        Dim maxheight As Double = FireHeight(1) + 0.5 * (RoomHeight(fireroom) - FireHeight(1))

        firehi = FireHeight(1) + 0.5 * (0.235 * QF ^ (2 / 5) - 1.02 * Sqrt(4 * QF / PI / HRRUA))
        'If stepcount = 30000 Then Stop
        If firehi > maxheight Then
            firehi = maxheight
        End If
        If firehi < FireHeight(1) Then
            firehi = FireHeight(1)
        End If

        '2014.13 and below
        ' firehi = FireHeight(1) + 0.23 * 0.5 * NewRadiantLossFraction(1) * HeatRelease(fireroom, stepcount, 2) ^ (2 / 5)
        'If firehi > RoomHeight(fireroom) Then
        'firehi = RoomHeight(fireroom)
        'End If

        'radiant heat flux striking the k'th rectangular wall segment due to the fire
        For k = 1 To 4
            If room = fireroom Then
                'PointSource = NewRadiantLossFraction(1) * Q * wAngle(k, Z) / (4 * PI)
                PointSource = NewRadiantLossFraction(1) * Q * wAngle(k, Z, firehi) 'kW directed toward the surface concerned
            Else
                PointSource = 0
            End If
            If k = 1 Then
                'ceiling
                Area = A1
                If firehi < Z Then
                    qfire = A * b * PointSource / Area 'assume fire in lower
                Else
                    qfire = A * PointSource / Area 'assume fire in upper layer
                End If
            ElseIf k = 2 Then
                'upper wall
                Area = A2
                If Area <> 0 Then
                    If firehi < Z Then
                        qfire = A * b * PointSource / Area
                    Else
                        qfire = A * PointSource / Area 'assume fire in upper layer
                    End If
                Else : qfire = 0
                End If

            ElseIf k = 3 Then
                'lower wall
                Area = A3
                If Area <> 0 Then
                    If firehi < Z Then
                        qfire = b * PointSource / Area
                    Else
                        qfire = A * b * PointSource / Area
                    End If
                Else : qfire = 0
                End If
            ElseIf k = 4 Then
                'floor
                Area = A4
                If firehi < Z Then
                    qfire = b * PointSource / Area
                Else
                    qfire = A * b * PointSource / Area
                End If
            End If

            'radiant heat absorbed by the upper layer due to point source fire
            'If k < 3 Then
            '    radfireUpper = radfireUpper + b * (1 - A) * PointSource 'kW
            'End If

            ''radiant heat absorbed by the lower layer due to point source fire
            'radfireLower = radfireLower + (1 - b) * PointSource 'kW

            If k < 3 Then
                If firehi < Z Then
                    radfireUpper = radfireUpper + b * (1 - A) * PointSource 'kW
                    radfireLower = radfireLower + (1 - b) * PointSource 'kW
                Else
                    radfireUpper = radfireUpper + (1 - A) * PointSource 'kW
                    radfireLower = radfireLower  'kW
                End If
            End If
            If k > 2 Then
                If firehi < Z Then
                    radfireUpper = radfireUpper  'kW
                    radfireLower = radfireLower + (1 - b) * PointSource 'kW
                Else
                    radfireUpper = radfireUpper + (1 - A) * PointSource 'kW
                    radfireLower = radfireLower + A * (1 - b) * PointSource 'kW
                End If
            End If


            'radiant heat flux striking the k'th rectangular wall segment due to emitting gas layer
            qgaslower = 0 'kW/m2
            qgasupper = 0 'kW/m2

            For j = 1 To 4
                If j <= 2 And k <= 2 Then
                    'upper to upper
                    qgaslower = qgaslower
                    qgasupper = F(k, j, room) * StefanBoltzmann / 1000 * (1 - A) * (UT ^ 4) + qgasupper
                ElseIf j <= 2 And k > 2 Then
                    'upper to lower
                    'from Forney's paper
                    'qgaslower = F(k, j, room) * StefanBoltzmann / 1000 * (1 - b) * LT ^ 4 + qgaslower
                    'qgasupper = F(k, j, room) * StefanBoltzmann / 1000 * b * (1 - A) * UT ^ 4 + qgasupper
                    'what I think they should be
                    qgaslower = F(k, j, room) * A * (1 - b) * StefanBoltzmann / 1000 * LT ^ 4 + qgaslower
                    qgasupper = F(k, j, room) * (1 - A) * StefanBoltzmann / 1000 * UT ^ 4 + qgasupper
                ElseIf j > 2 And k <= 2 Then
                    'lower to upper
                    'from Forney's paper
                    'qgaslower = F(k, j, room) * StefanBoltzmann / 1000 * (1 - b) * LT ^ 4 * b + qgaslower
                    'qgasupper = F(k, j, room) * StefanBoltzmann / 1000 * (1 - A) * UT ^ 4 + qgasupper
                    'what I think they should be
                    qgaslower = F(k, j, room) * (1 - b) * StefanBoltzmann / 1000 * LT ^ 4 + qgaslower
                    qgasupper = F(k, j, room) * b * (1 - A) * StefanBoltzmann / 1000 * UT ^ 4 + qgasupper
                ElseIf j > 2 And k > 2 Then
                    'lower to lower
                    qgaslower = F(k, j, room) * StefanBoltzmann / 1000 * (1 - b) * LT ^ 4 + qgaslower
                    qgasupper = qgasupper
                End If
            Next j
            matc(k, 1) = (qfire + qgasupper + qgaslower) 'kW/m2

        Next k
        On Error GoTo 0
        Application.DoEvents()

    End Sub
    Sub MatrixZ(ByVal room As Integer, ByVal Q As Double, ByVal Z As Double, ByVal A1 As Double, ByVal A2 As Double, ByVal A3 As Double, ByVal A4 As Double, ByRef matc(,) As Double, ByVal UT As Double, ByVal LT As Double)
        '************************************************************
        '*  Procedure defines the elements in the Matrix C. This includes
        '*  the radiant heat flux striking each wall due to a point source
        '*  fire and that due to an emitting gas layer.
        '*
        '*  Q = total heat release (kW)
        '*  Z = position of the layer interface above floor
        '*  A1 = ceiling area
        '*  A2 = upper wall area
        '*  A3 = lower wall area
        '*  A4 = floor area
        '*  MatZ() = the returned matrix 4 rows x 1 col
        '*  UT = upper layer temp
        '*  LT = lower layer temp
        '*
        '*  Called by: FourWallRad
        '*  Calls: wAngle
        '*
        '*  Revised: 24 November 1995 Colleen Wade
        '************************************************************

        Dim k As Integer
        Dim b, A, Area As Double
        Dim qgaslower, qfire, qgasupper As Double
        Dim j As Integer
        Dim PointSource, firehi As Double

        A = TransmissionFactor(2, 2, room) 'upper layer transmission fire to surface k
        b = TransmissionFactor(3, 3, room) 'lower layer transmission fire to surface k

        radfireUpper = 0
        radfireLower = 0

        On Error Resume Next
        'assume point source located at approx 1/2 flame height
        'Dim QF As Double = HeatRelease(fireroom, stepcount, 2)
        ' Dim HRRUA As Double = ObjectMLUA(2, 1)
        'Dim maxheight As Double = FireHeight(1) + 0.5 * (RoomHeight(fireroom) - FireHeight(1))

        'firehi = FireHeight(1) + 0.5 * (0.235 * QF ^ (2 / 5) - 1.02 * Sqrt(4 * QF / PI / HRRUA))
        'If stepcount = 30000 Then Stop
        'If firehi > maxheight Then
        '    firehi = maxheight
        'End If
        'If firehi < FireHeight(1) Then
        '    firehi = FireHeight(1)
        'End If

        '2014.13 and below
        ' firehi = FireHeight(1) + 0.23 * 0.5 * NewRadiantLossFraction(1) * HeatRelease(fireroom, stepcount, 2) ^ (2 / 5)
        'If firehi > RoomHeight(fireroom) Then
        'firehi = RoomHeight(fireroom)
        'End If

        'radiant heat flux striking the k'th rectangular wall segment due to the fire
        For k = 1 To 4
            '    If room = fireroom Then
            '        'PointSource = NewRadiantLossFraction(1) * Q * wAngle(k, Z) / (4 * PI)
            '        PointSource = NewRadiantLossFraction(1) * Q * wAngle(k, Z, firehi) 'kW directed toward the surface concerned
            '    Else
            '        PointSource = 0
            '    End If
            '    If k = 1 Then
            '        'ceiling
            '        Area = A1
            '        If firehi < Z Then
            '            qfire = A * b * PointSource / Area 'assume fire in lower
            '        Else
            '            qfire = A * PointSource / Area 'assume fire in upper layer
            '        End If
            '    ElseIf k = 2 Then
            '        'upper wall
            '        Area = A2
            '        If Area <> 0 Then
            '            If firehi < Z Then
            '                qfire = A * b * PointSource / Area
            '            Else
            '                qfire = A * PointSource / Area 'assume fire in upper layer
            '            End If
            '        Else : qfire = 0
            '        End If

            '    ElseIf k = 3 Then
            '        'lower wall
            '        Area = A3
            '        If Area <> 0 Then
            '            If firehi < Z Then
            '                qfire = b * PointSource / Area
            '            Else
            '                qfire = A * b * PointSource / Area
            '            End If
            '        Else : qfire = 0
            '        End If
            '    ElseIf k = 4 Then
            '        'floor
            '        Area = A4
            '        If firehi < Z Then
            '            qfire = b * PointSource / Area
            '        Else
            '            qfire = A * b * PointSource / Area
            '        End If
            '    End If

            'radiant heat absorbed by the upper layer due to point source fire
            'If k < 3 Then
            '    radfireUpper = radfireUpper + b * (1 - A) * PointSource 'kW
            'End If

            ''radiant heat absorbed by the lower layer due to point source fire
            'radfireLower = radfireLower + (1 - b) * PointSource 'kW

            'If k < 3 Then
            '    If firehi < Z Then
            '        radfireUpper = radfireUpper + b * (1 - A) * PointSource 'kW
            '        radfireLower = radfireLower + (1 - b) * PointSource 'kW
            '    Else
            '        radfireUpper = radfireUpper + (1 - A) * PointSource 'kW
            '        radfireLower = radfireLower  'kW
            '    End If
            'End If
            'If k > 2 Then
            '    If firehi < Z Then
            '        radfireUpper = radfireUpper  'kW
            '        radfireLower = radfireLower + (1 - b) * PointSource 'kW
            '    Else
            '        radfireUpper = radfireUpper + (1 - A) * PointSource 'kW
            '        radfireLower = radfireLower + A * (1 - b) * PointSource 'kW
            '    End If
            'End If


            'radiant heat flux striking the k'th rectangular wall segment due to emitting gas layer
            qgaslower = 0 'kW/m2
            qgasupper = 0 'kW/m2

            For j = 1 To 4
                If j <= 2 And k <= 2 Then
                    'upper to upper
                    qgaslower = qgaslower
                    qgasupper = F(k, j, room) * StefanBoltzmann / 1000 * (1 - A) * (UT ^ 4) + qgasupper
                ElseIf j <= 2 And k > 2 Then
                    'upper to lower
                    'from Forney's paper
                    'qgaslower = F(k, j, room) * StefanBoltzmann / 1000 * (1 - b) * LT ^ 4 + qgaslower
                    'qgasupper = F(k, j, room) * StefanBoltzmann / 1000 * b * (1 - A) * UT ^ 4 + qgasupper
                    'what I think they should be
                    qgaslower = F(k, j, room) * A * (1 - b) * StefanBoltzmann / 1000 * LT ^ 4 + qgaslower
                    qgasupper = F(k, j, room) * (1 - A) * StefanBoltzmann / 1000 * UT ^ 4 + qgasupper
                ElseIf j > 2 And k <= 2 Then
                    'lower to upper
                    'from Forney's paper
                    'qgaslower = F(k, j, room) * StefanBoltzmann / 1000 * (1 - b) * LT ^ 4 * b + qgaslower
                    'qgasupper = F(k, j, room) * StefanBoltzmann / 1000 * (1 - A) * UT ^ 4 + qgasupper
                    'what I think they should be
                    qgaslower = F(k, j, room) * (1 - b) * StefanBoltzmann / 1000 * LT ^ 4 + qgaslower
                    qgasupper = F(k, j, room) * b * (1 - A) * StefanBoltzmann / 1000 * UT ^ 4 + qgasupper
                ElseIf j > 2 And k > 2 Then
                    'lower to lower
                    qgaslower = F(k, j, room) * StefanBoltzmann / 1000 * (1 - b) * LT ^ 4 + qgaslower
                    qgasupper = qgasupper
                End If
            Next j
            matc(k, 1) = (qfire + qgasupper + qgaslower) 'kW/m2

        Next k
        On Error GoTo 0
        Application.DoEvents()

    End Sub

    Sub MatrixD(ByRef MatD(,) As Double, ByVal room As Integer)
        '************************************************************
        '*  Procedure defines the elements in the Matrix D, a scaling
        '*  matrix containing the emittance of each wall segment.
        '*
        '*  Called by: FourWallRad
        '*
        '*  Revised by Colleen Wade 4/7/95
        '************************************************************

        Dim j, k As Integer

        For j = 1 To 4
            For k = 1 To 4
                If k = j Then
                    MatD(j, k) = Surface_Emissivity(j, room)
                Else
                    MatD(j, k) = 0
                End If
            Next k
        Next j

    End Sub
	
    Sub MatrixE(ByVal CeilingTemp As Double, ByVal Upperwalltemp As Double, ByVal LowerWallTemp As Double, ByVal FloorTemp As Double, ByRef MatE(,) As Double)
        '************************************************************
        '*  Procedure defines the elements in the Matrix E
        '*
        '*  Called by: FourWallRad
        '*
        '*  Revised by Colleen Wade 4/7/95
        '************************************************************

        MatE(1, 1) = StefanBoltzmann / 1000 * CeilingTemp ^ 4 'kW/m2
        MatE(2, 1) = StefanBoltzmann / 1000 * Upperwalltemp ^ 4
        MatE(3, 1) = StefanBoltzmann / 1000 * LowerWallTemp ^ 4
        MatE(4, 1) = StefanBoltzmann / 1000 * FloorTemp ^ 4

    End Sub
    Public Sub LINEAR2(ByVal n As Integer, ByRef A(,) As Double, ByRef b(,) As Double, ByRef ier As Short)

        ' Solution of a system of linear equations subroutine

        ' Solves [ A ] * { X } = { B } using the Gauss-Jordan method

        ' Input

        '  N   = number of equations
        '  A() = matrix of coefficients   (N rows by N columns)
        '  B() = right hand column vector (N rows)

        ' Output

        '  A() = inverse of incoming matrix [ A ] (N rows by N columns)
        '  B() = solution vector of linear system (N rows)
        '  IER = error flag
        '    0 = no error
        '    1 = singular matrix

        Dim IPIVOT() As Integer
        Dim INDEXR() As Integer
        Dim INDEXC() As Integer
        Dim IR As Integer
        Dim IC As Integer
        Dim tmp, PMAX, PIVINV As Double
        Dim k, i, j, L As Integer
        Dim Ll As Integer
        ReDim IPIVOT(n)
        ReDim INDEXR(n)
        ReDim INDEXC(n)

        ier = 0

        For j = 1 To n
            IPIVOT(j) = 0
        Next j

        For i = 1 To n
            PMAX = 0.0#

            For j = 1 To n
                If (IPIVOT(j) <> 1) Then
                    For k = 1 To n
                        If (IPIVOT(k) = 0) Then
                            If (Abs(A(j, k)) >= PMAX) Then
                                PMAX = Abs(A(j, k))
                                IR = j
                                IC = k
                            End If
                        ElseIf (IPIVOT(k) > 1) Then
                            ier = 1
                            GoTo EXITSUB
                        End If
                    Next k
                End If
            Next j

            IPIVOT(IC) = IPIVOT(IC) + 1

            If (IR <> IC) Then
                For L = 1 To n
                    tmp = A(IR, L)
                    A(IR, L) = A(IC, L)
                    A(IC, L) = tmp
                Next L

                tmp = b(IR, 1)
                b(IR, 1) = b(IC, 1)
                b(IC, 1) = tmp
            End If

            INDEXR(i) = IR
            INDEXC(i) = IC

            If (A(IC, IC) = 0.0#) Then
                ier = 1
                GoTo EXITSUB
            End If

            PIVINV = 1.0# / A(IC, IC)
            A(IC, IC) = 1.0#

            For L = 1 To n
                A(IC, L) = A(IC, L) * PIVINV
            Next L

            b(IC, 1) = b(IC, 1) * PIVINV

            For Ll = 1 To n
                If (Ll <> IC) Then
                    tmp = A(Ll, IC)
                    A(Ll, IC) = 0.0#
                    For L = 1 To n
                        A(Ll, L) = A(Ll, L) - A(IC, L) * tmp
                    Next L
                    b(Ll, 1) = b(Ll, 1) - b(IC, 1) * tmp
                End If
            Next Ll
        Next i

        'calculates the matrix inverse which we don't need
        'For L = N To 1 Step -1
        '    If (INDEXR(L) <> INDEXC(L)) Then
        '       For k = 1 To N
        '           TMP = A(k, INDEXR(L))
        '           A(k, INDEXR(L)) = A(k, INDEXC(L))
        '           A(k, INDEXC(L)) = TMP
        '       Next k
        '    End If
        'Next L

EXITSUB:
        Erase IPIVOT
        Erase INDEXC
        Erase INDEXR

    End Sub
	
    Sub MatSol(ByRef A(,) As Double, ByRef b(,) As Double, ByVal n As Integer)
        '************************************************************
        '*  Procedure computes the solution of a linear system of
        '*  equations of the form A X = B, where A is an N x N matrix
        '*  and X and B are column vectors.
        '*
        '*  Solution also stored in B().
        '*
        '*  Revised by Colleen Wade 4 July 1995
        '************************************************************
        Dim indexx() As Double
        Dim D As Double
        ReDim indexx(n)

        'Dim ainv() As Double, ier As Integer, R() As Double
        'ReDim ainv(1 To N, 1 To N), R(1 To N, 1 To N)
        'Call INVERSE(N, A(), ainv(), ier)
        'Call Matprd(A(), ainv(), R(), ByVal N, ByVal N, ByVal N)
        'Dim i As Integer
        'For i = 1 To N
        '   Debug.Print R(N, N)
        'Next i
        'If N > 4 Then Stop

        'call procedure to LU decompose the matrix
        Call Matlud(A, n, indexx, D)

        'call procedure to perform forward and back substitution
        Call Matlub(A, n, indexx, b)

    End Sub
	
    Sub Matsub(ByRef A(,) As Double, ByRef b(,) As Double, ByRef R(,) As Double, ByVal n As Integer, ByVal m As Integer)
        '************************************************************
        '*  Procedure subtracts a general matrix B() from another A().
        '*  Solution held in matrix R() = A() - B().
        '*
        '*  A() first N x M matrix
        '*  B() second N x M matrix
        '*  R() output N x M matrix
        '*  N number of rows in A() B() RES()
        '*  M number of columns in A() B() RES()
        '*
        '*  Revised by Colleen Wade 4 July 1995
        '************************************************************

        Dim i, j As Integer

        For i = 1 To n
            For j = 1 To m
                R(i, j) = A(i, j) - b(i, j)
            Next j
        Next i

    End Sub
	
	Sub steam_table(ByRef Temps() As Double, ByRef Psat() As Double, ByVal n As Integer)
		'************************************************************
		'*  This function returns the value of the saturation pressure (kPa)
		'*  for water vapor corresponding to a particular air temperature.
		'*
		'*  Revised 1 December 1995 Colleen Wade
		'************************************************************
		
		n = 11
		
		Temps(1) = 0 + 273 : Psat(1) = 0.6122
		Temps(2) = 5 + 273 : Psat(2) = 0.8719
		Temps(3) = 10 + 273 : Psat(3) = 1.227
		Temps(4) = 15 + 273 : Psat(4) = 1.704
		Temps(5) = 20 + 273 : Psat(5) = 2.337
		Temps(6) = 25 + 273 : Psat(6) = 3.166
		Temps(7) = 30 + 273 : Psat(7) = 4.242
		Temps(8) = 35 + 273 : Psat(8) = 5.622
		Temps(9) = 40 + 273 : Psat(9) = 7.375
		Temps(10) = 45 + 273 : Psat(10) = 9.582
		Temps(11) = 50 + 273 : Psat(11) = 12.33
		
	End Sub
	
	'UPGRADE_NOTE: Property was upgraded to Property_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Function Thermal_Prop(ByRef material As Short, ByRef Property_Renamed As String) As Double
		'*  ======================================================================
		'*      This function returns the thermal properties for a selected
		'*      materials
		'*  ======================================================================
		Select Case material
			Case 0
				'gypsum plaster
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop = 0.16 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop = 900 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop = 790 'kg/m3
						Exit Function
				End Select
			Case 1
				'concrete
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop = 1.4 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop = 880 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop = 2300 'kg/m3
						Exit Function
				End Select
			Case 2
				'brick
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop = 0.72 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop = 835 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop = 1920 'kg/m3
						Exit Function
				End Select
			Case 3
				'plywood
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop = 0.12 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop = 1215 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop = 545 'kg/m3
						Exit Function
				End Select
			Case 4
				'acoustic tile
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop = 0.058 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop = 1340 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop = 290 'kg/m3
						Exit Function
				End Select
			Case 5
				'particle board low density
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop = 0.078 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop = 1300 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop = 590 'kg/m3
						Exit Function
				End Select
			Case 6
				'hardwood
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop = 0.16 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop = 1255 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop = 720 'kg/m3
						Exit Function
				End Select
			Case 7
				'softwood
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop = 0.12 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop = 1380 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop = 510 'kg/m3
						Exit Function
				End Select
			Case 8
				'ceramic fibre
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop = 0.09
						Exit Function
					Case "Specific Heat"
						Thermal_Prop = 1040
						Exit Function
					Case "Density"
						Thermal_Prop = 128
						Exit Function
				End Select
			Case Else
		End Select
		
	End Function
	
	'UPGRADE_NOTE: Property was upgraded to Property_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Function Thermal_Prop_Floor(ByRef material As Short, ByRef Property_Renamed As String) As Double
		'*  ======================================================================
		'*      This function returns the thermal properties for a selected
		'*      material of the floor
		'*  ======================================================================
		
		Select Case material
			Case 0
				'concrete
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop_Floor = 1.4 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop_Floor = 880 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop_Floor = 2300 'kg/m3
						Exit Function
				End Select
			Case 1
				'brick
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop_Floor = 0.72 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop_Floor = 835 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop_Floor = 1920 'kg/m3
						Exit Function
				End Select
			Case 2
				'plywood
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop_Floor = 0.12 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop_Floor = 1215 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop_Floor = 545 'kg/m3
						Exit Function
				End Select
			Case 3
				'particle board low density
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop_Floor = 0.078 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop_Floor = 1300 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop_Floor = 590 'kg/m3
						Exit Function
				End Select
			Case 4
				'hardwood
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop_Floor = 0.16 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop_Floor = 1255 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop_Floor = 720 'kg/m3
						Exit Function
				End Select
			Case 5
				'softwood
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop_Floor = 0.12 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop_Floor = 1380 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop_Floor = 510 'kg/m3
						Exit Function
				End Select
			Case 6
				'ceramic fibre
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop_Floor = 0.09
						Exit Function
					Case "Specific Heat"
						Thermal_Prop_Floor = 1040
						Exit Function
					Case "Density"
						Thermal_Prop_Floor = 128
						Exit Function
				End Select
			Case Else
		End Select
		
	End Function
	
	'UPGRADE_NOTE: Property was upgraded to Property_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Function Thermal_Prop_Substrates(ByRef material As Short, ByRef Property_Renamed As String) As Double
		'*  ======================================================================
		'*      This function returns the thermal properties for a selected
		'*      materials
		'*  30/5/97 C A Wade
		'*  ======================================================================
		Select Case material
			Case 0
				'gypsum plaster
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop_Substrates = 0.16 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop_Substrates = 900 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop_Substrates = 790 'kg/m3
						Exit Function
				End Select
			Case 1
				'concrete
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop_Substrates = 1.4 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop_Substrates = 880 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop_Substrates = 2300 'kg/m3
						Exit Function
				End Select
			Case 2
				'brick
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop_Substrates = 0.72 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop_Substrates = 835 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop_Substrates = 1920 'kg/m3
						Exit Function
				End Select
			Case 3
				'plywood
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop_Substrates = 0.12 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop_Substrates = 1215 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop_Substrates = 545 'kg/m3
						Exit Function
				End Select
			Case 4
				'particle board low density
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop_Substrates = 0.078 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop_Substrates = 1300 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop_Substrates = 590 'kg/m3
						Exit Function
				End Select
			Case 5
				'hardwood
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop_Substrates = 0.16 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop_Substrates = 1255 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop_Substrates = 720 'kg/m3
						Exit Function
				End Select
			Case 6
				'softwood
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop_Substrates = 0.12 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop_Substrates = 1380 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop_Substrates = 510 'kg/m3
						Exit Function
				End Select
			Case 7
				'lightweight concrete krc=0.15 Nordic ISO 9705 room material
				Select Case Property_Renamed
					Case "Conductivity"
						Thermal_Prop_Substrates = 0.21 'W/mK
						Exit Function
					Case "Specific Heat"
						Thermal_Prop_Substrates = 880 'J/kgK
						Exit Function
					Case "Density"
						Thermal_Prop_Substrates = 800 'kg/m3
						Exit Function
				End Select
				
			Case Else
		End Select
	End Function
	
	Sub Transmission(ByVal room As Integer, ByVal volume As Double, ByVal Z As Double, ByVal uppertemp As Double, ByVal lowertemp As Double, ByVal YCO2 As Double, ByVal YH2O As Double, ByVal YCO2Lower As Double, ByVal YH2Olower As Double)
		'************************************************************
		'*  Procedure to calculate the transmission factors for the
		'*  room. Called at each time step.
		'*
		'*  Z = layer interface height from floor
		'*
		'*  Subprocedures and functions called: gas_emissivity_upper
		'*  Called by: MatrixAB
		'*
		'*  Revised: 23 November 1995 Colleen Wade
		'************************************************************
		
		Dim A, b As Double
		
		If stepcount > 1 Then
            A = 1 - gas_emissivity_upper(room, volume, Z, uppertemp, YCO2, YH2O, OD_upper(room, stepcount - 1)) 'transmission upper layer (air+soot)
            b = 1 - gas_emissivity_lower(room, Z, lowertemp, YCO2Lower, YH2Olower, volume) 'transmission lower layer (air)
		End If
		
		TransmissionFactor(1, 1, room) = 1
		TransmissionFactor(1, 2, room) = A
		TransmissionFactor(1, 3, room) = A * b
		TransmissionFactor(1, 4, room) = A * b
		TransmissionFactor(2, 1, room) = A
		TransmissionFactor(2, 2, room) = A
		TransmissionFactor(2, 3, room) = A * b
		TransmissionFactor(2, 4, room) = A * b
		TransmissionFactor(3, 1, room) = A * b
		TransmissionFactor(3, 2, room) = A * b
		TransmissionFactor(3, 3, room) = b
		TransmissionFactor(3, 4, room) = b
		TransmissionFactor(4, 1, room) = A * b
		TransmissionFactor(4, 2, room) = A * b
		TransmissionFactor(4, 3, room) = b
		TransmissionFactor(4, 4, room) = 1
		
	End Sub
	
    Function Upper_Layer_Absorb(ByVal room As Integer, ByVal VentAreaUpper As Double, ByVal VentAreaLower As Double, ByVal Z As Double, ByRef prd(,) As Double, ByVal A1 As Double, ByVal A2 As Double, ByVal A3 As Double, ByVal A4 As Double, ByVal uppertemp As Double, ByVal lowertemp As Double, ByVal CeilingTemp As Double, ByVal Upperwalltemp As Double, ByVal LowerWallTemp As Double, ByVal FloorTemp As Double) As Double
        '************************************************************
        '*  This function returns the radiant heat absorbed by the
        '*  upper layer.
        '*
        '*  Revised Colleen Wade 8 January 2002
        '************************************************************

        Dim j, k As Integer
        Dim qout, gemm As Double
        Dim G1, ST, h As Double
        Dim A() As Double
        ReDim A(4)


        On Error GoTo errorhandler

        G1 = TransmissionFactor(2, 2, room) 'transmission upper
        h = TransmissionFactor(3, 3, room) 'transmission lower

        A(1) = A1
        A(2) = A2 'area of upper wall
        A(3) = A3 'area of lower wall
        A(4) = A4

        qout = 0
        gemm = 0

        'emitting surface
        For j = 1 To 4
            'ST = the applicable surface temperature
            If j = 1 Then ST = CeilingTemp 'assume no vents in ceiling
            If j = 2 And A(2) <> 0 Then ST = (Upperwalltemp ^ 4 - VentAreaUpper / A(2) * (Upperwalltemp ^ 4 - ExteriorTemp ^ 4)) ^ 0.25
            If j = 3 And A(3) <> 0 Then ST = (LowerWallTemp ^ 4 - VentAreaLower / A(3) * (LowerWallTemp ^ 4 - ExteriorTemp ^ 4)) ^ 0.25
            If j = 4 Then ST = FloorTemp 'assume no vents in floor

            'receiving surface
            For k = 1 To 4
                If j < 3 Then
                    'upper to upper or lower
                    'absorbtion due to heat emitting wall surface
                    qout = qout + (1 - G1) * A(j) * F(j, k, room) * (StefanBoltzmann / 1000 * (ST ^ 4) - (1 - Surface_Emissivity(j, room)) / Surface_Emissivity(j, room) * prd(j, 1))
                    'due to gas layer emission
                    gemm = gemm - A(j) * F(j, k, room) * (1 - G1) * StefanBoltzmann / 1000 * (uppertemp ^ 4)
                ElseIf j > 2 And k < 3 Then
                    'lower to upper
                    'absorbtion due to heat emitting wall surface
                    qout = qout + h * (1 - G1) * A(j) * F(j, k, room) * (StefanBoltzmann / 1000 * (ST ^ 4) - (1 - Surface_Emissivity(j, room)) / Surface_Emissivity(j, room) * prd(j, 1))
                    'due to gas layer emission
                    gemm = gemm + (1 - h) * (1 - G1) * StefanBoltzmann / 1000 * (lowertemp ^ 4) * A(j) * F(j, k, room) - A(j) * F(j, k, room) * (1 - G1) * StefanBoltzmann / 1000 * (uppertemp ^ 4)
                End If
            Next k
        Next j

        'Upper_Layer_Absorb = radfireUpper + qout + gemm

        'lets add in the radiation loss though the openings for the fire room only
        If room = fireroom Then '2/10/2002
            Upper_Layer_Absorb = radfireUpper + qout + gemm - (1 - G1) * VentAreaUpper * StefanBoltzmann / 1000 * (uppertemp ^ 4 - ExteriorTemp ^ 4) 'kW
            'Upper_Layer_Absorb = radfireUpper + qout + gemm  'kW
        Else
            Upper_Layer_Absorb = radfireUpper + qout + gemm 'kW
        End If
        Exit Function

errorhandler:
        MsgBox(ErrorToString(Err.Number) & " Upper_Layer_Absorb")
        ERRNO = Err.Number
        'System.Windows.Forms.Cursor.Current = Default_Renamed
        Exit Function

    End Function
	
	Function ViewFactor_Parallel(ByVal room As Integer, ByVal Z As Double) As Double
		'*  ====================================================================
		'*  This function returns the value of the radiation view factor for the
		'*  compartment floor receiving radiation from a hot gas layer which is
		'*  descending in the room. The analytical expression is taken from the
		'*  text "Fundamentals of Heat and Mass Transfer" by Incropera and deWitt
		'*  for radiation exchange between two parallel plates of finite size.
		'*
		'*  Z = height of the upper layer interface above the fire
		'*
		'*  ====================================================================
		
		Dim Xbar, Ybar As Double
		Dim C, A, b, D As Double
		
		'some useful ratios
		If Z > 0 Then
			Xbar = RoomLength(room) / Z
			Ybar = RoomWidth(room) / Z
			
			A = (1 + Xbar ^ 2) * (1 + Ybar ^ 2) / (1 + Xbar ^ 2 + Ybar ^ 2)
            b = Xbar * Sqrt(1 + Ybar ^ 2) * Atan(Xbar / Sqrt(1 + Ybar ^ 2))
            C = Ybar * Sqrt(1 + Xbar ^ 2) * Atan(Ybar / Sqrt(1 + Xbar ^ 2))
            D = -Xbar * Atan(Xbar) - Ybar * Atan(Ybar)
			
            ViewFactor_Parallel = 2 / (PI * Xbar * Ybar) * (Log(Sqrt(A)) + b + C + D)
		Else
			ViewFactor_Parallel = 1
		End If
		
	End Function
	
    Function wAngle(ByRef surface As Integer, ByRef Z As Double, ByRef firehi As Double) As Double
        '************************************************************
        '*  Procedure calculates the solid angles between a point source
        '*  fire and the room surfaces assuming that the fire is located
        '*  in the centre of the room.
        '*
        '*  surface = surface identification
        '*  Z = position of the layer interface above floor
        '*
        '*  Called by: MatrixC
        '*  Calls: Calc_Angle
        '*
        '*  Revised: 16 November 1995 Colleen Wade
        '************************************************************

        Dim wAngle1, wAngle2 As Double
        'Dim firehi As Double
        'assume point source located at approx 1/2 flame height
        'Dim QF As Double = HeatRelease(fireroom, stepcount, 2)
        'Dim HRRUA As Double = ObjectMLUA(2, 1)
        'Dim maxheight As Double = FireHeight(1) + 0.5 * (RoomHeight(fireroom) - FireHeight(1))

        'firehi = FireHeight(1) + 0.5 * (0.235 * QF ^ (2 / 5) - 1.02 * Sqrt(4 * QF / PI / HRRUA))

        'If firehi > maxheight Then
        '    firehi = maxheight
        'End If
        'If firehi < FireHeight(1) Then
        '    firehi = FireHeight(1)
        'End If

        ' note height of object no 1 only is used.
        'assume point source located at approximately 1/2 flame height
        'firehi = FireHeight(1) + 0.23 * NewRadiantLossFraction(1) * HeatRelease(fireroom, stepcount, 2) ^ (2 / 5) / 2
        'If firehi > RoomHeight(fireroom) Then firehi = RoomHeight(fireroom)

        If surface = 1 Then 'ceiling
            wAngle = 4 * Calc_Angle(RoomLength(fireroom) / 2, RoomWidth(fireroom) / 2, RoomHeight(fireroom) - firehi)
        ElseIf surface = 2 Then  'upper wall
            If firehi < Z Then
                wAngle1 = Calc_Angle(RoomLength(fireroom) / 2, RoomHeight(fireroom) - firehi, RoomWidth(fireroom) / 2) - Calc_Angle(RoomLength(fireroom) / 2, Z - firehi, RoomWidth(fireroom) / 2)
                wAngle2 = Calc_Angle(RoomWidth(fireroom) / 2, RoomHeight(fireroom) - firehi, RoomLength(fireroom) / 2) - Calc_Angle(RoomWidth(fireroom) / 2, Z - firehi, RoomLength(fireroom) / 2)
            Else
                wAngle1 = Calc_Angle(RoomLength(fireroom) / 2, RoomHeight(fireroom) - firehi, RoomWidth(fireroom) / 2) + Calc_Angle(RoomLength(fireroom) / 2, firehi - Z, RoomWidth(fireroom) / 2)
                wAngle2 = Calc_Angle(RoomWidth(fireroom) / 2, RoomHeight(fireroom) - firehi, RoomLength(fireroom) / 2) + Calc_Angle(RoomWidth(fireroom) / 2, firehi - Z, RoomLength(fireroom) / 2)
            End If
            wAngle = 4 * (wAngle1 + wAngle2)
        ElseIf surface = 3 Then  'lower wall
            If firehi < Z Then
                wAngle1 = Calc_Angle(RoomLength(fireroom) / 2, Z - firehi, RoomWidth(fireroom) / 2) + Calc_Angle(RoomLength(fireroom) / 2, firehi, RoomWidth(fireroom) / 2)
                wAngle2 = Calc_Angle(RoomWidth(fireroom) / 2, Z - firehi, RoomLength(fireroom) / 2) + Calc_Angle(RoomWidth(fireroom) / 2, firehi, RoomLength(fireroom) / 2)
            Else
                wAngle1 = Calc_Angle(RoomLength(fireroom) / 2, firehi, RoomWidth(fireroom) / 2) - Calc_Angle(RoomLength(fireroom) / 2, CSng(Z), RoomWidth(fireroom) / 2)
                wAngle2 = Calc_Angle(RoomWidth(fireroom) / 2, firehi, RoomLength(fireroom) / 2) - Calc_Angle(RoomWidth(fireroom) / 2, CSng(Z), RoomLength(fireroom) / 2)
            End If
            wAngle = 4 * (wAngle1 + wAngle2)
        ElseIf surface = 4 Then  'floor
            wAngle = 4 * Calc_Angle(RoomLength(fireroom) / 2, RoomWidth(fireroom) / 2, firehi)
        End If

    End Function
	
	
    Public Function ViewFactor_Flame(ByVal A As Double, ByVal b As Double, ByVal Z As Double) As Double
        '*  ====================================================================
        '*  This function returns the value of the radiation view factor for the
        '*  compartment floor receiving radiation from a hot gas layer which is
        '*  descending in the room. The analytical expression is taken from the
        '*  SFPE handbook
        '*  for radiation exchange between a surface and a parallel differential element
        '*  diff element located opposite the corner of a surface with dimensions a and b
        '*  Z = height between the top of the flame volume and the ceiling (=differential element and the surface)
        '*
        '*  ====================================================================

        Dim Xbar, Ybar As Double
        Dim E, C, D, F As Double

        'some useful ratios
        If Z > 0 Then
            Xbar = A / Z
            Ybar = b / Z

            E = Xbar / Sqrt(1 + Xbar ^ 2)
            F = Ybar / Sqrt(1 + Xbar ^ 2)
            C = Ybar / Sqrt(1 + Ybar ^ 2)
            D = Xbar / Sqrt(1 + Ybar ^ 2)

            ViewFactor_Flame = (E * Atan(F) + C * Atan(D)) / (2 * PI)
        Else
            ViewFactor_Flame = 1
        End If

    End Function
	
    Public Sub Radiation_to_Surface(ByVal room As Integer, ByVal layer As Double, ByVal utemp As Double, ByVal YCO2 As Double, ByVal YH2O As Double, ByRef rad As Double, ByRef volume As Double, ByRef OD As Double, ByRef dist As Double, ByRef emissivity As Double)
        '=======================================================================
        '   calculate the received radiation by a differential elemental surface
        '   located at the monitoring height, due to radiation received from the
        '   upper layer
        '
        '=======================================================================
        Dim phi, distance, layerz As Double

        If dist >= layer Then
            phi = 1
        Else
            distance = layer - dist
            'get the configuration factor
            phi = 4 * ViewFactor_Flame(RoomWidth(room) / 2, RoomLength(room) / 2, distance)
        End If
        layerz = layer
        emissivity = gas_emissivity_upper(room, volume, layerz, utemp, YCO2, YH2O, OD)

        rad = 0.001 * phi * emissivity * StefanBoltzmann * utemp ^ 4 'kW/m2
        'reradiation from the surface not included

    End Sub
	
    Public Function ViewFactor_flame_floor(ByVal r1 As Double, ByVal R2 As Double, ByVal h As Double) As Double
        '*  ====================================================================
        '*  This function returns the value of the radiation view factor for the
        '*  compartment floor receiving radiation from a cylindrical plume.
        '*  The analytical expression is taken from the
        '*  SFPE handbook
        '*  assumes the base of the plume is at floor level
        '* r1 = radius of plume cylinder
        '* r2 = relevant distance from the plume centreline
        '* h = flame height
        '*  ====================================================================

        Dim L, R, X As Double

        'some useful ratios
        If R2 <= r1 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object ViewFactor_flame_floor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ViewFactor_flame_floor = 1
            Exit Function
        End If
        If R2 <= 0 Then Exit Function
        If R2 > 0 Then R = r1 / R2
        If R2 > 0 Then L = h / R2
        X = Sqrt((1 + L ^ 2 + R ^ 2) ^ 2 - 4 * R ^ 2)

    End Function
End Module