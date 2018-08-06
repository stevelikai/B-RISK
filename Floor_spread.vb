Option Strict Off
Option Explicit On
Module Floor_spread
	
	
	
	Private Sub HRR_floorpeak(ByVal room As Integer, ByVal NetFlux_F As Single, ByRef QDotF As Double)
		'====================================================================
		' Determines the peak heat release rate from the compartment lining materials
		'
		'====================================================================
		Dim s() As Double
		Dim index() As Integer
        'Dim x1 As Single
        'Dim x2 As Single
		Dim x4 As Single
		Dim curve As Integer
        'Dim YW() As Double
        'Dim YC() As Double
        'Dim x3 As Single
		Dim YF() As Double
		Dim ExtFlux() As Double
		Dim j As Integer
		
		If CurveNumber_F(room) = 0 Then Exit Sub
		ReDim YF(CurveNumber_F(room))
		
		If CurveNumber_F(room) > 0 Then
			x4 = NetFlux_F
			If x4 < ExtFlux_F(1, room) Then x4 = ExtFlux_F(1, room)
			If x4 > ExtFlux_F(CurveNumber_F(room), room) Then x4 = ExtFlux_F(CurveNumber_F(room), room)
		End If
		
		'get the peak hrr for each cone input curve
		
		For curve = 1 To CurveNumber_F(room)
			YF(curve) = PeakFloorHRR(curve, room)
		Next curve
		
		
		ReDim ExtFlux(CurveNumber_F(room))
		For j = 1 To CurveNumber_F(room)
			ExtFlux(j) = ExtFlux_F(j, room)
		Next j
		
		ReDim s(CurveNumber_F(room))
		ReDim index(CurveNumber_F(room))
		If CurveNumber_F(room) > 1 Then
			Call CSCOEF(CurveNumber_F(room), ExtFlux, YF, s, index)
			Call CSFIT1(CurveNumber_F(room), ExtFlux, YF, s, index, x4, QDotF)
		Else
			QDotF = YF(1)
		End If
		If x4 > NetFlux_F Then
			QDotF = QDotF - QDotF * (x4 - NetFlux_F) / x4
		End If
		
	End Sub
	Public Function Floor_YPfront2(ByVal vheight As Double, ByVal room As Integer, ByVal YP As Double, ByVal YB As Single, ByVal QW As Double, ByVal QF As Double) As Object
		'*****************************************************************
		'*  Runge-Kutta function for the position of the Y pyrolysis front
		'   created 23/08/2001  C Wade
		'*****************************************************************
		
        Dim Qig, tig, qstar As Double
        Dim YFlame As Double
		
		tig = 100000
		If IgnCorrelation = vbJanssens Then
			If FloorConeDataFile(room) <> "null.txt" Then tig = PI / 4 * ThermalInertiaFloor(room) * ((IgTempF(room) - FloorTemp(room, stepcount)) / (QSpread)) ^ 2
		Else
			If FloorConeDataFile(room) <> "null.txt" Then tig = FloorFTP(room) / (QSpread - FloorQCrit(room)) ^ Floorn(room)
		End If
		
		If room = fireroom Then
			
			Qig = (BurnerFlameLength / FlameAreaConstant) ^ (1 / FlameLengthPower)
			
			'work out Yflame assuming flames are contiguous (Yflame should be the position of the flame tip relative to floor level)
			qstar = (HRR_total) / (1110 * BurnerWidth ^ (5 / 2))
			
			If FireLocation(1) = 0 Or FireLocation(1) = 1 Then
				'assumes burner in centre of room or against wall
				YFlame = -1.02 * BurnerDiameter + 0.235 * (HRR_total) ^ (2 / 5)
				If YFlame < 0 Then YFlame = 0
			Else
				'for corner configuration
				'wind-aided flame spread not possible in corner location
				YFlame = 0
			End If
			
			If YB < BurnerFlameLength Then
				'burner and flame continuous - use yflame from above
			Else
				'burner and wall flame are separated
				If YP > YB Then
					YFlame = YB + FlameAreaConstant * (QF * (YP - YB)) ^ FlameLengthPower
				Else
					YFlame = YB
				End If
			End If
		Else
			'room not fireroom
			If YP > YB Then
				YFlame = YB + FlameAreaConstant * (QF * (YP - YB)) ^ FlameLengthPower
			Else
				YFlame = YB
			End If
		End If
		
		If tig > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Floor_YPfront2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Floor_YPfront2 = (YFlame - YP) / tig
			Exit Function
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Floor_YPfront2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Floor_YPfront2 = 1000 'very fast
		End If
		
	End Function
	Public Function Floor_YBfront(ByVal room As Integer, ByVal YP As Double, ByVal YB As Double, ByVal QF As Double) As Object
		'**********************************************************************
		'  first order diferential equation to solve for the position of the
		'  burn out front
		'   created 23/8/2001   C Wade
		'**********************************************************************
		Dim tb As Double
		tb = 0
		If QF > 0 Then tb = AreaUnderFloorCurve(room) / QF 'burnout time applicable to wall material
		
		If tb > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Floor_YBfront. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Floor_YBfront = (YP - YB) / tb
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Floor_YBfront. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Floor_YBfront = 0
		End If
	End Function
	Function Floor_XPfront(ByVal room As Integer) As Object
		'*****************************************************************
		'*  Runge-Kutta function for the position of the X pyrolysis front
		'*  Revised 19/4/2002 Colleen Wade
		'*****************************************************************
		
		If FloorTemp(room, stepcount) >= FloorTSMin(room) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Floor_XPfront. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Floor_XPfront = FloorFlameSpreadParameter(room) / (ThermalInertiaFloor(room) * (IgTempF(room) - FloorTemp(room, stepcount)) ^ 2)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Floor_XPfront. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Floor_XPfront = 0
		End If
		
	End Function
	
    Public Sub floor_fireroom(ByVal i As Integer, ByVal sec As Double, ByRef fstart(,) As Double, ByRef X() As Double, ByRef V() As Double, ByRef fvelocity(,) As Double)
        '==================================================================
        ' set up the arrays for the floor flame spread in the fire room
        ' and call the ode solver
        '==================================================================

        'solve the flame spread ODE's for the fireroom
        If sec >= FloorIgniteTime(fireroom) Then
            X(1) = fstart(fireroom, 1) 'windaided floor
            X(3) = fstart(fireroom, 3) 'burnout floor
            If sec >= FloorIgniteTime(fireroom) Then X(2) = fstart(fireroom, 2) 'lateral floor
            V(1) = fvelocity(fireroom, 1)
            If sec >= FloorIgniteTime(fireroom) Then V(2) = fvelocity(fireroom, 2)
            V(3) = fvelocity(fireroom, 3)

            '*****************************************************************
            Call ODE_Spread_Solver(0, fireroom, 0, V, X) 'returns arrays velocity and xstart
            '*****************************************************************

            fstart(fireroom, 1) = X(1)
            fstart(fireroom, 3) = X(3)
            If sec >= FloorIgniteTime(fireroom) Then fstart(fireroom, 2) = X(2)
            fvelocity(fireroom, 1) = V(1)
            If sec >= FloorIgniteTime(fireroom) Then fvelocity(fireroom, 2) = V(2)
            fvelocity(fireroom, 3) = V(3)
            'returned values
            If i <> NumberTimeSteps + 1 Then
                'limit the maximum length of the Y fronts
                Yf_pyrolysis(fireroom, i + 1) = fstart(fireroom, 1) 'Y pyrolsis front
                If Yf_pyrolysis(fireroom, i + 1) < BurnerFlameLength Then
                    Yf_pyrolysis(fireroom, i + 1) = BurnerFlameLength
                End If
                fstart(fireroom, 1) = Yf_pyrolysis(fireroom, i + 1)
                Yf_burnout(fireroom, i + 1) = fstart(fireroom, 3) 'Y burnout front
                fstart(fireroom, 3) = Yf_burnout(fireroom, i + 1)
                'FlameVelocity(fireroom, 1, i + 1) = FloorVelocity(fireroom, 1) 'upward
                'FlameVelocity(fireroom, 2, i + 1) = FloorVelocity(fireroom, 2) 'lateral
                'FlameVelocity(fireroom, 3, i + 1) = FloorVelocity(fireroom, 3) 'burnout
                Xf_pyrolysis(fireroom, i + 1) = fstart(fireroom, 2) 'X pyrolysis front
                If Xf_pyrolysis(fireroom, i + 1) > RoomWidth(fireroom) Then
                    Xf_pyrolysis(fireroom, i + 1) = RoomWidth(fireroom)
                End If
                fstart(fireroom, 2) = Xf_pyrolysis(fireroom, i + 1)
            End If

            'record the time when the Y_pyrolysis front reaches the ceiling
            'If Y_pyrolysis(fireroom, i + 1) >= RoomHeight(fireroom) And YPFlag = 0 Then
            '    timeH = tim(i, 1)
            '    YPFlag = 1
            'End If
        Else
            'if the floor has not yet ignited
            Yf_pyrolysis(fireroom, i + 1) = Yf_pyrolysis(fireroom, i)
            Xf_pyrolysis(fireroom, i + 1) = Xf_pyrolysis(fireroom, i)
            Yf_burnout(fireroom, i + 1) = Yf_burnout(fireroom, i)
            'FlameVelocity(fireroom, 1, i + 1) = 0
            'FlameVelocity(fireroom, 2, i + 1) = 0
            'FlameVelocity(fireroom, 3, i + 1) = 0
        End If
    End Sub
End Module