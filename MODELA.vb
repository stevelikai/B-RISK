Option Strict Off
Option Explicit On
Module MODELA
	
    Sub Cone_Data(ByVal room As Long, ByVal CurveID As Integer, ByVal conecurve As Integer, ByVal tim As Double, ByRef Q As Double)
        '********************************************************
        '*  Get cone calorimeter data from data file at time=tim.
        '*
        '*  Revised: 8 October 1996 Colleen Wade
        '********************************************************

        Dim i As Integer
        'conecurve = 1

        If CurveID = 1 Then 'this is a wall
            Dim TempY(ConeNumber_W(conecurve, room)) As Double
            Dim Tempx(ConeNumber_W(conecurve, room)) As Double
            'return hrr at time tim as Q
            'ConeXW() is HRR
            'ConeYW() is time
            For i = 1 To ConeNumber_W(conecurve, room)
                ReDim Preserve TempY(i)
                TempY(i) = ConeYW(room, conecurve, i)
                Tempx(i) = ConeXW(room, conecurve, i)
            Next i
            Call Interpolate_D(TempY, Tempx, ConeNumber_W(conecurve, room), tim, Q)
        ElseIf CurveID = 2 Then  'this is a ceiling
            Dim TempY() As Double
            Dim TempX() As Double
            ReDim TempY(ConeNumber_C(conecurve, room))
            ReDim TempX(ConeNumber_C(conecurve, room))
            'return hrr at time tim as Q
            'ConeXC() is HRR
            'ConeYC() is time
            For i = 1 To ConeNumber_C(conecurve, room)
                TempY(i) = ConeYC(room, conecurve, i)
                TempX(i) = ConeXC(room, conecurve, i)
            Next i
            'UPGRADE_WARNING: Couldn't resolve default property of object room. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object conecurve. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call Interpolate_D(TempY, TempX, ConeNumber_C(conecurve, room), tim, Q)
        Else 'this is a floor
            'UPGRADE_WARNING: Couldn't resolve default property of object room. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object conecurve. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Dim TempY() As Double
            Dim TempX() As Double
            ReDim TempY(ConeNumber_F(conecurve, room))
            'UPGRADE_WARNING: Couldn't resolve default property of object room. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object conecurve. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ReDim TempX(ConeNumber_F(conecurve, room))
            'return hrr at time tim as Q
            'ConeXC() is HRR
            'ConeYC() is time
            'UPGRADE_WARNING: Couldn't resolve default property of object room. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object conecurve. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For i = 1 To ConeNumber_F(conecurve, room)
                'UPGRADE_WARNING: Couldn't resolve default property of object conecurve. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object room. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                TempY(i) = ConeYF(room, conecurve, i)
                'UPGRADE_WARNING: Couldn't resolve default property of object conecurve. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object room. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                TempX(i) = ConeXF(room, conecurve, i)
            Next i
            'UPGRADE_WARNING: Couldn't resolve default property of object room. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object conecurve. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call Interpolate_D(TempY, TempX, ConeNumber_F(conecurve, room), tim, Q)
        End If
    End Sub
	
	Sub flame_velocity2(ByRef tim As Object, ByRef Tau As Object, ByRef TauW As Object, ByRef TauC As Object, ByRef A0 As Object, ByRef Q As Object, ByRef QCIntegral As Object)
		'********************************************************
		'*  Solve for the Flame Spread Rate by Iteration.
		'*
		'*  Revised: 14 February 1997 Colleen Wade
		'********************************************************
		
		Static spreadrate As Double
		Dim sum1, VLHS, vrhs, sum2 As Double
		Dim First As Short
		Dim upperbound, lowerbound, VA As Double
		Dim ErrorCode, n, T1 As Short
		Dim T2 As Double
		ReDim Preserve flamespread(stepcount)
		ReDim Preserve product(stepcount)
		
		On Error GoTo errorhandler
		
		If FlameVelocityFlag = 0 Then spreadrate = 0
		
		FlameVelocityFlag = 1
		First = 0
		
		'guess VLHS
		
		vrhs = 0#
		upperbound = spreadrate
		lowerbound = spreadrate
		sum1 = 0
		sum2 = 0
		
		Do Until System.Math.Abs(spreadrate - vrhs) < 0.01 And First = 1
			First = 1
			flamespread(stepcount) = spreadrate
			'UPGRADE_WARNING: Couldn't resolve default property of object Q. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			product(stepcount) = spreadrate * Q
			
			'Call Simpson(flamespread(), Stepcount, timestep, sum1)
			'Call Simpson(product(), Stepcount, timestep, sum2)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object tim. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object TauC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object TauW. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call Integrate_Velocity(flamespread, (TauW + TauC), (TauW + TauC + tim), sum1)
			'UPGRADE_WARNING: Couldn't resolve default property of object tim. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object TauC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object TauW. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call Integrate_Velocity(product, (TauW + TauC), (TauW + TauC + tim), sum2)
			
			'If sum1 < 0 Or sum2 < 0 Then Stop
			
			'UPGRADE_WARNING: Couldn't resolve default property of object Tau. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Q. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object A0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			vrhs = (FlameAreaConstant * (A0 * Q + sum2) - sum1) / Tau
			
			If spreadrate < vrhs Then
				'increase
				lowerbound = spreadrate
				upperbound = vrhs
			Else
				'reduce
				upperbound = spreadrate
				lowerbound = vrhs
			End If
			
			spreadrate = (upperbound + lowerbound) / 2
			
			'dont let spreadrate go -ve
			If spreadrate < 0 Then
				spreadrate = 0
				vrhs = 0
			End If
			
			If spreadrate > 500 Then
				'Stop
				spreadrate = 500
				Exit Do
			End If
			
		Loop 
		
		'Debug.Print tim + TauW + TauC, spreadrate
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QCIntegral. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QCIntegral = sum2
		
		Exit Sub
		
errorhandler: 
		MsgBox(ErrorToString(Err.Number) & " flame_velocity2")
		ERRNO = Err.Number
		Exit Sub
		
	End Sub
	
    Function Get_Burner(ByRef id As Short, ByRef tim As Object) As Double
        '*  ==============================================================================
        '*  Returns the heat release rate (kW) at a given time (sec) for the
        '*  burner. Linearly interpolates between data points.
        '*  For the room corner test, the heat release rate of the burner is defined as the first
        '*  fire object, input by the user. It can be as per ISO 9705 or something else.
        '*
        '*  tim = Time (seconds)
        '*  ID = an integer identifying the fire object ID (=1 for room corner burner)
        '*
        '*  Revised 3 October 1996 Colleen Wade
        '*  ==============================================================================

        Dim i As Short
        Dim b, A, C As Double

        'get the current heat release rate for object number 1.
        For i = 1 To NumberDataPoints(id) - 1
            If tim >= HeatReleaseData(1, i, id) Then
                If tim < HeatReleaseData(1, i + 1, id) Then
                    A = tim - HeatReleaseData(1, i, id)
                    b = HeatReleaseData(1, i + 1, id) - HeatReleaseData(1, i, id)
                    C = HeatReleaseData(2, i + 1, id) - HeatReleaseData(2, i, id)
                    'Get_Burner = DMax1(0.01, HeatReleaseData(2, i, id) + (a / B) * C)
                    If HeatReleaseData(2, i, id) + (A / b) * C <= 0.001 Then
                        Get_Burner = 0.001
                    Else
                        Get_Burner = HeatReleaseData(2, i, id) + (A / b) * C
                    End If
                    Exit Function
                End If
            End If
        Next i

        'check to see if time corresponds to last data point
        If NumberDataPoints(id) <> 0 Then
            If tim = HeatReleaseData(1, NumberDataPoints(id), id) Then
                If 0.001 < HeatReleaseData(2, NumberDataPoints(id), id) Then
                    Get_Burner = HeatReleaseData(2, NumberDataPoints(id), id)
                    'Get_Burner = HeatReleaseData(2, NumberDataPoints(ID), ID)
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object Get_Burner. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Get_Burner = 0.001
                End If
                Exit Function
            End If
        End If

        'The heat release rate has not been specified up to the specified time.
        'The last specified value of Q is used."
        If NumberDataPoints(id) <> 0 Then
            Get_Burner = DMax1(0.001, HeatReleaseData(2, NumberDataPoints(id), id))
        End If

    End Function
	
    Function HRR_ModelA(ByRef id As Short, ByRef tim As Object, ByRef Q As Object) As Object
        '******************************************************
        '*  Find the total heat release rate for use by Model A
        '*
        '*  QW = hrr per unit area for the wall area only (kW/m2).
        '*
        '*  Revised 29 March 1998 Colleen Wade
        '******************************************************

        Dim Tau, ceilinghrr As Double
        Dim QW, hrr, A0 As Double
        Static wallflag, ceilingflag As Object
        Static TauC, TauW As Double
        Static HRRlast, timlast, Qlast As Double
        Static Delay As Double

        If stepcount = 1 Then
            wallflag = 0
            ceilingflag = 0
            Delay = 0
        End If

        If tim = timlast And stepcount <> 1 Then
            HRR_ModelA = HRRlast
            Q = Qlast
            Exit Function
        End If

        If wallflag = 0 Then
            TauW = PI * ThermalInertiaWall(fireroom) * (IgTempW(fireroom) - LowerWallTemp(fireroom, 1)) ^ 2 / (4 * WallHeatFlux ^ 2)
        End If

        Tau = PI * ThermalInertiaCeiling(fireroom) * (IgTempC(fireroom) - CeilingTemp(fireroom, stepcount)) ^ 2 / (4 * CeilingHeatFlux ^ 2)

        'get heat release from the wall area
        If tim > TauW Then Call Cone_Data(fireroom, 1, 1, tim - TauW, QW)

        Qburner = Get_Burner(id, tim)

        'initial pyrolysis area in the ceiling
        A0 = FlameAreaConstant * (Qburner + WallAreaBurner * QW - 150)

        'calculate ignition time of ceiling
        If ceilingflag = 0 And A0 > 0 Then
            TauC = Tau
        ElseIf ceilingflag = 0 And A0 <= 0 Then
            'since a0 is not positive, delay ignition of ceiling???
            TauC = tim 'a long value
            Delay = tim - TauW
        End If

        If tim <= TauW Then
            'only the burner
            hrr = Qburner
        ElseIf tim > TauW And tim <= (TauC + TauW + Delay) Then
            wallflag = 1
            'burner and the wall lining
            hrr = Qburner + WallAreaBurner * QW
        ElseIf tim > (TauC + TauW + Delay) Then
            If ceilingflag = 0 Then
                ceilingflag = 1
            End If

            ceilinghrr = QC_ModelA(tim - (TauW + TauC + Delay), CeilingTemp(fireroom, stepcount - 1), Tau, TauW, TauC, A0)

            hrr = Qburner + WallAreaBurner * QW + ceilinghrr
        End If

        HRR_ModelA = hrr

        'max hrr to use in the plume correlation
        If FireLocation(id) = 0 Then
            'centre of room, assume walls not burning
            Q = Qburner
        ElseIf FireLocation(id) = 1 Then
            'against wall
            Q = Qburner + BurnerWidth * (layerheight(fireroom, stepcount) - FireHeight(id)) * QW
        Else
            'corner
            'hrr to use in the plume correlation
            Q = Qburner + 2 * BurnerWidth * (layerheight(fireroom, stepcount) - FireHeight(id)) * QW
        End If

        'Q can't be greater than the total hrr
        If Q > hrr Then Q = hrr

        'if the layer drops to the base of the fire, no entrainment
        If layerheight(fireroom, stepcount - 1) - FireHeight(id) < 0 Then
            Q = 0
        End If

        timlast = tim
        HRRlast = hrr
        Qlast = Q

    End Function
	
	Sub Integrate_Velocity(ByRef TempY() As Double, ByRef low As Single, ByRef high As Single, ByRef IO As Double)
		'************************************************************
		'*  A procedure to find the area under the array supplied
		'*
		'*  Revised 8 May 1997 Colleen Wade
		'************************************************************
		
		Dim i, n As Integer
		Dim sum As Double
		Dim ErrorCode As Integer
		Dim yinterp As Double
		Dim X() As Double
		Dim W() As Double
		n = 50 'polynomial order
		
		ReDim X(n + 500)
		ReDim W(n + 500)
		
		Call GAULEG2(low, high, X, W, n, ErrorCode)
		
		If ErrorCode <> 0 Then
			MsgBox("Internal error in routine GAULEG2")
			Exit Sub
		End If
		
		sum = 0
		
		For i = 1 To n
			Call Interpolate2(tim, TempY, stepcount, X(i), yinterp)
			sum = sum + W(i) * yinterp
		Next i
		
		IO = sum
		'Debug.Print "velocity integral = "; IO
		
	End Sub
	
	Function QC_ModelA(ByRef tim As Double, ByRef CeilingTemp As Object, ByRef Tau As Object, ByRef TauW As Object, ByRef TauC As Object, ByRef A0 As Double) As Object
		'*  ====================================================================
		'*  Find the heat release from the ceiling material
		'*
		'*  tim = time from ignition of ceiling material (sec)
		'*
		'*  Function calls: flame_velocity2; Cone_Data
		'*  Called by: HRR_ModelA
		'*
		'*  Created 22 May 1997 by Colleen Wade
		'*  ====================================================================
		
		Dim Q As Double
		Dim QCIntegral As Double
		
		'initial pyrolysis area cannot be less than 0
		If A0 < 0 Then A0 = 0
		
		'max initial pyrolysis area cannot be greater than the ceiling area
		A0 = DMin1(RoomLength(fireroom) * RoomWidth(fireroom), A0)
		
        Call Cone_Data(fireroom, 1, 1, tim, Q)
		
		'solve for the velocity spread rate (m2/s)
		Call flame_velocity2(tim, Tau, TauW, TauC, A0, Q, QCIntegral)
		
		'limit the maximum hrr from ceiling
		QC_ModelA = DMin1(QCIntegral + A0 * Q, Q * RoomLength(fireroom) * RoomWidth(fireroom))
		
	End Function
	
	Sub Read_Cone_Data(ByRef room As Integer, ByRef id As Integer, ByRef high As Single, ByRef opendatafile As Object, ByRef radflux() As Double, ByRef IgnitionTime() As Double, ByRef IgnPoints As Short, ByRef Peak() As Double, ByRef IgnitionTemp As Object, ByRef ThermalInertia As Object, ByRef HeatofGasification As Object, ByRef AreaUnderCurve As Object, ByRef QCritical As Object)
		'*****************************************************************
		'*  Reads in HRR data from the cone calorimeter data file
		'*
		'*  Revised 2 December 1997 Colleen Wade
		'*****************************************************************
		
        Dim i, curve As Integer
		Dim peakval As Double
        Dim Description As String = ""
        Dim Heading As String = ""
        Dim Action As Short = 0
		Dim Flux, secs, hrr, TTI As Double
        Dim dummy As String = ""
		Static maxc, maxw, maxf As Integer
		Static filename As String
		
        Try

 
            'if file exists, open it
            If Confirm_File(UserPersonalDataFolder & gcs_folder_ext & "\cone\" + opendatafile, readfile, 0) = 1 Then
                FileOpen(1, (UserPersonalDataFolder & gcs_folder_ext & "\cone\" + opendatafile), OpenMode.Input)

                For i = 1 To MaxConeCurves
                    If id = 1 Or id = 4 Then
                        PeakWallHRR(i, room) = 0
                    ElseIf id = 2 Or id = 5 Then
                        PeakCeilingHRR(i, room) = 0
                    ElseIf id = 3 Or id = 6 Then
                        PeakFloorHRR(i, room) = 0
                    End If
                Next i

                high = 0

                'get data from file
                Input(1, Description)
                If id = 1 Or id = 4 Then
                    Input(1, dummy)
                    Input(1, dummy)
                    If IsNumeric(dummy) = True Then
                        CurveNumber_W(room) = CShort(dummy)
                    Else
                        MsgBox("The cone data file is not in the correct format. The number of heat release curves is required to be described as a two character string (line 2).")
                        Exit Sub
                    End If
                    ReDim Preserve ExtFlux_W(MaxConeCurves, NumberRooms)
                    ReDim Preserve ConeNumber_W(MaxConeCurves, NumberRooms)
                    maxw = 0
                    For curve = 1 To CurveNumber_W(room)
                        Input(1, dummy)
                        Input(1, ExtFlux_W(curve, room))
                        Input(1, dummy)
                        Input(1, ConeNumber_W(curve, room))
                        If ConeNumber_W(curve, room) > maxw Then maxw = ConeNumber_W(curve, room)
                        If curve = 1 And room = 1 Then
                            ReDim ConeXW(MaxNumberRooms, MaxConeCurves, maxw)
                            ReDim ConeYW(MaxNumberRooms, MaxConeCurves, maxw)
                        Else
                            If Not IsNothing(ConeYW) Then
                                If maxw < UBound(ConeYW, 3) Then maxw = UBound(ConeYW, 3) 'added 1/9/2005
                            End If

                            ReDim Preserve ConeXW(MaxNumberRooms, MaxConeCurves, maxw)
                            ReDim Preserve ConeYW(MaxNumberRooms, MaxConeCurves, maxw)
                        End If
                        Input(1, Heading)
                        For i = 1 To ConeNumber_W(curve, room)
                            'input time, HRR
                            Input(1, secs)
                            Input(1, hrr)
                            'sensitivity analysis for hrr
                            'hrr = 1# * hrr
                            ConeYW(room, curve, i) = secs
                            ConeXW(room, curve, i) = hrr

                            'If curve = 1 Then
                            'keep track of the peak heat release rate
                            If ConeXW(room, curve, i) > PeakWallHRR(curve, room) Then PeakWallHRR(curve, room) = ConeXW(room, curve, i)
                            'keep track of max time
                            If ConeYW(room, curve, i) > high Then high = ConeYW(room, curve, i)
                            'End If
                        Next i
                    Next curve
                ElseIf id = 2 Or id = 5 Then
                    Input(1, dummy)
                    Input(1, dummy)
                    If IsNumeric(dummy) = True Then
                        CurveNumber_C(room) = CShort(dummy)
                    Else
                        MsgBox("The cone data file is not in the correct format. The number of heat release curves is required to be described as a two character string (line 2).")
                        Exit Sub
                    End If

                    ReDim Preserve ExtFlux_C(MaxConeCurves, NumberRooms)
                    ReDim Preserve ConeNumber_C(MaxConeCurves, NumberRooms)

                    For curve = 1 To CurveNumber_C(room)
                        Input(1, dummy)
                        Input(1, ExtFlux_C(curve, room))
                        Input(1, dummy)
                        Input(1, ConeNumber_C(curve, room))
                        If curve = 1 And room = 1 Then

                            maxc = ConeNumber_C(curve, room)

                            ReDim ConeXC(NumberRooms, MaxConeCurves, maxc)
                            ReDim ConeYC(NumberRooms, MaxConeCurves, maxc)
                        Else
                            If maxc < ConeNumber_C(curve, room) Then maxc = ConeNumber_C(curve, room)

                            If Not IsNothing(ConeYC) Then
                                If maxc < UBound(ConeYC, 3) Then maxc = UBound(ConeYC, 3) 'added 1/9/2005
                            End If

                            ReDim Preserve ConeXC(NumberRooms, MaxConeCurves, maxc)
                            ReDim Preserve ConeYC(NumberRooms, MaxConeCurves, maxc)
                        End If
                        Input(1, Heading)
                        For i = 1 To ConeNumber_C(curve, room)
                            'input time, HRR
                            Input(1, secs)
                            Input(1, hrr)
                            ConeYC(room, curve, i) = secs
                            ConeXC(room, curve, i) = hrr
                            'sensitivity analysis for hrr
                            'hrr = 1# * hrr
                            'If curve = 1 Then
                            If ConeXC(room, curve, i) > PeakCeilingHRR(curve, room) Then PeakCeilingHRR(curve, room) = ConeXC(room, curve, i)
                            'keep track of max time
                            If ConeYC(room, curve, i) > high Then high = ConeYC(room, curve, i)
                            'End If
                        Next i
                    Next curve
                ElseIf id = 3 Or id = 6 Then  'floor
                    Input(1, dummy)
                    Input(1, dummy)
                    If IsNumeric(dummy) = True Then
                        CurveNumber_F(room) = CShort(dummy)
                    Else
                        MsgBox("The cone data file is not in the correct format. The number of heat release curves is required to be described as a two character string (line 2).")
                        Exit Sub
                    End If
                    ReDim Preserve ExtFlux_F(MaxConeCurves, NumberRooms)
                    ReDim Preserve ConeNumber_F(MaxConeCurves, NumberRooms)
                    For curve = 1 To CurveNumber_F(room)
                        Input(1, dummy)
                        Input(1, ExtFlux_F(curve, room))
                        Input(1, dummy)
                        Input(1, ConeNumber_F(curve, room))
                        If curve = 1 And room = 1 Then
                            maxf = ConeNumber_F(curve, room)
                            ReDim ConeXF(NumberRooms, MaxConeCurves, maxf)
                            ReDim ConeYF(NumberRooms, MaxConeCurves, maxf)
                        Else

                            If maxf < ConeNumber_F(curve, room) Then maxf = ConeNumber_F(curve, room)

                            ReDim Preserve ConeXF(NumberRooms, MaxConeCurves, maxf)
                            ReDim Preserve ConeYF(NumberRooms, MaxConeCurves, maxf)
                        End If
                        Input(1, Heading)
                        For i = 1 To ConeNumber_F(curve, room)
                            'input time, HRR
                            Input(1, secs)
                            Input(1, hrr)
                            ConeYF(room, curve, i) = secs
                            ConeXF(room, curve, i) = hrr
                            'If curve = 1 Then
                            If ConeXF(room, curve, i) > PeakFloorHRR(curve, room) Then PeakFloorHRR(curve, room) = ConeXF(room, curve, i)
                            'keep track of max time
                            If ConeYF(room, curve, i) > high Then high = ConeYF(room, curve, i)
                            'End If
                        Next i
                    Next curve

                End If

                Input(1, Description)
                Input(1, dummy)
                Input(1, IgnPoints)

                If IgnPoints > 0 Then
                    If IgnPoints > 1 Then
                        Input(1, dummy)
                        For i = 1 To IgnPoints
                            'input radiant flux, time to ignition
                            Input(1, Flux)
                            Input(1, TTI)
                            Input(1, peakval)
                            'sensitivity analysis for hrr
                            'peakval = 1# * peakval
                            'sensitivity analysis for ignition time
                            'TTI = 1# * TTI

                            radflux(i) = Flux
                            Peak(i) = peakval
                            IgnitionTime(i) = TTI
                            'If TTI <> 0 Then IgnitionTime(i) = (1 / TTI) ^ .547
                        Next i
                    Else
                        'need at least three flux - time pairs for a sensible regression
                        MsgBox("The cone data file must contain at least two pairs of heat flux - time to ignition data.", MB_ICONSTOP)
                        FileClose(1)
                        Exit Sub
                    End If
                Else
                    'input radiant flux, time to ignition etc
                    Input(1, dummy)
                    Input(1, IgnitionTemp)
                    Input(1, dummy)
                    Input(1, ThermalInertia)
                    Input(1, dummy)
                    Input(1, HeatofGasification)
                    Input(1, dummy)
                    Input(1, AreaUnderCurve)
                    Input(1, dummy)
                    Input(1, QCritical)
                End If
                FileClose(1)

            End If
            filename = opendatafile
            Exit Sub

ConfirmFileError:
            If Err.Number = 55 Then
                FileClose(1)
                MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in Read_Cone_Data")
                'Resume
            End If

            Action = File_Errors(Err.Number)

            Select Case Action
                Case 0
                    FileClose(1)
                    Exit Sub
                Case 1
                    MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in Read_Cone_Data")
                    'Resume Next
                Case 2
                    FileClose(1)
                    Exit Sub
                Case Else
                    FileClose(1)
                    Exit Sub
            End Select

        Catch ex As Exception
            FileClose(1)
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in Read_Cone_Data")

        End Try
	End Sub
	
	Sub RegressionPow(ByRef X() As Double, ByRef Y() As Double, ByRef n As Short, ByRef A As Double, ByRef b As Double, ByRef R As Double)
		
		'Dim i As Integer, pmarraybase As Integer
		'ReDim Xt(N), Yt(N)
		
		'         For i = pmarraybase To N
		'          Xt(i) = Log(X(i))
		'          Yt(i) = Log(Y(i))
		'        Next i
		'        RegressionL Xt(), Yt(), N, A, B, R
		'        A = Exp(A)
	End Sub
	
	Sub Simpson(ByRef Y() As Double, ByRef n As Object, ByRef h As Object, ByRef result As Object)
		'************************************************************************
		'  Simpson Subroutine Uses Simpson'S Integration Algorithm. The Function*
		'  Y(N) - The Array Containing The (Equally Spaced) Function Value      *
		'  N - Number Of Points In The Array. N Must Be Odd  .                  *
		'  H - The Independent Variable Mesh Size (=dx)                         *
		'  Sum - Is The Integration Answer
		'  From ProMath 2.0
		'************************************************************************
		
		Dim Nbegin, NPanel, Nend As Short
		Dim Nhalf As Double
		Dim i As Short
		
        NPanel = n - 1
		Nhalf = NPanel / 2
		Nbegin = 1
        result = 0.0#
		If (NPanel - 2 * Nhalf) = 0 Then GoTo 5
		
		' Number Of Panels Is Odd. Use 3/8 Rule On First Three Panels, 1/3 Rule On
		' The Rest.
        result = (3.0# * h / 8.0#) * (Y(1) + 3.0# * Y(2) + 3.0# * Y(3) + Y(4))
		Nbegin = 4
		
		' Apply 1/3 Rule - Add In First,second And Last Values.
5: 'UPGRADE_WARNING: Couldn't resolve default property of object n. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        result = result + (h / 3.0#) * (Y(Nbegin) + 4.0# * Y(Nbegin + 1) + Y(n))
		Nbegin = Nbegin + 2
        If Nbegin = n Then Return
		
		' The Pattern After Nbegin+2 Is Repititive. Get Nend, The Place To Stop
        Nend = n - 2
		For i = Nbegin To Nend Step 2
            result = result + (h / 3.0#) * (2.0# * Y(i) + 4.0# * Y(i + 1))
		Next i
		
	End Sub
	
    Function Thermal_Lining(ByRef material As Short, ByRef Property_Renamed As String) As Double
        '*  ======================================================================
        '*  This function returns the properties for a selected lining
        '*
        '*  Revised : 11 October 1996 Colleen Wade
        '*  ======================================================================

        Select Case material
            Case 1
                'painted gypsum paper-faced plasterboard
                Select Case Property_Renamed
                    Case "Cone Data File"
                        'UPGRADE_WARNING: Couldn't resolve default property of object Thermal_Lining. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Thermal_Lining = "EUMAT1.TXT"
                        Exit Function
                End Select
            Case 2
                'ordinary plywood
                Select Case Property_Renamed
                    Case "Cone Data File"
                        'UPGRADE_WARNING: Couldn't resolve default property of object Thermal_Lining. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Thermal_Lining = "EUMAT2.TXT"
                        Exit Function
                End Select
            Case 3
                'textile wallcovering on gypsum paper-faced plasterboard
                Select Case Property_Renamed
                    Case "Cone Data File"
                        'UPGRADE_WARNING: Couldn't resolve default property of object Thermal_Lining. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Thermal_Lining = "EUMAT3.TXT"
                        Exit Function
                End Select
            Case 4
                'Melamine-faced high density noncombustible board
                Select Case Property_Renamed
                    Case "Cone Data File"
                        'UPGRADE_WARNING: Couldn't resolve default property of object Thermal_Lining. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Thermal_Lining = "EUMAT4.TXT"
                        Exit Function
                End Select
            Case 5
                'plastic-faced steel sheet on mineral wool
                Select Case Property_Renamed
                    Case "Cone Data File"
                        'UPGRADE_WARNING: Couldn't resolve default property of object Thermal_Lining. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Thermal_Lining = "EUMAT5.TXT"
                        Exit Function
                End Select
            Case 6
                'FR Particleboard
                Select Case Property_Renamed
                    Case "Cone Data File"
                        'UPGRADE_WARNING: Couldn't resolve default property of object Thermal_Lining. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Thermal_Lining = "EUMAT6.TXT"
                        Exit Function
                End Select
            Case 7
                'polyurethane foam covered with steel sheets
                Select Case Property_Renamed
                    Case "Cone Data File"
                        'UPGRADE_WARNING: Couldn't resolve default property of object Thermal_Lining. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Thermal_Lining = "EUMAT9.TXT"
                        Exit Function
                End Select
            Case 8
                'pvc wallcarpet on gypsum paper-faced plasterboard
                Select Case Property_Renamed
                    Case "Cone Data File"
                        'UPGRADE_WARNING: Couldn't resolve default property of object Thermal_Lining. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Thermal_Lining = "EUMAT10.TXT"
                        Exit Function
                End Select
            Case 9
                'FR polystyrene foam
                Select Case Property_Renamed
                    Case "Cone Data File"
                        'UPGRADE_WARNING: Couldn't resolve default property of object Thermal_Lining. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Thermal_Lining = "EUMAT11.TXT"
                        Exit Function
                End Select
            Case Else
        End Select

    End Function
	
	
	Public Sub Read_dmpfile(ByRef opendatafile As Object)
		'*****************************************************************
		'*  Reads in HRR data from the cone calorimeter data file
		'*
		'*  Revised 2 December 1997 Colleen Wade
		'*****************************************************************
		
		Dim count As Integer
		Dim X As Single
		Dim Action As Short
		Dim sline(5000) As Double
		Dim starttime, endtime As String
		
		count = 1
		'if file exists, open it
		'UPGRADE_WARNING: Couldn't resolve default property of object opendatafile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Confirm_File(opendatafile, readfile, 0) = 1 Then
			FileOpen(1, (opendatafile), OpenMode.Input)
			On Error GoTo ConfirmFileError
			Do Until EOF(1)
				sline(count) = LineInput(1)
				count = count + 1
			Loop 
			FileClose(1)
			starttime = CStr(sline(1))
			endtime = CStr(sline(count - 1))
			X = 2 'line 1 in dump file
			
		End If
		Exit Sub
		
ConfirmFileError: 
		Action = File_Errors(Err.Number)
		Select Case Action
			Case 0
				Exit Sub
			Case 1
				Resume Next
			Case 2
				'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
				'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                'System.Windows.Forms.Cursor.Current = Default_Renamed
				Exit Sub
			Case Else
				'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
				'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                'System.Windows.Forms.Cursor.Current = Default_Renamed
				Exit Sub
		End Select
		
	End Sub
End Module