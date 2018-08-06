Option Strict Off
Option Explicit On
Module DAVE
	Sub CSFIT1(ByVal n As Integer, ByRef X() As Double, ByRef Y() As Double, ByRef s() As Double, ByRef index() As Integer, ByVal x1 As Double, ByRef SX As Double)
		
		' Cublic spline interpolation subroutine
		
		' Input
		
		'  N       = number of X and Y data points
		'  X()     = array of X data points (N rows by 1 column)
		'  Y()     = array of Y data points (N rows by 1 column)
		'  S()     = array of cubic spline coefficients
		'            (N rows by 1 column)
		'  INDEX() = array of indices (N rows by 1 column)
		'  X1      = X data value to fit
		
		' Output
		
		'  SX = cubic spline interpolated value for X
		Static IL, Ii, i, L, ILP1 As Integer
		Static A, b As Double
		Static HL As Double
		
		For i = 2 To n
			Ii = index(i)
			If (x1 <= X(Ii)) Then Exit For
		Next i
		
		L = i - 1
		IL = index(L)
		ILP1 = index(L + 1)
		
		A = X(ILP1) - x1
		b = x1 - X(IL)
		
		HL = X(ILP1) - X(IL)
		
		SX = A * s(L) * (A * A / HL - HL) / 6# + b * s(L + 1) * (b * b / HL - HL) / 6# + (A * Y(IL) + b * Y(ILP1)) / HL
		
	End Sub
	
	Sub CSCOEF(ByVal n As Integer, ByRef X() As Double, ByRef Y() As Double, ByRef s() As Double, ByRef index() As Integer)
		
		' Cublic spline coefficients subroutine
		
		' Input
		
		'  N   = number of X and Y data points
		'  X() = array of X data points (N rows by 1 column)
		'  Y() = array of Y data points (N rows by 1 column)
		
		' Output
		
		'  S()     = array of cubic spline coefficients
		'            (N rows by 1 column)
		'  INDEX() = array of indices (N rows by 1 column)
		
		Static Rho() As Double
		Static Tau() As Double
		Static NM1 As Integer
		Static IP1 As Integer
		Static Ii As Integer
		Static IJ As Integer
		Static IIM1, ITEMP, NM2, IIP1 As Integer
		Static HIM1, HI As Double
		Static temp, D As Double
		Static i, IB, j As Integer
		ReDim Rho(n)
		ReDim Tau(n)
		
		NM1 = n - 1
		
		For i = 1 To n
			index(i) = i
		Next i
		
		' ascending order data sort
		
		For i = 1 To NM1
			IP1 = i + 1
			For j = IP1 To n
				Ii = index(i)
				IJ = index(j)
				If (X(Ii) > X(IJ)) Then
					ITEMP = index(i)
					index(i) = index(j)
					index(j) = ITEMP
				End If
			Next j
		Next i
		
		NM2 = n - 2
		
		Rho(2) = 0#
		Tau(2) = 0#
		
		For i = 2 To NM1
			IIM1 = index(i - 1)
			Ii = index(i)
			IIP1 = index(i + 1)
			HIM1 = X(Ii) - X(IIM1)
			HI = X(IIP1) - X(Ii)
			temp = (HIM1 / HI) * (Rho(i) + 2#) + 2#
			Rho(i + 1) = -1# / temp
			D = 6# * ((Y(IIP1) - Y(Ii)) / HI - (Y(Ii) - Y(IIM1)) / HIM1) / HI
			Tau(i + 1) = (D - HIM1 * Tau(i) / HI) / temp
		Next i
		
		s(1) = 0#
		s(n) = 0#
		
		' compute cubic spline coefficients
		
		For i = 1 To NM2
			IB = n - i
			s(IB) = Rho(IB + 1) * s(IB + 1) + Tau(IB + 1)
		Next i
		
		Erase Rho
		Erase Tau
		
	End Sub
	
	
	
	Function hrr_dave(ByRef timenow As Double, ByRef floorflux_f As Single, ByRef floorflux_b As Single, ByRef QP_wall As Double, ByRef QP_wall_B As Double, ByRef QP_ceiling As Double, ByRef QNormalFB As Double, ByRef QNormalF As Double, ByRef QNormalW As Double, ByRef QNormalWB As Double, ByRef QNormalC As Double, ByRef QWall As Single, ByRef QCeiling As Single, ByRef QFloor As Single) As Single
		'====================================================================
		' Determines the heat release rate from the compartment lining materials
		' by storing the change in pyrolysis area for each time step and
		' using the elapsed time to determine the heat release rate for each
		' burning segment from the cone calorimeter input
		'
		' Revised Jan 20, 1997.  By: David LeBlanc and James Ierardi
		'   16/6/97 C A Wade
		'====================================================================
		
		Dim i As Integer
		Dim hrr As Double
		Dim Q, elapsed, burnarea, IO As Double
		Static TimeFactorW As Double
		Static TimefactorWB As Double
		Static TimefactorFB As Double
		Static TimeFactorC As Double
		Static TimeFactorF As Double
		Dim UpperTimeFactor As Double
		Dim LowerTimeFactor As Double
		On Error GoTo errorhandlerdave
		
		'new bit
		If FirstTime(fireroom) = True Then
			TimeFactorW = 1
			TimeFactorC = 1
			TimefactorWB = 1
			TimefactorFB = 1
			TimeFactorF = 1
			FirstTime(fireroom) = False
		End If
		
		LowerTimeFactor = 0
		UpperTimeFactor = 3
		
		IO = 0
		If AreaUnderWallCurve(fireroom) > 0 Then
			Do While WallIgniteStep(fireroom) > 0 And System.Math.Abs(AreaUnderWallCurve(fireroom) - IO) / AreaUnderWallCurve(fireroom) > 0.01 '1%
				If flagstop = 1 Then Exit Function
				'calculate the new area under the curve
				Call Integrate_New_HRR(fireroom, 1, 0, wallhigh, TimeFactorW, IO, QP_wall)
				
				If System.Math.Abs(AreaUnderWallCurve(fireroom) - IO) / AreaUnderWallCurve(fireroom) <= 0.01 Then Exit Do
				
				'compare with cone area under curve
				If IO > AreaUnderWallCurve(fireroom) Then
					'contract time scale
					UpperTimeFactor = TimeFactorW
					TimeFactorW = (UpperTimeFactor + LowerTimeFactor) / 2
					'calculate the new area under the curve
					Call Integrate_New_HRR(fireroom, 1, 0, wallhigh, TimeFactorW, IO, QP_wall)
				Else
					'expand time scale
					LowerTimeFactor = TimeFactorW
					TimeFactorW = (TimeFactorW + UpperTimeFactor) / 2
					Call Integrate_New_HRR(fireroom, 1, 0, wallhigh, TimeFactorW, IO, QP_wall)
				End If
				
				If UpperTimeFactor - LowerTimeFactor <= 0 Then
					UpperTimeFactor = UpperTimeFactor * 2
				End If
				If UpperTimeFactor > 10000 Then Exit Do
				If LowerTimeFactor < 0.0001 Then Exit Do
			Loop 
			
			LowerTimeFactor = 0
			UpperTimeFactor = 3
			IO = 0
			Do While QP_wall_B > 0 And System.Math.Abs(AreaUnderWallCurve(fireroom) - IO) / AreaUnderWallCurve(fireroom) > 0.01 '1%
				'calculate the new area under the curve
				Call Integrate_New_HRR(fireroom, 1, 0, wallhigh, TimefactorWB, IO, QP_wall_B)
				
				If System.Math.Abs(AreaUnderWallCurve(fireroom) - IO) / AreaUnderWallCurve(fireroom) <= 0.01 Then Exit Do
				
				'compare with cone area under curve
				If IO > AreaUnderWallCurve(fireroom) Then
					'contract time scale
					UpperTimeFactor = TimefactorWB
					TimefactorWB = (UpperTimeFactor + LowerTimeFactor) / 2
					'calculate the new area under the curve
					Call Integrate_New_HRR(fireroom, 1, 0, wallhigh, TimefactorWB, IO, QP_wall_B)
				Else
					'expand time scale
					LowerTimeFactor = TimefactorWB
					TimefactorWB = (TimefactorWB + UpperTimeFactor) / 2
					Call Integrate_New_HRR(fireroom, 1, 0, wallhigh, TimefactorWB, IO, QP_wall_B)
				End If
				
				If UpperTimeFactor - LowerTimeFactor <= 0 Then
					UpperTimeFactor = UpperTimeFactor * 2
				End If
				If UpperTimeFactor > 10000 Then Exit Do
				If LowerTimeFactor < 0.0001 Then Exit Do
			Loop 
		End If
		
		If AreaUnderCeilingCurve(fireroom) > 0 Then
			If WallConeDataFile(fireroom) <> CeilingConeDataFile(fireroom) Then
				LowerTimeFactor = 0
				UpperTimeFactor = 3
				IO = 0
				Do While CeilingIgniteStep(fireroom) > 0 And System.Math.Abs(AreaUnderCeilingCurve(fireroom) - IO) / AreaUnderCeilingCurve(fireroom) > 0.01 '1%
					'calculate the new area under the curve
					Call Integrate_New_HRR(fireroom, 2, 0, ceilinghigh, TimeFactorC, IO, QP_ceiling)
					
					If System.Math.Abs(AreaUnderCeilingCurve(fireroom) - IO) / AreaUnderCeilingCurve(fireroom) <= 0.01 Then Exit Do
					
					'compare with cone area under curve
					If IO > AreaUnderCeilingCurve(fireroom) Then
						'contract time scale
						UpperTimeFactor = TimeFactorC
						TimeFactorC = (UpperTimeFactor + LowerTimeFactor) / 2
						'calculate the new area under the curve
						Call Integrate_New_HRR(fireroom, 2, 0, ceilinghigh, TimeFactorC, IO, QP_ceiling)
					Else
						'expand time scale
						LowerTimeFactor = TimeFactorC
						TimeFactorC = (TimeFactorC + UpperTimeFactor) / 2
						Call Integrate_New_HRR(fireroom, 2, 0, ceilinghigh, TimeFactorC, IO, QP_ceiling)
					End If
					
					If UpperTimeFactor - LowerTimeFactor <= 0 Then
						UpperTimeFactor = UpperTimeFactor * 2
					End If
					If UpperTimeFactor > 10000 Then Exit Do
					If LowerTimeFactor < 0.0001 Then Exit Do
				Loop 
			Else
				TimeFactorC = TimeFactorW
			End If
		End If
		
		If AreaUnderFloorCurve(fireroom) > 0 Then
			If FloorConeDataFile(fireroom) <> WallConeDataFile(fireroom) Then
				LowerTimeFactor = 0
				UpperTimeFactor = 3
				IO = 0
				Do While FloorIgniteStep(fireroom) > 0 And System.Math.Abs(AreaUnderFloorCurve(fireroom) - IO) / AreaUnderFloorCurve(fireroom) > 0.01 '1%
					'calculate the new area under the curve
					Call Integrate_New_HRR(fireroom, 3, 0, floorhigh, TimeFactorF, IO, floorflux_f)
					
					If System.Math.Abs(AreaUnderFloorCurve(fireroom) - IO) / AreaUnderFloorCurve(fireroom) <= 0.01 Then Exit Do
					
					'compare with cone area under curve
					If IO > AreaUnderFloorCurve(fireroom) Then
						'contract time scale
						UpperTimeFactor = TimeFactorF
						TimeFactorF = (UpperTimeFactor + LowerTimeFactor) / 2
						'calculate the new area under the curve
						Call Integrate_New_HRR(fireroom, 3, 0, floorhigh, TimeFactorF, IO, floorflux_f)
					Else
						'expand time scale
						LowerTimeFactor = TimeFactorF
						TimeFactorF = (TimeFactorF + UpperTimeFactor) / 2
						Call Integrate_New_HRR(fireroom, 3, 0, floorhigh, TimeFactorF, IO, floorflux_f)
					End If
					
					If UpperTimeFactor - LowerTimeFactor <= 0 Then
						UpperTimeFactor = UpperTimeFactor * 2
					End If
					If UpperTimeFactor > 10000 Then Exit Do
					If LowerTimeFactor < 0.0001 Then Exit Do
				Loop 
			Else
				TimeFactorF = TimeFactorW
			End If
		End If
		
		'Debug.Print "Converged", TimeFactor
		'start at the timestep when the wall first ignited
		If WallIgniteTime(fireroom) < CeilingIgniteTime(fireroom) Then
			If WallIgniteTime(fireroom) < FloorIgniteTime(fireroom) Then
				i = WallIgniteStep(fireroom) + 1
			Else
				i = FloorIgniteStep(fireroom) + 1
			End If
		Else
			If CeilingIgniteTime(fireroom) < FloorIgniteTime(fireroom) Then
				i = CeilingIgniteStep(fireroom) + 1
			Else
				i = FloorIgniteStep(fireroom) + 1
			End If
		End If
		
		hrr = 0
		QWall = 0
		QCeiling = 0
		QFloor = 0
		Do While i < stepcount
			If flagstop = 1 Then Exit Function
			elapsed = timenow - tim(i, 1)
			burnarea = delta_area(fireroom, 1, i) + delta_area(fireroom, 2, i) + delta_area(fireroom, 3, i)
			If delta_area(fireroom, 1, i) > 0 Or delta_area(fireroom, 2, i) > 0 Then
				Call Get_Normal_Data2(fireroom, 1, elapsed, QNormalW, TimeFactorW)
				If WallConeDataFile(fireroom) <> CeilingConeDataFile(fireroom) Then
					If delta_area(fireroom, 2, i) > 0 Then
						Call Get_Normal_Data2(fireroom, 2, elapsed, QNormalC, TimeFactorC)
					Else
						QNormalC = 0
					End If
				Else
					QNormalC = QNormalW
				End If
				QWall = QWall + QP_wall * QNormalW * delta_area(fireroom, 1, i)
				QCeiling = QCeiling + QP_ceiling * QNormalC * delta_area(fireroom, 2, i)
			End If
			If delta_area(fireroom, 3, i) > 0 Then
				Call Get_Normal_Data2(fireroom, 3, elapsed, QNormalF, TimeFactorF)
				QFloor = QFloor + floorflux_f * QNormalF * delta_area(fireroom, 3, i)
			End If
			i = i + 1
		Loop 
		
		If WallIgniteStep(fireroom) > 0 And (WallConeDataFile(fireroom) <> "" And WallConeDataFile(fireroom) <> "null.txt") Then
			'Call Get_Normal_Data2(ByVal fireroom, 3, timenow - tim(WallIgniteStep(fireroom), 1), QNormalWB, TimefactorWB)
			Call Get_Normal_Data2(fireroom, 1, timenow - tim(WallIgniteStep(fireroom), 1), QNormalWB, TimefactorWB)
			QWall = QWall + QP_wall_B * QNormalWB * delta_area(fireroom, 1, WallIgniteStep(fireroom))
		End If
		
		'If FloorIgniteStep(fireroom) > 0 And (FloorConeDataFile(fireroom) <> "" And FloorConeDataFile(fireroom) <> "null.txt") Then
		'    QFloor = floorflux_f * RoomFloorArea(fireroom)
		'End If
		If FloorIgniteStep(fireroom) > 0 And (FloorConeDataFile(fireroom) <> "" And FloorConeDataFile(fireroom) <> "null.txt") Then
			Call Get_Normal_Data2(fireroom, 3, timenow - tim(FloorIgniteStep(fireroom), 1), QNormalFB, TimefactorFB)
			QFloor = QFloor + floorflux_b * QNormalFB * delta_area(fireroom, 3, FloorIgniteStep(fireroom))
		End If
		
		hrr_dave = QWall + QCeiling + QFloor
		Exit Function
		
errorhandlerdave: 
		
HRRerrorhandler: 
		MsgBox(ErrorToString(Err.Number) & "Dave")
		ERRNO = Err.Number
        'System.Windows.Forms.Cursor.Current = Default_Renamed
		
	End Function
	
	
    Public Function hrr_dave_nextroom(ByVal room As Integer, ByVal timenow As Double, ByVal QP_wall As Double, ByVal QP_wall_B As Double, ByVal QP_floor As Double, ByVal QP_floor_B As Double, ByVal QP_ceiling As Double, ByVal QNormalW As Double, ByVal QNormalWB As Double, ByVal QNormalC As Double, ByVal QNormalF As Double, ByVal QNormalFB As Double) As Double
        '====================================================================
        ' Determines the heat release rate from the ceiling in the adjacent room
        '
        ' 30/5/99
        '====================================================================

        Dim i As Short
        Dim QCeiling, hrr, QWall, QFloor As Double
        Dim IO, elapsed, burnarea As Double
        Static TimeFactorF, TimeFactorW, TimeFactorC, TimefactorWB, TimefactorFB As Double
        Dim UpperTimeFactor As Double
        Dim LowerTimeFactor As Double
        On Error GoTo errorhandlerdave

        'new bit
        If FirstTime(room) = True Then
            TimeFactorC = 1
            TimeFactorW = 1
            TimefactorFB = 1
            TimeFactorF = 1
            TimefactorWB = 1
            FirstTime(room) = False
        End If

        LowerTimeFactor = 0
        UpperTimeFactor = 3
        IO = 0

        If AreaUnderWallCurve(room) > 0 Then
            Do While WallIgniteStep(room) > 0 And System.Math.Abs(AreaUnderWallCurve(room) - IO) / AreaUnderWallCurve(room) > 0.01 '1%
                'calculate the new area under the curve
                Call Integrate_New_HRR(room, 1, 0, wallhigh, TimeFactorW, IO, QP_wall)

                If System.Math.Abs(AreaUnderWallCurve(room) - IO) / AreaUnderWallCurve(room) <= 0.01 Then Exit Do

                'compare with cone area under curve
                If IO > AreaUnderWallCurve(room) Then
                    'contract time scale
                    UpperTimeFactor = TimeFactorW
                    TimeFactorW = (UpperTimeFactor + LowerTimeFactor) / 2
                    'calculate the new area under the curve
                    Call Integrate_New_HRR(room, 1, 0, wallhigh, TimeFactorW, IO, QP_wall)
                Else
                    'expand time scale
                    LowerTimeFactor = TimeFactorW
                    TimeFactorW = (TimeFactorW + UpperTimeFactor) / 2
                    Call Integrate_New_HRR(room, 1, 0, wallhigh, TimeFactorW, IO, QP_wall)
                End If

                If UpperTimeFactor - LowerTimeFactor <= 0 Then
                    UpperTimeFactor = UpperTimeFactor * 2
                End If
                If UpperTimeFactor > 10000 Then Exit Do
                If LowerTimeFactor < 0.00001 Then Exit Do

            Loop

            LowerTimeFactor = 0
            UpperTimeFactor = 3
            IO = 0
            Do While QP_wall_B > 0 And System.Math.Abs(AreaUnderWallCurve(room) - IO) / AreaUnderWallCurve(room) > 0.01 '1%
                'calculate the new area under the curve
                Call Integrate_New_HRR(room, 1, 0, wallhigh, TimefactorWB, IO, QP_wall_B)

                If System.Math.Abs(AreaUnderWallCurve(room) - IO) / AreaUnderWallCurve(room) <= 0.01 Then Exit Do

                'compare with cone area under curve
                If IO > AreaUnderWallCurve(room) Then
                    'contract time scale
                    UpperTimeFactor = TimefactorWB
                    TimefactorWB = (UpperTimeFactor + LowerTimeFactor) / 2
                    'calculate the new area under the curve
                    Call Integrate_New_HRR(room, 1, 0, wallhigh, TimefactorWB, IO, QP_wall_B)
                Else
                    'expand time scale
                    LowerTimeFactor = TimefactorWB
                    TimefactorWB = (TimefactorWB + UpperTimeFactor) / 2
                    Call Integrate_New_HRR(room, 1, 0, wallhigh, TimefactorWB, IO, QP_wall_B)
                End If

                If UpperTimeFactor - LowerTimeFactor <= 0 Then
                    UpperTimeFactor = UpperTimeFactor * 2
                End If
                If UpperTimeFactor > 10000 Then Exit Do
                If LowerTimeFactor < 0.00001 Then Exit Do

            Loop
        End If

        If AreaUnderCeilingCurve(room) > 0 Then
            LowerTimeFactor = 0
            UpperTimeFactor = 3
            IO = 0
            Do While CeilingIgniteStep(room) > 0 And System.Math.Abs(AreaUnderCeilingCurve(room) - IO) / AreaUnderCeilingCurve(room) > 0.01 '1%
                'calculate the new area under the curve
                Call Integrate_New_HRR(room, 2, 0, ceilinghigh, TimeFactorC, IO, QP_ceiling)

                If System.Math.Abs(AreaUnderCeilingCurve(room) - IO) / AreaUnderCeilingCurve(room) <= 0.01 Then Exit Do

                'compare with cone area under curve
                If IO > AreaUnderCeilingCurve(room) Then
                    'contract time scale
                    UpperTimeFactor = TimeFactorC
                    TimeFactorC = (UpperTimeFactor + LowerTimeFactor) / 2
                    'calculate the new area under the curve
                    Call Integrate_New_HRR(room, 2, 0, ceilinghigh, TimeFactorC, IO, QP_ceiling)
                Else
                    'expand time scale
                    LowerTimeFactor = TimeFactorC
                    TimeFactorC = (TimeFactorC + UpperTimeFactor) / 2
                    Call Integrate_New_HRR(room, 2, 0, ceilinghigh, TimeFactorC, IO, QP_ceiling)
                End If

                If UpperTimeFactor - LowerTimeFactor <= 0 Then
                    UpperTimeFactor = UpperTimeFactor * 2
                End If
                If UpperTimeFactor > 10000 Then Exit Do
                If LowerTimeFactor < 0.00001 Then Exit Do
            Loop
        End If

        If AreaUnderFloorCurve(room) > 0 Then
            Do While FloorIgniteStep(room) > 0 And System.Math.Abs(AreaUnderFloorCurve(room) - IO) / AreaUnderFloorCurve(room) > 0.01 '1%
                'calculate the new area under the curve
                Call Integrate_New_HRR(room, 3, 0, floorhigh, TimeFactorF, IO, QP_floor)

                If System.Math.Abs(AreaUnderFloorCurve(room) - IO) / AreaUnderFloorCurve(room) <= 0.01 Then Exit Do

                'compare with cone area under curve
                If IO > AreaUnderFloorCurve(room) Then
                    'contract time scale
                    UpperTimeFactor = TimeFactorF
                    TimeFactorF = (UpperTimeFactor + LowerTimeFactor) / 2
                    'calculate the new area under the curve
                    Call Integrate_New_HRR(room, 3, 0, floorhigh, TimeFactorF, IO, QP_floor)
                Else
                    'expand time scale
                    LowerTimeFactor = TimeFactorF
                    TimeFactorF = (TimeFactorF + UpperTimeFactor) / 2
                    Call Integrate_New_HRR(room, 3, 0, floorhigh, TimeFactorF, IO, QP_floor)
                End If

                If UpperTimeFactor - LowerTimeFactor <= 0 Then
                    UpperTimeFactor = UpperTimeFactor * 2
                End If
                If UpperTimeFactor > 10000 Then Exit Do
                If LowerTimeFactor < 0.00001 Then Exit Do

            Loop

            LowerTimeFactor = 0
            UpperTimeFactor = 3
            IO = 0
            Do While QP_floor_B > 0 And System.Math.Abs(AreaUnderFloorCurve(room) - IO) / AreaUnderFloorCurve(room) > 0.01 '1%
                'calculate the new area under the curve
                Call Integrate_New_HRR(room, 3, 0, floorhigh, TimefactorFB, IO, QP_floor_B)

                If System.Math.Abs(AreaUnderFloorCurve(room) - IO) / AreaUnderFloorCurve(room) <= 0.01 Then Exit Do

                'compare with cone area under curve
                If IO > AreaUnderFloorCurve(room) Then
                    'contract time scale
                    UpperTimeFactor = TimefactorFB
                    TimefactorFB = (UpperTimeFactor + LowerTimeFactor) / 2
                    'calculate the new area under the curve
                    Call Integrate_New_HRR(room, 3, 0, floorhigh, TimefactorFB, IO, QP_floor_B)
                Else
                    'expand time scale
                    LowerTimeFactor = TimefactorFB
                    TimefactorFB = (TimefactorFB + UpperTimeFactor) / 2
                    Call Integrate_New_HRR(room, 3, 0, floorhigh, TimefactorFB, IO, QP_floor_B)
                End If

                If UpperTimeFactor - LowerTimeFactor <= 0 Then
                    UpperTimeFactor = UpperTimeFactor * 2
                End If
                If UpperTimeFactor > 10000 Then Exit Do
                If LowerTimeFactor < 0.00001 Then Exit Do

            Loop
        End If

        'start at the timestep when a lining first ignited
        If WallIgniteTime(room) < CeilingIgniteTime(room) Then
            If FloorIgniteTime(room) < WallIgniteTime(room) Then
                i = FloorIgniteStep(room) + 1
            Else
                i = WallIgniteStep(room) + 1
            End If
        Else
            If FloorIgniteTime(room) < CeilingIgniteTime(room) Then
                i = FloorIgniteStep(room) + 1
            Else
                i = CeilingIgniteStep(room) + 1
            End If
        End If

        hrr = 0
        QWall = 0
        QCeiling = 0
        QFloor = 0
        Do While i < stepcount
            If flagstop = 1 Then Exit Function
            elapsed = timenow - tim(i, 1)
            burnarea = delta_area(room, 1, i) + delta_area(room, 2, i) + delta_area(room, 3, i)
            If delta_area(room, 1, i) > 0 Or delta_area(room, 2, i) > 0 Then
                Call Get_Normal_Data2(room, 1, elapsed, QNormalW, TimeFactorW)
                If WallConeDataFile(room) <> CeilingConeDataFile(room) Then
                    If delta_area(room, 2, i) > 0 Then
                        Call Get_Normal_Data2(room, 2, elapsed, QNormalC, TimeFactorC)
                    Else
                        QNormalC = 0
                    End If
                Else
                    QNormalC = QNormalW
                End If
                QWall = QWall + QP_wall * QNormalW * delta_area(room, 1, i)
                QCeiling = QCeiling + QP_ceiling * QNormalC * delta_area(room, 2, i)
            End If
            If delta_area(room, 3, i) > 0 Then
                Call Get_Normal_Data2(room, 3, elapsed, QNormalF, TimeFactorF)
                QWall = QFloor + QP_floor * QNormalF * delta_area(room, 3, i)
            End If

            i = i + 1
        Loop

        If WallIgniteStep(room) > 0 And (WallConeDataFile(room) <> "" And WallConeDataFile(room) <> "null.txt") Then
            Call Get_Normal_Data2(room, 1, timenow - tim(WallIgniteStep(room), 1), QNormalWB, TimefactorWB)
            QWall = QWall + QP_wall_B * QNormalWB * delta_area(room, 1, WallIgniteStep(room))
        End If
        If FloorIgniteStep(room) > 0 And (FloorConeDataFile(room) <> "" And FloorConeDataFile(room) <> "null.txt") Then
            Call Get_Normal_Data2(room, 3, timenow - tim(FloorIgniteStep(room), 1), QNormalFB, TimefactorFB)
            QFloor = QFloor + QP_floor_B * QNormalFB * delta_area(room, 3, FloorIgniteStep(room))
        End If

        hrr_dave_nextroom = QWall + QCeiling + QFloor
        Exit Function

errorhandlerdave:

HRRerrorhandler:
        MsgBox(ErrorToString(Err.Number) & "hrr_Dave_nextroom")
        ERRNO = Err.Number
        'System.Windows.Forms.Cursor.Current = Default_Renamed

    End Function
End Module