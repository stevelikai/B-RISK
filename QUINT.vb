Option Strict Off
Option Explicit On
Imports System.Math

Module QUINT
	
    Sub Derk_Spread_fireroom(ByVal flux_wall_AR As Double, ByVal fheight As Double, ByVal room As Integer, ByRef DYDX() As Double, ByRef Xstart() As Double, ByVal Nvariables As Integer, ByVal x1 As Double, ByVal x2 As Double, ByVal EPS As Double, ByVal H1 As Double, ByVal HMIN As Double, ByVal hnext As Double, ByVal kmax As Integer)
        '////////////////////////////////////////////////////////////////////////
        '   Derk2 Integrates A Set Of Diff. Eqs. Over An Interval Using
        '   Fifth-order Runge-kutta Method With Adaptive Step Size.
        '   This Routine Calls Rk5 To Do The Integration
        '*
        '*  Revised 4 April 1997 Colleen Wade
        '***************************************************************************
        Dim i, maxsteps, Kount, nstp As Integer
        Dim tiny As Double
        Dim flagcount As Integer
        Dim Y() As Double
        Dim Yscal() As Double

        flagcount = 0

        maxsteps = 10000 'good
        'maxsteps = 1000
        tiny = 1.0E-30
        Dim Hdid, X, h As Double
        ReDim Yscal(Nvariables)
        'ReDim DYDX(Nvariables)
        ReDim Y(Nvariables)

        X = x1
        h = Sign(H1, x2 - x1)
        Kount = 0
        For i = 1 To Nvariables : Y(i) = Xstart(i) : Next i
        For nstp = 1 To maxsteps
            Call Spread_Derivs_fireroom(flux_wall_AR, fheight, room, X, Y, DYDX, Nvariables)
            For i = 1 To Nvariables
                Yscal(i) = Abs(Y(i)) + Abs(h * DYDX(i)) + tiny
            Next i
            If (X + h - x2) * (X + h - x1) > 0.0# Then h = x2 - X
            Call Rk5_Spread(flux_wall_AR, fheight, room, Y, DYDX, Nvariables, X, h, EPS, Yscal, Hdid, hnext)
            If (X - x2) * (x2 - x1) >= 0.0# Then
                For i = 1 To Nvariables
                    Xstart(i) = Y(i)
                Next i
                If kmax Then Kount = Kount + 1
                Exit Sub
            End If

            If Abs(hnext) < HMIN Then
                flagcount = flagcount + 1
                If flagcount > 2 Then
                    flagstop = 1
                    MsgBox("Stepsize Smaller Than Minimum. Convergence is not reached and the simulation is terminated.")
                End If
                Exit Sub

            End If
            h = hnext

            If flagstop = 1 Then
                Exit For
            End If

        Next nstp

        If flagstop <> 1 Then
            Dim Message As String = Format(X, "0.0") & " sec: Too Many Steps in Derk_Spread_fireroom"

            frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
            For i = 1 To Nvariables
                Xstart(i) = Y(i)
            Next i
        End If



    End Sub
	
	Sub EHoG_Correlation(ByRef radflux() As Double, ByRef Peak() As Double, ByRef IgnPoints As Short, ByRef EffectiveHeatofCombustion As Object, ByRef HeatofGasification As Object)
		'******************************************************
		'*  Find the effetive heat of gasification for the lining material
		'*  from a correlation of cone peak hrr data using the
		'*  method of Quintiere.
		'*
		'*  Revised 23 February 1997 Colleen Wade
		'******************************************************
		
        Dim slope, Yintercept, R2 As Double
		
		'fit a linear regression line to the data, plotting peak hrr vs external flux
		Call RegressionL(radflux, Peak, IgnPoints, Yintercept, slope, R2)
		
		'determine the heat of gasification from the slope of the line and the effective heat of combustion
		
		'vary slope for sensitivity analysis
		
		If slope > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object EffectiveHeatofCombustion. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object HeatofGasification. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			HeatofGasification = EffectiveHeatofCombustion / slope
		Else
			If frmQuintiere.optUseOneCurve.Checked = True Then
				MsgBox("The heat of gasification was not able to be sensibly determined due to poor correlation of peak heat release rates")
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object HeatofGasification. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			HeatofGasification = 1#
		End If
		
	End Sub
    Sub GetLegendreRootsWeights(ByRef i As Integer, ByRef XM As Double, ByRef XL As Double, ByRef Z As Double, ByRef xx() As Double, ByRef x1 As Single, ByRef x2 As Single, ByRef X() As Double, ByRef W() As Double, ByRef n As Integer, ByRef ErrorCode As Integer)
        ' Input:
        Dim EPS As Double
        Dim DP As Double
        Dim P3, p1, p2, OneMinZ2 As Double
        Dim j As Integer
        Dim Z1 As Double

        EPS = 0.000000000000005 ' New accuracy
        ErrorCode = 0 'Init with OK

        Do  ' Do Loop
            p1 = 1.0#
            p2 = 0.0#
            For j = 1 To n
                P3 = p2
                p2 = p1
                p1 = ((2.0# * j - 1.0#) * Z * p2 - (j - 1.0#) * P3) / j
            Next j
            OneMinZ2 = 1.0# - Z * Z
            DP = -n * (Z * p1 - p2) / OneMinZ2
            Z1 = Z
            Z = Z1 - p1 / DP
        Loop While Abs(Z - Z1) > EPS ' Check convergence

        xx(i) = Z ' Store original root
        X(i) = XM - XL * Z
        X(n + 1 - i) = XM + XL * Z
        W(i) = 2.0# * XL / (OneMinZ2 * DP * DP)
        W(n + 1 - i) = W(i)

    End Sub
	Sub GAULEG2(ByRef x1 As Single, ByRef x2 As Single, ByRef X() As Double, ByRef W() As Double, ByRef n As Integer, ByRef ErrorCode As Integer)
		' Input:  Lower and upper limits of integration X1 and X2, and polynomial
		'         order N%.
		' Output: Returns arrays X() and W() of length N%, the
		'         abscissas and weights of the Gauss-Legendre N-point quadrature
		'         formula.
		'         ErrorCode=0  OK
		'         ErrorCode=1  Didn't converge
		' It is a modification of the PROMATH9.BAS routine, GAUSLEG.BAS, and uses
		' quadratic and cubic root guessing routines.
		'EPS = .00000000000003#                                      ' Old Promath accuracy.
		'*  Revised 8 May 1997 Colleen Wade
		'**************************************************
        Dim sum, convergence As Double
		Dim m, MLow As Integer
        Dim conv, XM, XL, Z As Double
		Dim i As Integer
		

		m = (n + 1) / 2
        Dim Xx(m) As Double ' Dimension original root variable
		If m < 3 Then MLow = m Else MLow = 3
		XM = 0.5 * (x2 + x1)
		XL = 0.5 * (x2 - x1)
		'  M1 = 0
		For i = 1 To MLow
			'       M1 = M1 + 1
            Z = Cos(PI * (i - 0.25) / (n + 0.5)) ' Guesses for 1st 3 roots
            Call GetLegendreRootsWeights(i, XM, XL, Z, Xx, x1, x2, X, W, n, ErrorCode) ' Get roots, weights
		Next i
		
		If m < 4 Then GoTo CheckAccuracy
		
		i = 4
				Z = 3# * Xx(i - 1) - 3# * Xx(i - 2) + Xx(i - 3) ' Quadratic extrapolation
        Call GetLegendreRootsWeights(i, XM, XL, Z, Xx, x1, x2, X, W, n, ErrorCode) ' Get roots, weights
		
		For i = 5 To m
						Z = 4# * Xx(i - 1) - 6# * Xx(i - 2) + 4# * Xx(i - 3) - Xx(i - 4)
			' CUBIC extrapolation for next root
            Call GetLegendreRootsWeights(i, XM, XL, Z, Xx, x1, x2, X, W, n, ErrorCode) ' Get roots, weights
        Next i

		System.Windows.Forms.Application.DoEvents()

CheckAccuracy:
		sum = 0
		For i = 1 To n
			sum = sum + W(i)
		Next i
		conv = 1# - (x2 - x1) / sum
		convergence = 0.00000000000001 ' Accuracy required in SUM test
        If Abs(conv) > convergence Then ' Make accuracy check
            ErrorCode = 1 ' Didn't converge
        End If
		
	End Sub
	
	
	
	
	Sub Integrate_HRR(ByVal room As Integer, ByRef id As Integer, ByRef low As Single, ByRef high As Single, ByRef IO As Double)
		'************************************************************
		'*  A procedure to find the area under the cone heat release
		'*  rate curve
		'*
		'*  Revised 8 May 1997 Colleen Wade
		'************************************************************
        Dim Tempx() As Double
        Dim TempY() As Double
        Dim i, n As Integer
		Dim sum As Double
		Dim curve, max As Short
		Dim ErrorCode As Integer
		Dim yinterp As Double
		Dim X() As Double
		Dim W() As Double
		n = 50
		
		ReDim X(n + 500)
		ReDim W(n + 500)
		'curve = 1  'uses the first curve only
		If id = 1 Or id = 4 Then
			max = CurveNumber_W(room)
		ElseIf id = 2 Or id = 5 Then 
			max = CurveNumber_C(room)
		ElseIf id = 3 Or id = 6 Then 
			max = CurveNumber_F(room)
		End If
		
		For curve = 1 To max
			If id = 1 Or id = 4 Then
                ReDim Tempx(ConeNumber_W(curve, room))
                ReDim TempY(ConeNumber_W(curve, room))
			ElseIf id = 2 Or id = 5 Then 
				ReDim Tempx(ConeNumber_C(curve, room))
				ReDim TempY(ConeNumber_C(curve, room))
			Else : id = 3 Or id = 6
				ReDim Tempx(ConeNumber_F(curve, room))
				ReDim TempY(ConeNumber_F(curve, room))
			End If
			
			Call GAULEG2(low, high, X, W, n, ErrorCode)
			
			If ErrorCode <> 0 Then
				MsgBox("Internal error in routine GAULEG2 in Integrate_HRR")
				flagstop = 1
				Exit Sub
			End If
			
			sum = 0
			If id = 1 Or id = 4 Then
				For i = 1 To ConeNumber_W(curve, room)
					Tempx(i) = ConeYW(room, curve, i) 'time
					TempY(i) = ConeXW(room, curve, i) 'hrr
				Next i
				For i = 1 To n
					Call Interpolate_D(Tempx, TempY, ConeNumber_W(curve, room), X(i), yinterp)
					sum = sum + W(i) * yinterp
				Next i
            ElseIf id = 2 Or id = 5 Then

                For i = 1 To ConeNumber_C(curve, room)
                    Tempx(i) = ConeYC(room, curve, i) 'time
                    TempY(i) = ConeXC(room, curve, i) 'hrr

                Next i
                For i = 1 To n
                    Call Interpolate_D(Tempx, TempY, ConeNumber_C(curve, room), X(i), yinterp)
                    sum = sum + W(i) * yinterp
                Next i

			ElseIf id = 3 Or id = 6 Then 
				For i = 1 To ConeNumber_F(curve, room)
					Tempx(i) = ConeYF(room, curve, i) 'time
					TempY(i) = ConeXF(room, curve, i) 'hrr
				Next i
				For i = 1 To n
					Call Interpolate_D(Tempx, TempY, ConeNumber_F(curve, room), X(i), yinterp)
					sum = sum + W(i) * yinterp
				Next i
			End If
			
			IO = IO + sum
            Debug.Print("Area Under HRR Curve = " & IO)
		Next curve
		IO = IO / max 'average area under the curve for all curves provided
	End Sub
	
	Sub ODE_Spread_Solver(ByVal flux_wall_AR As Double, ByVal room As Integer, ByVal fheight As Double, ByRef velocity() As Double, ByRef Xstart() As Double)
		'*  ====================================================================
		'*  Solve the ODE's for the flame spread using a 5th order Runge-Kutta numerical solution
		'*  'For use with Quintiere's room corner model
		'*  ====================================================================
		
		Dim Nvariables As Integer
		Dim x2, x1, EPS As Double
		Dim HMIN, H1, hnext As Double
		Dim kmax As Integer
		Dim NStep As Integer
		
		On Error GoTo ODEerrorhandler
		
		
		Nvariables = 6 'number of equations
        x1 = tim(stepcount, 1) 'initial time
        kmax = 5000 'max number of steps to be attempted in adapting step size
        x2 = x1 + Timestep  'final time

        '26/6/15
        EPS = 0.00001 'decimal place accuracy
        H1 = Timestep / 10  'suggested time step
        HMIN = 0 'minimum time step

        'previous settings
        'EPS = 0.000001 'decimal place accuracy
        'H1 = Timestep   'suggested time step
        'hnext = Timestep 'suggested time step
        'HMIN = 0.00000000001 'minimum time step

		Call Derk_Spread_fireroom(flux_wall_AR, fheight, room, velocity, Xstart, Nvariables, x1, x2, EPS, H1, HMIN, hnext, kmax)
		
		'Call RKV(ByVal Nvariables, ByVal x1, ByVal x2, ByVal H1, ByVal EPS, Xstart(), ByVal flux_wall_AR, ByVal fheight, ByVal room, velocity())
		
		Exit Sub
		
ODEerrorhandler: 
		MsgBox(ErrorToString(Err.Number) & " ODE_Spread_Solver")
		ERRNO = Err.Number

	End Sub
	
	Private Sub Pyrol_Area_NB(ByRef T As Double, ByRef WallArea As Double, ByRef CeilingArea As Double, ByRef FloorArea As Double)
		'=================================================================
		'=  This function calculates the wall pyrolysis area not including
		'=  burnout following Quintiere's description of the areas
		'=
		'=  Function called by HRR_Quintiere
		'=
		'=  Revised 14 August 1998 Colleen Wade
		'=================================================================
		
		Dim z_pyrol, jetdepth As Double
		Dim APJ1, AP1, APC1 As Double
		
		On Error GoTo AREAerrorhandler
		
		jetdepth = 0.12 'depth of ceiling jet 12% of fire-ceiling height
		'jetdepth = 0.08 'depth of ceiling jet 12% of fire-ceiling height
		
		If T >= WallIgniteTime(fireroom) Then
			
			'the wall has ignited
			If Y_pyrolysis(fireroom, stepcount) < RoomHeight(fireroom) Then
				'Case 1a
				'The ignitor burner has caused ignition of wall
				'The pyrolysis height is less than the ceiling height
				If FireLocation(1) = 2 Then
					'corner
					WallArea = 2 * ((Y_pyrolysis(fireroom, stepcount) - FireHeight(1)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (Y_pyrolysis(fireroom, WallIgniteStep(fireroom)) - FireHeight(1)) + 0.5 * (Y_pyrolysis(fireroom, stepcount) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))))
				ElseIf FireLocation(1) = 1 Then 
					'wall
					'WallArea = ((Y_pyrolysis(fireroom, stepcount) - FireHeight(1)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (Y_pyrolysis(fireroom, WallIgniteStep(fireroom)) - FireHeight(1)) + 0.5 * (Y_pyrolysis(fireroom, stepcount) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))))
					'ver2002.2
					WallArea = (Y_pyrolysis(fireroom, stepcount) - FireHeight(1)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + 2 * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (Y_pyrolysis(fireroom, WallIgniteStep(fireroom)) - FireHeight(1)) + (Y_pyrolysis(fireroom, stepcount) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom)))
				Else
					'centre of room
					WallArea = 0
				End If
				CeilingArea = 0
				
				'Exit Sub
			Else
				'Case 1b
				'The ignitor burner has caused ignition of wall
				'The pyrolysis height is greater than the ceiling height
				If CeilingIgniteFlag(fireroom) = 0 And FireLocation(1) <> 0 And CeilingEffectiveHeatofCombustion(fireroom) > 0 Then
					CeilingIgniteFlag(fireroom) = 1
                    'MDIFrmMain.ToolStripStatusLabel3.Text = "Ceiling in Room " & CStr(fireroom) & " has ignited at " & CStr(T) & " seconds."

                    Dim Message As String = "Ceiling in Room " & CStr(fireroom) & " has ignited at " & CStr(T) & " seconds."
                    MDIFrmMain.ToolStripStatusLabel2.Text = Message.ToString
                    If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                    CeilingIgniteTime(fireroom) = T
					CeilingIgniteStep(fireroom) = stepcount
				End If
				
				If FireLocation(1) = 2 Then
					AP1 = 2 * ((RoomHeight(fireroom) - FireHeight(1)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (Y_pyrolysis(fireroom, WallIgniteStep(fireroom)) - FireHeight(1)) + 0.5 * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (RoomHeight(fireroom) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom))))
				ElseIf FireLocation(1) = 1 Then 
					'AP1 = ((RoomHeight(fireroom) - FireHeight(1)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (Y_pyrolysis(fireroom, WallIgniteStep(fireroom)) - FireHeight(1)) + 0.5 * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (RoomHeight(fireroom) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom))))
					'v2002.2
					AP1 = 2 * ((RoomHeight(fireroom) - FireHeight(1)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (Y_pyrolysis(fireroom, WallIgniteStep(fireroom)) - FireHeight(1)) + 0.5 * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (RoomHeight(fireroom) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom)))) - (RoomHeight(fireroom) - FireHeight(1)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom))
				Else
					'centre of room
					AP1 = 0
				End If
				
				'The z_pyrolysis length cannot be greater than (1-0.12) * room height
				'z_pyrol = (1 - jetdepth) * (RoomHeight(fireroom) - FireHeight(1))
				z_pyrol = (1 - jetdepth) * (RoomHeight(fireroom))
				If z_pyrol > X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom)) Then
					z_pyrol = X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))
				End If
				If T = tim(stepcount, 1) Then
					Z_pyrolysis(fireroom, stepcount) = z_pyrol
				End If
				
				If z_pyrol > 0 Then
					'corner
					APJ1 = 2 * ((Y_pyrolysis(fireroom, stepcount) - RoomHeight(fireroom)) * (jetdepth * (RoomHeight(fireroom) - FireHeight(1))) + 0.5 * z_pyrol * (Y_pyrolysis(fireroom, stepcount) - RoomHeight(fireroom)) - 0.5 * (jetdepth * (RoomHeight(fireroom) - FireHeight(1)) + z_pyrol) ^ 2 * ((X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) / (RoomHeight(fireroom) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom)))))
				Else
					'corner
					APJ1 = 2 * (Y_pyrolysis(fireroom, stepcount) - RoomHeight(fireroom)) * (jetdepth * (RoomHeight(fireroom) - FireHeight(1)))
				End If
				If APJ1 < 0 Then APJ1 = 0
				
				WallArea = AP1 + APJ1
				If WallArea > 2 * (RoomHeight(fireroom) * RoomLength(fireroom) + RoomHeight(fireroom) * RoomWidth(fireroom)) Then
					WallArea = 2 * (RoomHeight(fireroom) * RoomLength(fireroom) + RoomHeight(fireroom) * RoomWidth(fireroom))
				End If
				CeilingArea = PI / 4 * (Y_pyrolysis(fireroom, stepcount) - RoomHeight(fireroom)) ^ 2
				If CeilingArea > RoomFloorArea(fireroom) Then CeilingArea = RoomFloorArea(fireroom)
			End If
		Else
			WallArea = 0
			CeilingArea = 0
		End If
		
		If T >= CeilingIgniteTime(fireroom) Then
			'the ceiling has ignited
			If FireLocation(1) = 0 Then
				'centre of room
				If CeilingArea < PI * (Y_pyrolysis(fireroom, stepcount) - RoomHeight(fireroom)) ^ 2 Then
					CeilingArea = PI * (Y_pyrolysis(fireroom, stepcount) - RoomHeight(fireroom)) ^ 2
				End If
			ElseIf FireLocation(1) = 2 Then 
				'corner location
				If CeilingArea < PI / 4 * (Y_pyrolysis(fireroom, stepcount) - RoomHeight(fireroom)) ^ 2 Then
					CeilingArea = PI / 4 * (Y_pyrolysis(fireroom, stepcount) - RoomHeight(fireroom)) ^ 2
				End If
			Else
				'wall location
				If CeilingArea < PI / 2 * (Y_pyrolysis(fireroom, stepcount) - RoomHeight(fireroom)) ^ 2 Then
					CeilingArea = PI / 2 * (Y_pyrolysis(fireroom, stepcount) - RoomHeight(fireroom)) ^ 2
				End If
			End If
			If CeilingArea > RoomFloorArea(fireroom) Then CeilingArea = RoomFloorArea(fireroom)
		End If
		
		'=======================
		' floor
		'=======================
		If T >= FloorIgniteTime(fireroom) Then
			'the floor has ignited
			'If frmQuintiere.chkFloorFlameSpread.Value = vbChecked Then
			If frmQuintiere.optWindAided.Checked = True Then
				'this is when we have wind-aided spread and maybe lateral spread as well similar to a wall material
				'ie the flame lies down parallel to the floor
				If FireLocation(1) = 0 Then
					'centre of room
					'wind-aided spread + lateral flow
					FloorArea = Yf_pyrolysis(fireroom, stepcount) * BurnerWidth + 2 * (Yf_pyrolysis(fireroom, FloorIgniteStep(fireroom)) * (Xf_pyrolysis(fireroom, stepcount) - Xf_pyrolysis(fireroom, FloorIgniteTime(fireroom))) + 0.5 * (Xf_pyrolysis(fireroom, stepcount) - Xf_pyrolysis(fireroom, FloorIgniteStep(fireroom))) * (Yf_pyrolysis(fireroom, stepcount) - Yf_pyrolysis(fireroom, FloorIgniteStep(fireroom))))
					'limit max area to 1/2 the room floor area
					If FloorArea > 0.5 * RoomFloorArea(fireroom) Then FloorArea = 0.5 * RoomFloorArea(fireroom)
				ElseIf FireLocation(1) = 2 Then 
					'corner location
					FloorArea = 0 'no wind-aided spread from a corner location
				Else
					'wall location
					'wind-aided spread + opposed flow
					FloorArea = Yf_pyrolysis(fireroom, stepcount) * BurnerWidth + 2 * (Yf_pyrolysis(fireroom, FloorIgniteStep(fireroom)) * (Xf_pyrolysis(fireroom, stepcount) - Xf_pyrolysis(fireroom, FloorIgniteTime(fireroom))) + 0.5 * (Xf_pyrolysis(fireroom, stepcount) - Xf_pyrolysis(fireroom, FloorIgniteStep(fireroom))) * (Yf_pyrolysis(fireroom, stepcount) - Yf_pyrolysis(fireroom, FloorIgniteStep(fireroom))))
					'limit max area to the room floor area
					If FloorArea > RoomFloorArea(fireroom) Then FloorArea = RoomFloorArea(fireroom)
				End If
			ElseIf frmQuintiere.optOpposedFlow.Checked = True Then 
				'opposed flow spread
				If FireLocation(1) = 0 Then
					'centre of room
					FloorArea = PI * Xf_pyrolysis(fireroom, stepcount) ^ 2 'area of a circle
				ElseIf FireLocation(1) = 2 Then 
					'corner location
					FloorArea = PI / 4 * Xf_pyrolysis(fireroom, stepcount) ^ 2 'area of a quarter-circle
				Else
					'against a wall
					FloorArea = PI / 2 * Xf_pyrolysis(fireroom, stepcount) ^ 2 'area of a semi-circle
				End If
				
				'limit max area to the room floor area
				If FloorArea > RoomFloorArea(fireroom) Then FloorArea = RoomFloorArea(fireroom)
				
			Else
				'the whole room floor area has ignited, assuming uniform radiation from a hot layer
				FloorArea = RoomFloorArea(fireroom)
			End If
		End If
		Exit Sub
		
AREAerrorhandler: 
		MsgBox(ErrorToString(Err.Number) & "in Pyrol_Area_NB")
		ERRNO = Err.Number

	End Sub
	
    Function RK_XPfront(ByVal room As Integer, ByRef runaway As Boolean) As Double
        '*****************************************************************
        '*  Runge-Kutta function for the position of the X pyrolysis front
        '*  Called by: Spread_Derivs
        '*****************************************************************
        runaway = False
        If Upperwalltemp(room, stepcount) >= WallTSMin(room) Then

            If PessimiseCombWall = True Then
                If (IgTempW(room) - Upperwalltemp(room, stepcount)) > 0 Then
                    RK_XPfront = WallFlameSpreadParameter(room) / (ThermalInertiaWall(room) * (IgTempW(room) - Upperwalltemp(room, stepcount)) ^ 2)
                Else
                    'everything ignites, runaway
                    RK_XPfront = 10 'let's limit it to 10 m/s
                    runaway = True
                End If
            Else
                If layerheight(room, stepcount) > wallVFSlimit Then
                    If (IgTempW(room) - LowerWallTemp(room, stepcount)) > 0 Then
                        RK_XPfront = WallFlameSpreadParameter(room) / (ThermalInertiaWall(room) * (IgTempW(room) - LowerWallTemp(room, stepcount)) ^ 2)
                    Else
                        'everything ignites, runaway
                        RK_XPfront = 10 'let's limit it to 10 m/s
                        runaway = True
                    End If
                Else
                    If (IgTempW(room) - Upperwalltemp(room, stepcount)) > 0 Then
                        RK_XPfront = WallFlameSpreadParameter(room) / (ThermalInertiaWall(room) * (IgTempW(room) - Upperwalltemp(room, stepcount)) ^ 2)
                    Else
                        'everything ignites, runaway
                        RK_XPfront = 10 'let's limit it to 10 m/s
                        runaway = True
                    End If
                End If

            End If

        Else
            RK_XPfront = 0
        End If
        If RK_XPfront > 0.1 Then RK_XPfront = 0.1

    End Function
    Function RK_XPfront2(ByVal room As Integer, ByRef runaway As Boolean) As Double
        '*****************************************************************
        '*  Runge-Kutta function for the position of the X pyrolysis front
        '*
        '*  Called by: Spread_Derivs
        '*
        '*  Revised 20/6/97 Colleen Wade
        '*****************************************************************
        runaway = False
        If LowerWallTemp(room, stepcount) >= WallTSMin(room) Then

            If PessimiseCombWall = True Then
                If (IgTempW(room) - LowerWallTemp(room, stepcount)) > 0 Then
                    RK_XPfront2 = WallFlameSpreadParameter(room) / (ThermalInertiaWall(room) * (IgTempW(room) - LowerWallTemp(room, stepcount)) ^ 2)
                Else
                    'everything ignites, runaway
                    RK_XPfront2 = 0.1 'let limit it to 10 m/s
                    runaway = True
                End If
            Else
                ' If layerheight(room, stepcount) > wallVFSlimit Then
                If (IgTempW(room) - LowerWallTemp(room, stepcount)) > 0 Then
                    RK_XPfront2 = WallFlameSpreadParameter(room) / (ThermalInertiaWall(room) * (IgTempW(room) - LowerWallTemp(room, stepcount)) ^ 2)
                Else
                    'everything ignites, runaway
                    RK_XPfront2 = 0.1 'let limit it to 10 m/s
                    runaway = True
                End If
                ' Else
                'If (IgTempW(room) - Upperwalltemp(room, stepcount)) > 0 Then
                '    RK_XPfront2 = WallFlameSpreadParameter(room) / (ThermalInertiaWall(room) * (IgTempW(room) - Upperwalltemp(room, stepcount)) ^ 2)
                'Else
                '    'everything ignites, runaway
                '    RK_XPfront = 10 'let limit it to 10 m/s
                '    runaway = True
                'End If
                ' End If

            End If

        Else
            RK_XPfront2 = 0
        End If

    End Function
    Sub RK4_Spread(ByVal flux_wall_AR As Double, ByVal fheight As Double, ByVal room As Integer, ByRef Y() As Double, ByRef DYDX() As Double, ByVal n As Integer, ByRef X As Double, ByVal h As Double, ByRef Yout() As Double)
        '///////////////////////////////////////////////////////////////////////
        '   Rk4 Advances A Solution Vector Y(N) Of A Set Of Ordinary Diff.
        '   Eqs. Over A Single Small Interval H Using Fourth-order Runge-kutta
        '   Method.
        '   Revised: 4 October 1996 Colleen Wade
        '//////////////////////////////////////////////////////////////////////

        Dim H6, Hh, XH As Double
        Dim i As Integer
        Dim Yt() As Double
        Dim Dyt() As Double
        Dim Dym() As Double
        ReDim Yt(n)
        ReDim Dyt(n)
        ReDim Dym(n)

        Hh = h * 0.5
        H6 = h / 6.0#
        XH = X + Hh

        For i = 1 To n
            Yt(i) = Y(i) + Hh * DYDX(i)
        Next i

        Call Spread_Derivs_fireroom(flux_wall_AR, fheight, room, XH, Yt, Dyt, n)

        For i = 1 To n
            Yt(i) = Y(i) + Hh * Dyt(i)
        Next i

        Call Spread_Derivs_fireroom(flux_wall_AR, fheight, room, XH, Yt, Dym, n)

        For i = 1 To n
            Yt(i) = Y(i) + h * Dym(i)
            Dym(i) = Dyt(i) + Dym(i)
        Next i

        Call Spread_Derivs_fireroom(flux_wall_AR, fheight, room, X + h, Yt, Dyt, n)

        For i = 1 To n
            Yout(i) = Y(i) + H6 * (DYDX(i) + Dyt(i) + 2.0# * Dym(i))
        Next i

    End Sub
	
	Sub Rk5_Spread(ByVal flux_wall_AR As Double, ByVal fheight As Double, ByVal room As Integer, ByRef Y() As Double, ByRef DYDX() As Double, ByVal n As Integer, ByRef X As Double, ByRef Htry As Double, ByVal EPS As Double, ByRef Yscal() As Double, ByRef Hdid As Double, ByRef hnext As Double)
		'////////////////////////////////////////////////////////////////////////
		'   Rk5 Performs A Single Step Of Fifth-order Runge-kutta Integration
		'   With Local Truncation Error Estimate And Corresponding Step
		'   Size Adjustment. Routine Adapted From Numerical Recipes By
		'   William H. Price Et Al.
		'   From Promath 2.0.
		'
		'*  Calls RK4_Spread
		'*  Revised by Colleen Wade 4 October 1996
		'///////////////////////////////////////////////////////////////////////
		
		Dim Errcon, Fcor, Safety, PGrow As Double
		Dim h, Pshrnk, Xsaved, Hh As Double
		Dim i As Integer
		Dim errmax As Double
		
		Fcor = 0.0666666667
		Safety = 0.9
		Errcon = 0.0006
		
		Dim Ytemp() As Double
		Dim YSAV() As Double
		Dim DYSAV() As Double
		ReDim Ytemp(n)
		ReDim YSAV(n)
		ReDim DYSAV(n)
		
		PGrow = -0.2
		Pshrnk = -0.25
		Xsaved = X
		
		For i = 1 To n ' Save Initial Values
			YSAV(i) = Y(i)
			DYSAV(i) = DYDX(i)
		Next i
		
		h = Htry ' Set Step Size To Initial Trial Size
1: Hh = 0.5 * h ' Take Two Half Step Sizes
		
		Call RK4_Spread(flux_wall_AR, fheight, room, YSAV, DYSAV, n, Xsaved, Hh, Ytemp)
		
		X = Xsaved + Hh
		
		Call Spread_Derivs_fireroom(flux_wall_AR, fheight, room, X, Ytemp, DYDX, n)
		
		Call RK4_Spread(flux_wall_AR, fheight, room, Ytemp, DYDX, n, X, Hh, Y)
		
		X = Xsaved + h
		
		If X = Xsaved Then
			MsgBox("Stepsize Not Significant In Rkqc.")
			flagstop = 1
			Exit Sub
		End If
		
		Call RK4_Spread(flux_wall_AR, fheight, room, YSAV, DYSAV, n, Xsaved, h, Ytemp)
		
		errmax = 0#
		
		For i = 1 To n
			Ytemp(i) = Y(i) - Ytemp(i)
            If Abs(Ytemp(i) / Yscal(i)) > errmax Then
                errmax = Abs(Ytemp(i) / Yscal(i))
            End If
		Next i
		
		errmax = errmax / EPS
		
		If errmax > 1# Then
			h = Safety * h * errmax ^ Pshrnk
			GoTo 1
		Else
			Hdid = h
			If errmax > Errcon Then
				hnext = Safety * h * errmax ^ PGrow
			Else
				hnext = 4# * h
			End If
		End If
		
		For i = 1 To n
			Y(i) = Y(i) + Ytemp(i) * Fcor
		Next i
		
		
	End Sub
	
	Private Sub Spread_Derivs_fireroom(ByVal flux_wall_AR As Double, ByVal fheight As Double, ByVal room As Integer, ByRef X As Double, ByRef YZ() As Double, ByRef Dyz() As Double, ByVal n As Integer)
		'*********************************************************************
		'*  This procedure calculates the ODE functions
		'*  For use with Quintiere's room corner model without burnout
		'*
		'*  Called by: Derk_Spread_fireroom
		'*  Revised 19/6/97 Colleen Wade
		'*********************************************************************
		'Dim room As Integer
        Dim runaway As Boolean
        Dim MaxLength As Double
		Dim floorflux_f, NetFlux_F, NetFLux_B, floorflux_b As Single
		Dim QP_wall_B, QP_floor, QP_wall, QP_ceiling, QP_floor_B As Double

        If frmQuintiere.optUseOneCurve.Checked = True Then
            Call NewPeak(flux_wall_AR, room, QP_wall, QP_ceiling, QP_wall_B, QP_floor, QP_floor_B)
        Else
            Call Get_NetFlux(flux_wall_AR, room, NetFlux_F, NetFLux_B, floorflux_f, floorflux_b)

            'FloorFlux = Target(fireroom, stepcount)
            Call HRR_peak(flux_wall_AR, room, floorflux_f, NetFLux_B, NetFlux_F, QP_wall, QP_ceiling, QP_wall_B, QP_floor)
        End If
		
		If room = fireroom Then
			If WallIgniteFlag(fireroom) = 0 Then
				QP_wall_B = 0
				QP_wall = 0
            End If

			'================================================
			'solve for the position of the Y pyrolysis front
			'================================================

            If FireLocation(burner_id) = 0 Then
                'fire in centre of room
                If CeilingIgniteFlag(room) = 1 Then
                    If YZ(1) >= DMax1(RoomHeight(room) + 0.5 * RoomLength(room), RoomHeight(room) + 0.5 * RoomWidth(room)) Then
                        'the pyrolysis front is now limited by the room size
                        'and no further spread is permitted
                        Dyz(1) = 0
                    Else
                        'does rk_ypfront need to be changed if fire not in corner?
                        Dyz(1) = RK_YPfront2(burner_id, fheight, room, YZ(1), YZ(3), QP_wall, QP_ceiling)
                        If Dyz(1) < 0 Then Dyz(1) = 0
                    End If
                Else
                    Dyz(1) = 0
                End If

                If FloorIgniteFlag(room) = 1 Then
                    If frmQuintiere.optWindAided.Checked = True Then
                        If YZ(4) > 0.5 * RoomLength(room) Then
                            Dyz(4) = 0 'no further windaided floor spread permitted
                        Else
                            Dyz(4) = Floor_YPfront2(fheight, room, YZ(4), YZ(5), QP_wall, QP_floor)
                            If Dyz(4) < 0 Then Dyz(1) = 0
                        End If
                    End If
                Else
                    Dyz(4) = 0
                End If

                '17/11/2011
                If WallIgniteFlag(room) = 1 Then
                    If YZ(1) >= RoomHeight(room) + RoomLength(room) + RoomWidth(room) Then
                        'the pyrolysis front is now limited by the room size
                        'and no further spread is permitted
                        Dyz(1) = 0
                    Else
                        Dyz(1) = RK_YPfront2(burner_id, fheight, room, YZ(1), YZ(3), QP_wall, QP_ceiling)
                        If Dyz(1) < 0 Then Dyz(1) = 0
                    End If
                End If
                '===========================
            Else
                'fire not in centre of room
                If WallIgniteFlag(room) = 1 Or CeilingIgniteFlag(room) = 1 Then
                    If YZ(1) >= RoomHeight(room) + RoomLength(room) + RoomWidth(room) Then
                        'the pyrolysis front is now limited by the room size
                        'and no further spread is permitted
                        Dyz(1) = 0
                    Else
                        Dyz(1) = RK_YPfront2(burner_id, fheight, room, YZ(1), YZ(3), QP_wall, QP_ceiling)
                        If Dyz(1) < 0 Then Dyz(1) = 0
                    End If
                End If

                If FloorIgniteFlag(room) = 1 Then
                    If YZ(4) > RoomLength(room) Then
                        Dyz(4) = 0 'no further windaided floor spread permitted
                    Else
                        Dyz(4) = Floor_YPfront2(fheight, room, YZ(4), YZ(5), QP_wall, QP_floor)
                        If Dyz(4) < 0 Then Dyz(4) = 0
                    End If
                Else
                    Dyz(4) = 0
                End If
            End If

            '================================================
            'solve for the position of the Y burnout front
            '================================================
            If FireLocation(burner_id) = 0 Then
                'fire in centre of room
                If CeilingIgniteFlag(room) = 1 Then
                    Dyz(3) = RK_YBfront(fheight, room, YZ(1), YZ(3), QP_wall, QP_ceiling)
                Else
                    Dyz(3) = 0
                End If

                If FloorIgniteFlag(room) = 1 Then
                    If frmQuintiere.optWindAided.Checked = True Then
                        If YZ(5) > RoomLength(room) / 2 Then
                            Dyz(5) = 0 'no further windaided floor spread permitted
                        Else
                            Dyz(5) = Floor_YBfront(room, YZ(4), YZ(5), QP_floor)
                            If Dyz(5) < 0 Then Dyz(5) = 0
                        End If
                    End If
                Else
                    Dyz(5) = 0
                End If

                '17/11/2011
                If WallIgniteFlag(room) = 1 Then
                    If YZ(3) >= RoomHeight(room) + RoomLength(room) + RoomWidth(room) Then
                        'the pyrolysis front is now limited by the room size
                        'and no further spread is permitted
                        Dyz(3) = 0
                    Else
                        Dyz(3) = RK_YBfront(fheight, room, YZ(1), YZ(3), QP_wall, QP_ceiling)
                        If Dyz(3) < 0 Then Dyz(3) = 0
                    End If
                End If
            Else
                'fire not in centre of room
                If WallIgniteFlag(room) = 1 Or CeilingIgniteFlag(room) = 1 Then
                    If YZ(3) >= RoomHeight(room) + RoomLength(room) + RoomWidth(room) Then
                        'the pyrolysis front is now limited by the room size
                        'and no further spread is permitted
                        Dyz(3) = 0
                    Else
                        Dyz(3) = RK_YBfront(fheight, room, YZ(1), YZ(3), QP_wall, QP_ceiling)
                        If Dyz(3) < 0 Then Dyz(3) = 0
                    End If
                End If

                If FloorIgniteFlag(room) = 1 Then
                    If YZ(5) > RoomLength(room) Then
                        Dyz(5) = 0 'no further windaided floor spread permitted
                    Else
                        Dyz(5) = Floor_YBfront(room, YZ(4), YZ(5), QP_floor)
                        If Dyz(5) < 0 Then Dyz(5) = 0
                    End If
                Else
                    Dyz(5) = 0
                End If
            End If

            '==================================================
            'solve for the position of the X pyrolysis front
            '==================================================
            If frmQuintiere.chkDisableLateralSpread.CheckState = UNCHECKED Then
                If WallIgniteFlag(room) = 1 Or CeilingIgniteFlag(room) = 1 Then
                    If YZ(2) >= RoomLength(room) + RoomWidth(room) Then
                        'the pyrolysis front is now limited by the room size
                        'and no further spread is permitted
                        Dyz(2) = 0
                        YZ(2) = RoomLength(room) + RoomWidth(room)
                    Else
                        Dyz(2) = RK_XPfront(room, runaway)
                        If runaway = True Then
                            Dyz(2) = 0
                            'YZ(2) = RoomLength(room) + RoomWidth(room) '26/4/16 commented out
                        End If
                    End If
                End If
                If FloorIgniteFlag(room) = 1 Then
                    If FireLocation(burner_id) = 0 Then
                        'fire source in centre of room
                        If YZ(6) >= 0.5 * RoomWidth(room) And YZ(6) >= 0.5 * RoomLength(room) Then
                            'the pyrolysis front is now limited by the room size
                            'and no further spread is permitted
                            Dyz(6) = 0
                        Else
                            Dyz(6) = Floor_XPfront(room)
                            If Dyz(6) < 0 Then Dyz(6) = 0
                        End If
                    ElseIf FireLocation(burner_id) = 1 Then
                        'fire against wall
                        If YZ(6) >= RoomWidth(room) And YZ(6) >= RoomLength(room) Then
                            'the pyrolysis front is now limited by the room size
                            'and no further spread is permitted
                            Dyz(6) = 0
                        Else
                            Dyz(6) = Floor_XPfront(room)
                            If Dyz(6) < 0 Then Dyz(6) = 0
                        End If
                    Else
                        'fire in corner
                        If YZ(6) >= Sqrt(RoomWidth(room) ^ 2 + RoomLength(room) ^ 2) Then
                            'the pyrolysis front is now limited by the room size
                            'and no further spread is permitted
                            Dyz(6) = 0
                            YZ(6) = Sqrt(RoomWidth(room) ^ 2 + RoomLength(room) ^ 2)
                        Else
                            Dyz(6) = Floor_XPfront(room)
                            If Dyz(6) < 0 Then Dyz(6) = 0
                        End If
                    End If
                End If
            End If
        Else
            '============================
            'spread to adjacent rooms
            '============================
            If WallIgniteFlag(room) = 1 Or CeilingIgniteFlag(room) = 1 Then
                'solve for the position of the Y pyrolysis front on walls+ceiling
                If RoomWidth(room) = WallLength2(fireroom, room, 1) Then
                    MaxLength = RoomLength(room) + RoomHeight(room)
                Else
                    MaxLength = RoomWidth(room) + RoomHeight(room)
                End If
                If YZ(1) >= MaxLength Then
                    'the pyrolysis front is now limited by the room size
                    'and no further spread is permitted
                    Dyz(1) = 0
                Else
                    Dyz(1) = RK_YPfront2(burner_id, fheight, room, YZ(1), YZ(3), QP_wall, QP_ceiling)
                    If Dyz(1) < 0 Then Dyz(1) = 0
                End If

                'solve for the position of the X pyrolysis front on walls
                If YZ(2) >= RoomLength(room) + RoomWidth(room) Then
                    'the pyrolysis front is now limited by the height of the walls
                    'and no further spread is permitted
                    Dyz(2) = 0
                    YZ(2) = RoomLength(room) + RoomWidth(room)
                Else
                    Dyz(2) = RK_XPfront(room, runaway)
                    If runaway = True Then
                        Dyz(2) = 0
                        YZ(2) = RoomLength(room) + RoomWidth(room)
                    End If

                    If Dyz(2) < 0 Then Dyz(2) = 0
                End If
            End If

            If FloorIgniteFlag(room) = 1 Then
                'solve for the position of the Y pyrolysis front on floor
                If RoomWidth(room) = WallLength2(fireroom, room, 1) Then
                    MaxLength = RoomLength(room)
                Else
                    MaxLength = RoomWidth(room)
                End If
                If YZ(4) >= MaxLength Then
                    'the pyrolysis front is now limited by the room size
                    'and no further spread is permitted
                    Dyz(4) = 0
                Else
                    Dyz(4) = Floor_YPfront2(fheight, room, YZ(4), YZ(5), QP_wall, QP_floor)
                    If Dyz(4) < 0 Then Dyz(4) = 0
                End If

                'burnout
                If YZ(5) >= MaxLength Then
                    Dyz(5) = 0 'no further windaided floor spread permitted
                Else
                    Dyz(5) = Floor_YBfront(room, YZ(4), YZ(5), QP_floor)
                    If Dyz(5) < 0 Then Dyz(5) = 0
                End If

                'lateral spread
                If YZ(6) >= 0.5 * RoomWidth(room) Then
                    'the pyrolysis front is now limited by the room size
                    'and no further spread is permitted
                    Dyz(6) = 0
                    YZ(6) = 0.5 * RoomWidth(room)
                Else
                    Dyz(6) = Floor_XPfront(room)
                    If Dyz(6) < 0 Then Dyz(6) = 0
                End If
            End If
        End If
		
	End Sub
	
    Sub Wall_Ignite(ByVal i As Short, ByRef IgWallNode(,,) As Double)
        '*  ================================================================
        '*  This function updates the surface temperatures on the wall next to
        '*  the burner, using an implicit finite difference method.
        '*  For use with Quintiere's room corner model.
        '*
        '*  IgWallNode = an array storing the temperatures within the wall adjacent to the burner
        '*
        '*  Colleen Wade 24/7/98
        '*  ================================================================

        Dim k, j As Integer
        Dim Qnet As Double
        Dim UW(Wallnodes, Wallnodes) As Double
        Dim wx(Wallnodes, 1) As Double

        UW(1, 1) = 1 + 2 * WallFourier(fireroom)
        UW(1, 2) = -2 * WallFourier(fireroom)

        k = 2
        For j = 2 To Wallnodes - 1
            UW(j, k - 1) = -WallFourier(fireroom)
            UW(j, k) = 1 + 2 * WallFourier(fireroom)
            UW(j, k + 1) = -WallFourier(fireroom)
            k = k + 1
        Next j

        UW(Wallnodes, Wallnodes - 1) = -2 * WallFourier(fireroom)
        UW(Wallnodes, Wallnodes) = 1 + 2 * WallFourier(fireroom)

        'net heat flux into the wall adjacent to the burner flame
        'Qnet = -QFB + QLowerWall(fireroom, i) - StefanBoltzmann / 1000 * Surface_Emissivity(3, fireroom) * (LowerWallTemp(fireroom, i) ^ 4 - (IgWallNode(fireroom, 1, i)) ^ 4) 'kW/m2

        'switched to use upper wall 9June2003
        Qnet = -QFB + QUpperWall(fireroom, i) - StefanBoltzmann / 1000 * Surface_Emissivity(2, fireroom) * (Upperwalltemp(fireroom, i) ^ 4 - (IgWallNode(fireroom, 1, i)) ^ 4) 'kW/m2
        'Qnet = -QFB + QLowerWall(i) + stefanboltzmann / 1000 * Surface_Emissivity(3) * (IgWallNode(1, i) ^ 4)'kW/m2

        wx(1, 1) = -2 * Qnet * 1000 * WallFourier(fireroom) * WallDeltaX(fireroom) / WallConductivity(fireroom) + IgWallNode(fireroom, 1, i)

        For k = 2 To Wallnodes - 1
            wx(k, 1) = IgWallNode(fireroom, k, i)
        Next k

        'insulated rear face
        'WX(WallNodes, 1) = IgWallNode(WallNodes, i)
        'exterior boundary conditions
        wx(Wallnodes, 1) = 2 * WallFourier(fireroom) * WallOutsideBiot(fireroom) * (ExteriorTemp - IgWallNode(fireroom, Wallnodes, i)) + IgWallNode(fireroom, Wallnodes, i)

        Dim ier As Short
        If frmOptions1.optLUdecom.Checked = True Then
            'find surface temperatures for the next timestep
            'using method of LU decomposition
            Call MatSol(UW, wx, Wallnodes) 'wall
        Else
            'find surface temperatures for the next timestep
            'using method of Gauss-Jordan elimination
            Call LINEAR2(Wallnodes, UW, wx, ier)
            If ier = 1 Then MsgBox("singular matrix in implicit_surface_temps")
        End If

        For j = 1 To Wallnodes
            IgWallNode(fireroom, j, i + 1) = wx(j, 1)
        Next j

        'store surface temps in another array
        IgWallTemp(fireroom, i + 1) = IgWallNode(fireroom, 1, i + 1)
        'Debug.Print tim(i); IgWallNode(1, i) - 273, IgWallNode(WallNodes, i) - 273

    End Sub
	
    Sub Wall_Ignite2(ByVal i As Short, ByRef IgWallNode(,,) As Double, ByRef IncidentFlux As Double)
        '*  ================================================================
        '*  This function updates the surface temperatures on the wall next to
        '*  the burner, using an implicit finite difference method.
        '*  For use with Quintiere's room corner model.
        '*
        '*  IgWallNode = an array storing the temperatures within the wall adjacent to the burner
        '*
        '*  Colleen Wade 24/7/98
        '*  ================================================================

        Dim k, j As Integer
        Dim WallFourier2, Qnet, WallOutsideBiot2 As Double
        Dim UW(2 * Wallnodes - 1, 2 * Wallnodes - 1) As Double
        Dim wx(2 * Wallnodes - 1, 1) As Double

        WallFourier2 = WallSubConductivity(fireroom) * Timestep / (((WallSubThickness(fireroom) / 1000) / (Wallnodes - 1)) ^ 2 * WallSubDensity(fireroom) * WallSubSpecificHeat(fireroom))
        WallOutsideBiot2 = OutsideConvCoeff * ((WallSubThickness(fireroom) / 1000) / (Wallnodes - 1)) / WallSubConductivity(fireroom)

        UW(1, 1) = 1 + 2 * WallFourier(fireroom)
        UW(1, 2) = -2 * WallFourier(fireroom)

        k = 2
        For j = 2 To Wallnodes - 1
            UW(j, k - 1) = -WallFourier(fireroom)
            UW(j, k) = 1 + 2 * WallFourier(fireroom)
            UW(j, k + 1) = -WallFourier(fireroom)
            k = k + 1
        Next j

        j = Wallnodes
        UW(j, k - 1) = -WallFourier(fireroom)
        UW(j, k) = 1 + WallFourier(fireroom) + WallFourier2
        UW(j, k + 1) = -WallFourier2
        k = k + 1

        For j = Wallnodes + 1 To 2 * Wallnodes - 2
            UW(j, k - 1) = -WallFourier2
            UW(j, k) = 1 + 2 * WallFourier2
            UW(j, k + 1) = -WallFourier2
            k = k + 1
        Next j

        UW(2 * Wallnodes - 1, 2 * Wallnodes - 2) = -2 * WallFourier2
        UW(2 * Wallnodes - 1, 2 * Wallnodes - 1) = 1 + 2 * WallFourier2

        'net heat flux into the wall adjacent to the burner flame
        'Qnet = -QFB + QLowerWall(fireroom, i) - StefanBoltzmann / 1000 * Surface_Emissivity(3, fireroom) * (LowerWallTemp(fireroom, i) ^ 4 - IgWallNode(fireroom, 1, i) ^ 4) 'kW/m2
        'Qnet = -QFB + QLowerWall(fireroom, i) - StefanBoltzmann / 1000 * Surface_Emissivity(3, fireroom) * (LowerWallTemp(fireroom, i) ^ 4 - (IgWallNode(fireroom, 1, i)) ^ 4)  'kW/m2
        'Qnet = -QFB + QLowerWall(i) + stefanboltzmann / 1000 * Surface_Emissivity(3) * (IgWallNode(1, i) ^ 4)'kW/m2
        'IncidentFlux = QFB - QLowerWall(fireroom, i)

        Qnet = -QFB + QUpperWall(fireroom, i) - StefanBoltzmann / 1000 * Surface_Emissivity(2, fireroom) * (Upperwalltemp(fireroom, i) ^ 4 - IgWallNode(fireroom, 1, i) ^ 4) 'kW/m2
        IncidentFlux = QFB - QUpperWall(fireroom, i)

        wx(1, 1) = -2 * Qnet * 1000 * WallFourier(fireroom) * WallDeltaX(fireroom) / WallConductivity(fireroom) + IgWallNode(fireroom, 1, i)

        For k = 2 To 2 * Wallnodes - 2
            wx(k, 1) = IgWallNode(fireroom, k, i)
        Next k

        'insulated rear face
        'WX(2 * WallNodes - 1, 1) = IgWallNode(2 * WallNodes - 1, i)
        'exterior boundary conditions
        wx(2 * Wallnodes - 1, 1) = 2 * WallFourier2 * WallOutsideBiot2 * (ExteriorTemp - IgWallNode(fireroom, 2 * Wallnodes - 1, i)) + IgWallNode(fireroom, 2 * Wallnodes - 1, i)

        Dim ier As Short
        If frmOptions1.optLUdecom.Checked = True Then
            'find surface temperatures for the next timestep
            'using method of LU decomposition
            Call MatSol(UW, wx, 2 * Wallnodes - 1) 'wall
        Else
            'find surface temperatures for the next timestep
            'using method of Gauss-Jordan elimination
            Call LINEAR2(2 * Wallnodes - 1, UW, wx, ier)
            If ier = 1 Then MsgBox("singular matrix in implicit_surface_temps")
        End If

        For j = 1 To 2 * Wallnodes - 1
            IgWallNode(fireroom, j, i + 1) = wx(j, 1)
        Next j

        'store surface temps in another array
        IgWallTemp(fireroom, i + 1) = IgWallNode(fireroom, 1, i + 1)
        'Debug.Print tim(i, 1); IgWallTemp(fireroom, i + 1) - 273, Qnet

    End Sub
	
	
	Public Sub Integrate_New_HRR(ByVal room As Integer, ByVal id As Integer, ByVal low As Single, ByVal high As Single, ByVal TimeFactor As Double, ByRef IO As Double, ByVal Q As Double)
		'************************************************************
		'*  A procedure to find the area under the cone heat release
		'*  rate curve
		'*
		'*  Created 11 November 1997 Colleen Wade
		'************************************************************
        Dim Tempx(1) As Double
        Dim TempY(1) As Double
		Dim i, n As Integer
		Dim sum As Double
		Dim ErrorCode As Integer
		Dim yinterp As Double
		Dim X() As Double
		Dim W() As Double
		n = 50
		
		On Error GoTo errorhandler
		
		ReDim X(n + 500)
		ReDim W(n + 500)
		If id = 1 Then
            ReDim Tempx(ConeNumber_W(1, room))
            ReDim TempY(ConeNumber_W(1, room))
		ElseIf id = 2 Then 
			ReDim Tempx(ConeNumber_C(1, room))
			ReDim TempY(ConeNumber_C(1, room))
		ElseIf id = 3 Then 
			ReDim Tempx(ConeNumber_F(1, room))
			ReDim TempY(ConeNumber_F(1, room))
		End If
        Call GAULEG2(low, TimeFactor * high, X, W, n, ErrorCode)
		
		If ErrorCode <> 0 Then
			MsgBox("Internal error in routine GAULEG2 in Integrate_New_HRR")
			Exit Sub
		End If
		
		sum = 0
		
		If id = 1 Then
			For i = 1 To ConeNumber_W(1, room)
				Tempx(i) = TimeFactor * ConeYW(room, 1, i) 'time
				TempY(i) = Q * NormalHRR_W(1, i) 'hrr
			Next i
			For i = 1 To n
				Call Interpolate_D(Tempx, TempY, ConeNumber_W(1, room), X(i), yinterp)
				sum = sum + W(i) * yinterp
			Next i
		ElseIf id = 2 Then 
			For i = 1 To ConeNumber_C(1, room)
				Tempx(i) = TimeFactor * ConeYC(room, 1, i) 'time
				TempY(i) = Q * NormalHRR_C(1, i) 'hrr
			Next i
			
			For i = 1 To n
				Call Interpolate_D(Tempx, TempY, ConeNumber_C(1, room), X(i), yinterp)
				sum = sum + W(i) * yinterp
			Next i
		ElseIf id = 3 Then 
			For i = 1 To ConeNumber_F(1, room)
				Tempx(i) = TimeFactor * ConeYF(room, 1, i) 'time
				TempY(i) = Q * NormalHRR_F(1, i) 'hrr
			Next i
			For i = 1 To n
				Call Interpolate_D(Tempx, TempY, ConeNumber_F(1, room), X(i), yinterp)
				sum = sum + W(i) * yinterp
			Next i
		End If
		IO = sum
		Exit Sub
		
errorhandler: 
		MsgBox(ErrorToString(Err.Number) & " in subroutine Integrate_New_Hrr, Module QUINT")
		Dim response, Help, Style, msg, Title, Ctxt, MyString As Object
        msg = "Do you want to continue ?" ' Define message.
        Style = MsgBoxStyle.YesNo + MsgBoxStyle.Critical + MsgBoxStyle.DefaultButton2 ' Define buttons.
				response = MsgBox(msg, Style, Title)
		If response = MsgBoxResult.Yes Then ' User chose Yes.
			'MyString = "Yes"   ' Perform some action.
		Else ' User chose No.
			flagstop = 1
		End If
		
	End Sub
	
	Public Sub Get_Normal_Data2(ByVal room As Integer, ByVal id As Integer, ByVal tim As Double, ByRef NormalQ As Double, ByRef TimeFactor As Double)
		'********************************************************
		'*  Get normalised cone calorimeter data from data file at time=tim.
		'*
		'*  Created 24 April 1997 Colleen Wade
		'********************************************************
        Dim TempY() As Double
        Dim Tempx() As Double
		Dim i As Integer
		System.Windows.Forms.Application.DoEvents()
		'return hrr at time tim as Q
		'NormalHRR_W() is normalised HRR
		'ConeYW() is time
		If id = 1 Then
            ReDim TempY(ConeNumber_W(1, room))
            ReDim Tempx(ConeNumber_W(1, room))
			For i = 1 To ConeNumber_W(1, room)
				TempY(i) = TimeFactor * ConeYW(room, 1, i) 'time
				Tempx(i) = NormalHRR_W(1, i) 'hrr
			Next i
			Call Interpolate_D(TempY, Tempx, ConeNumber_W(1, room), tim, NormalQ)
			Exit Sub
		ElseIf id = 2 Then 
			ReDim TempY(ConeNumber_C(1, room))
			ReDim Tempx(ConeNumber_C(1, room))
			For i = 1 To ConeNumber_C(1, room)
				TempY(i) = TimeFactor * ConeYC(room, 1, i) 'time
				Tempx(i) = NormalHRR_C(1, i) 'hrr
			Next i
			Call Interpolate_D(TempY, Tempx, ConeNumber_C(1, room), tim, NormalQ)
			Exit Sub
		ElseIf id = 3 Then 
			ReDim TempY(ConeNumber_F(1, room))
			ReDim Tempx(ConeNumber_F(1, room))
			For i = 1 To ConeNumber_F(1, room)
				TempY(i) = TimeFactor * ConeYF(room, 1, i) 'time
				Tempx(i) = NormalHRR_F(1, i) 'hrr
			Next i
			Call Interpolate_D(TempY, Tempx, ConeNumber_F(1, room), tim, NormalQ)
			Exit Sub
		End If
	End Sub
	
	Public Sub NewPeak(ByVal flux_wall_AR As Double, ByVal room As Integer, ByRef QP_wall As Double, ByRef QP_ceiling As Double, ByRef QP_wall_B As Double, ByRef QP_floor As Double, ByRef QP_floor_B As Double)
		'*********************************************************
		'*  This function calculates the heat release rate per unit area of the wall material.
		'*  For use in Quintiere's room corner model.
		'*
		'*  L = heat of gasification of the material (kJ/g)
		'*
		'*  5 April 1998 Colleen Wade
		'********************************************************
		
		Dim NetFlux_F, NetFLux_B As Single
		'function returns the heat release rate per unit area
		'wall is burning
		
		If WallIgniteFlag(room) = 1 Then
			If WallHeatofGasification(room) > 0 And stepcount > 1 Then
				'If room = fireroom Then
				'If IgTempW(room) + Delta_TV > Upperwalltemp((room), stepcount) Then
				'    max = IgTempW(room) + Delta_TV
				'Else
				'    max = Upperwalltemp((room), stepcount)
				'End If
				
				'NetFlux_F = (QPyrol - QUpperWall((room), stepcount) + StefanBoltzmann / 1000 * Surface_Emissivity(3, (room)) * (Upperwalltemp((room), stepcount) ^ 4 - max ^ 4))
				'NetFlux_B = (QFB - QUpperWall((room), stepcount) + StefanBoltzmann / 1000 * Surface_Emissivity(3, (room)) * (Upperwalltemp((room), stepcount) ^ 4 - max ^ 4))
				'without reradiation term
				NetFlux_F = (QPyrol - QUpperWall(room, stepcount) + StefanBoltzmann / 1000 * Surface_Emissivity(2, room) * (Upperwalltemp(room, stepcount) ^ 4))
				If room = fireroom Then
					NetFLux_B = (QFB - QUpperWall(room, stepcount) + StefanBoltzmann / 1000 * Surface_Emissivity(2, room) * (Upperwalltemp(room, stepcount) ^ 4))
				Else
					NetFLux_B = (flux_wall_AR - QUpperWall(room, stepcount) + StefanBoltzmann / 1000 * Surface_Emissivity(2, room) * (Upperwalltemp(room, stepcount) ^ 4))
				End If
				
				If WallHeatofGasification(room) > 0 And stepcount > 1 Then
					QP_wall = WallEffectiveHeatofCombustion(room) / WallHeatofGasification(room) * NetFlux_F 'kW/m2
					QP_wall_B = WallEffectiveHeatofCombustion(room) / WallHeatofGasification(room) * NetFLux_B 'kW/m2
				Else
					QP_wall = 0
					QP_wall_B = 0
				End If
				'Else
				'QP_wall = 0
				'QP_wall_B = 0
				'End If
				'QP_wall = 0
				' QP_wall_B = 0
			Else
				
			End If
		End If
		
		If CeilingIgniteFlag(room) = 1 Then
			If CeilingHeatofGasification(room) > 0 And stepcount > 1 Then
				'ceiling is burning
				QP_ceiling = CeilingEffectiveHeatofCombustion(room) / CeilingHeatofGasification(room) * (QPyrol - QCeiling(room, stepcount) + StefanBoltzmann / 1000 * Surface_Emissivity(1, room) * (CeilingTemp(room, stepcount) ^ 4)) 'kW/m2
			Else
				QP_ceiling = 0
			End If
		End If
		
		If FloorIgniteFlag(room) = 1 Then
			If FloorHeatofGasification(room) > 0 And stepcount > 1 Then
				QP_floor = FloorEffectiveHeatofCombustion(room) / FloorHeatofGasification(room) * (QPyrol - QFloor(room, stepcount) + StefanBoltzmann / 1000 * Surface_Emissivity(4, room) * (FloorTemp(room, stepcount) ^ 4)) 'kW/m2
				QP_floor_B = FloorEffectiveHeatofCombustion(room) / FloorHeatofGasification(room) * (QFB - QFloor(room, stepcount) + StefanBoltzmann / 1000 * Surface_Emissivity(4, room) * (FloorTemp(room, stepcount) ^ 4)) 'kW/m2
			Else
				QP_floor = 0
				QP_floor_B = 0
			End If
		End If
	End Sub
	
    Public Sub Ceiling_Ignite(ByVal i As Short, ByRef IgCeilingNode(,,) As Double)
        '*  ================================================================
        '*  This function updates the surface temperatures on the ceiling above the
        '*  the burner, using an implicit finite difference method.
        '*
        '*  IgCeilingNode = an array storing the temperatures within the ceiling.
        '*
        '*  Colleen Wade 24/7/98
        '*  ================================================================

        Dim k, j As Integer
        Dim gastemp, Qnet, h, VF As Double
        Dim UC(Ceilingnodes, Ceilingnodes) As Double
        Dim cx(Ceilingnodes, 1) As Double

        UC(1, 1) = 1 + 2 * CeilingFourier(fireroom)
        UC(1, 2) = -2 * CeilingFourier(fireroom)

        k = 2
        For j = 2 To Ceilingnodes - 1
            UC(j, k - 1) = -CeilingFourier(fireroom)
            UC(j, k) = 1 + 2 * CeilingFourier(fireroom)
            UC(j, k + 1) = -CeilingFourier(fireroom)
            k = k + 1
        Next j

        UC(Ceilingnodes, Ceilingnodes - 1) = -2 * CeilingFourier(fireroom)
        UC(Ceilingnodes, Ceilingnodes) = 1 + 2 * CeilingFourier(fireroom)

        'get ceiling jet temperature
        'gastemp = CeilingJet_MaxTemp(HeatRelease(fireroom, stepcount, 2))
        'UPGRADE_WARNING: Couldn't resolve default property of object JET_CJmax(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        gastemp = JET_CJmax(HeatRelease(fireroom, stepcount, 2), fireroom, (tim(stepcount, 1)), layerheight(fireroom, stepcount), uppertemp(fireroom, stepcount), lowertemp(fireroom, stepcount), 0)
        'convective heat transfer coefficent
        h = Get_Convection_Coefficient2(1, gastemp, IgCeilingTemp(fireroom, stepcount), CStr(HORIZONTAL))

        If (2 / 3) * BurnerFlameLength > RoomHeight(fireroom) - FireHeight(1) Then
            'net heat flux into the ceiling above the burner flame
            'uses the same radiant component that would be used for the wall
            'Qnet = -(3 * RadiantLossFraction * Qburner / (4 * BurnerWidth * BurnerFlameLength) * (4 * BurnerFlameLength / (3 * BurnerWidth + 8 * BurnerFlameLength))) - (h / 1000 * (gastemp - IgCeilingTemp(fireroom, stepcount))) + QCeiling(fireroom, i) - StefanBoltzmann / 1000 * Surface_Emissivity(1, fireroom) * (CeilingTemp(fireroom, i) ^ 4 - IgCeilingNode(fireroom, 1, i) ^ 4) 'kW/m2
            'config factor between flame and ceiling =1
            Qnet = -ObjectRLF(1) * Qburner / (BurnerWidth ^ 2 + 4 * 2 / 3 * BurnerFlameLength * BurnerWidth) - (h / 1000 * (gastemp - IgCeilingTemp(fireroom, stepcount))) + QCeiling(fireroom, i) - StefanBoltzmann / 1000 * Surface_Emissivity(1, fireroom) * (CeilingTemp(fireroom, i) ^ 4 - IgCeilingNode(fireroom, 1, i) ^ 4) 'kW/m2
        Else
            'net heat flux into the ceiling above the burner flame
            'Qnet = -StefanBoltzmann / 1000 * Surface_Emissivity(1) * (gastemp ^ 4 - IgCeilingNode(1, i) ^ 4) - (h / 1000 * (gastemp - IgCeilingTemp(stepcount))) + QCeiling(i) - StefanBoltzmann / 1000 * Surface_Emissivity(1) * (CeilingTemp(i) ^ 4 - IgCeilingNode(1, i) ^ 4)  'kW/m2
            'Qnet = -RadiantLossFraction * Qburner * (BurnerWidth / (BurnerWidth + 8 / 3 * BurnerFlameLength) / BurnerWidth ^ 2) - (h / 1000 * (gastemp - IgCeilingTemp(fireroom, stepcount))) + QCeiling(fireroom, i) - StefanBoltzmann / 1000 * Surface_Emissivity(1, fireroom) * (CeilingTemp(fireroom, i) ^ 4 - IgCeilingNode(fireroom, 1, i) ^ 4) 'kW/m2
            'get config factor
            'UPGRADE_WARNING: Couldn't resolve default property of object ViewFactor_Flame(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VF = ViewFactor_Flame(BurnerWidth / 2, BurnerWidth / 2, RoomHeight(fireroom) - FireHeight(1) - 2 / 3 * BurnerFlameLength)
            Qnet = -4 * VF * ObjectRLF(1) * Qburner / (BurnerWidth ^ 2 + 4 * 2 / 3 * BurnerFlameLength * BurnerWidth) - (h / 1000 * (gastemp - IgCeilingTemp(fireroom, stepcount))) + QCeiling(fireroom, i) - StefanBoltzmann / 1000 * Surface_Emissivity(1, fireroom) * (CeilingTemp(fireroom, i) ^ 4 - IgCeilingNode(fireroom, 1, i) ^ 4) 'kW/m2
        End If

        cx(1, 1) = -2 * Qnet * 1000 * CeilingFourier(fireroom) * CeilingDeltaX(fireroom) / CeilingConductivity(fireroom) + IgCeilingNode(fireroom, 1, i)

        For k = 2 To Ceilingnodes - 1
            cx(k, 1) = IgCeilingNode(fireroom, k, i)
        Next k

        'insulated rear face
        'exterior boundary conditions
        cx(Ceilingnodes, 1) = 2 * CeilingFourier(fireroom) * CeilingOutsideBiot(fireroom) * (ExteriorTemp - IgCeilingNode(fireroom, Ceilingnodes, i)) + IgCeilingNode(fireroom, Ceilingnodes, i)

        'find surface temperatures for the next timestep
        Call MatSol(UC, cx, Ceilingnodes)

        Dim ier As Short
        If frmOptions1.optLUdecom.Checked = True Then
            'find surface temperatures for the next timestep
            'using method of LU decomposition
            Call MatSol(UC, cx, Ceilingnodes)
        Else
            'find surface temperatures for the next timestep
            'using method of Gauss-Jordan elimination
            Call LINEAR2(Ceilingnodes, UC, cx, ier)
            If ier = 1 Then MsgBox("singular matrix in subroutine ceiling_ignite.")
        End If

        For j = 1 To Ceilingnodes
            IgCeilingNode(fireroom, j, i + 1) = cx(j, 1)
        Next j

        'store surface temps in another array
        IgCeilingTemp(fireroom, i + 1) = IgCeilingNode(fireroom, 1, i + 1)

    End Sub
	
    Public Sub Ceiling_Ignite2(ByVal i As Short, ByRef IgCeilingNode(,,) As Double)
        '*  ================================================================
        '*  This function updates the surface temperatures on the ceiling above
        '*  the burner, using an implicit finite difference method.
        '*  For use with Quintiere's room corner model and a two-layer ceiling
        '*
        '*  IgCeilingNode = an array storing the temperatures within the ceiling.
        '*
        '*  Colleen Wade 24/7/98
        '*  ================================================================

        Dim k, j As Integer
        Dim CeilingFourier2, Qnet, CeilingOutsideBiot2 As Double
        Dim gastemp, h, VF As Double
        Dim UC(,) As Double
        Dim cx(,) As Double
        ReDim UC(2 * Ceilingnodes - 1, 2 * Ceilingnodes - 1)
        ReDim cx(2 * Ceilingnodes - 1, 1)

        CeilingFourier2 = CeilingSubConductivity(fireroom) * Timestep / (((CeilingSubThickness(fireroom) / 1000) / (Ceilingnodes - 1)) ^ 2 * CeilingSubDensity(fireroom) * CeilingSubSpecificHeat(fireroom))
        CeilingOutsideBiot2 = OutsideConvCoeff * ((CeilingSubThickness(fireroom) / 1000) / (Ceilingnodes - 1)) / CeilingSubConductivity(fireroom)

        UC(1, 1) = 1 + 2 * CeilingFourier(fireroom)
        UC(1, 2) = -2 * CeilingFourier(fireroom)

        k = 2
        For j = 2 To Ceilingnodes - 1
            UC(j, k - 1) = -CeilingFourier(fireroom)
            UC(j, k) = 1 + 2 * CeilingFourier(fireroom)
            UC(j, k + 1) = -CeilingFourier(fireroom)
            k = k + 1
        Next j

        j = Ceilingnodes
        UC(j, k - 1) = -CeilingFourier(fireroom)
        UC(j, k) = 1 + CeilingFourier(fireroom) + CeilingFourier2
        UC(j, k + 1) = -CeilingFourier2
        k = k + 1

        For j = Ceilingnodes + 1 To 2 * Ceilingnodes - 2
            UC(j, k - 1) = -CeilingFourier2
            UC(j, k) = 1 + 2 * CeilingFourier2
            UC(j, k + 1) = -CeilingFourier2
            k = k + 1
        Next j

        UC(2 * Ceilingnodes - 1, 2 * Ceilingnodes - 2) = -2 * CeilingFourier2
        UC(2 * Ceilingnodes - 1, 2 * Ceilingnodes - 1) = 1 + 2 * CeilingFourier2

        'get ceiling jet temperature
        'gastemp = CeilingJet_MaxTemp(HeatRelease(fireroom, stepcount, 2))
        'UPGRADE_WARNING: Couldn't resolve default property of object JET_CJmax(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        gastemp = JET_CJmax(HeatRelease(fireroom, stepcount, 2), fireroom, tim(i, 1), layerheight(fireroom, i), uppertemp(fireroom, i), lowertemp(fireroom, i), 0)
        'convective heat transfer coefficent
        h = Get_Convection_Coefficient2(1, gastemp, IgCeilingTemp(fireroom, stepcount), CStr(HORIZONTAL))
        'radiant flame flux
        If (2 / 3) * BurnerFlameLength > RoomHeight(fireroom) - FireHeight(1) Then
            'net heat flux into the ceiling above the burner flame
            'uses the same radiant component that would be used for the wall
            'Qnet = -(3 * RadiantLossFraction * Qburner / (4 * BurnerWidth * BurnerFlameLength) * (4 * BurnerFlameLength / (3 * BurnerWidth + 8 * BurnerFlameLength))) - (h / 1000 * (gastemp - IgCeilingTemp(fireroom, stepcount))) + QCeiling(fireroom, i) - StefanBoltzmann / 1000 * Surface_Emissivity(1, fireroom) * (CeilingTemp(fireroom, i) ^ 4 - IgCeilingNode(fireroom, 1, i) ^ 4) 'kW/m2
            'config factor between flame and ceiling =1
            Qnet = -ObjectRLF(1) * Qburner / (BurnerWidth ^ 2 + 4 * 2 / 3 * BurnerFlameLength * BurnerWidth) - (h / 1000 * (gastemp - IgCeilingTemp(fireroom, stepcount))) + QCeiling(fireroom, i) - StefanBoltzmann / 1000 * Surface_Emissivity(1, fireroom) * (CeilingTemp(fireroom, i) ^ 4 - IgCeilingNode(fireroom, 1, i) ^ 4) 'kW/m2
            'Qnet = -RadiantLossFraction * Qburner / (BurnerWidth ^ 2 + 4 * BurnerFlameLength * BurnerWidth) - (h / 1000 * (gastemp - IgCeilingTemp(fireroom, stepcount))) + QCeiling(fireroom, i) - StefanBoltzmann / 1000 * Surface_Emissivity(1, fireroom) * (CeilingTemp(fireroom, i) ^ 4 - IgCeilingNode(fireroom, 1, i) ^ 4) 'kW/m2
        Else
            'net heat flux into the ceiling above the burner flame
            'get config factor
            'UPGRADE_WARNING: Couldn't resolve default property of object ViewFactor_Flame(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VF = ViewFactor_Flame(BurnerWidth / 2, BurnerWidth / 2, RoomHeight(fireroom) - FireHeight(1) - 2 / 3 * BurnerFlameLength)
            Qnet = -4 * VF * ObjectRLF(1) * Qburner / (BurnerWidth ^ 2 + 4 * 2 / 3 * BurnerFlameLength * BurnerWidth) - (h / 1000 * (gastemp - IgCeilingTemp(fireroom, stepcount))) + QCeiling(fireroom, i) - StefanBoltzmann / 1000 * Surface_Emissivity(1, fireroom) * (CeilingTemp(fireroom, i) ^ 4 - IgCeilingNode(fireroom, 1, i) ^ 4) 'kW/m2
            'vf = ViewFactor_Flame(BurnerWidth / 2, BurnerWidth / 2, RoomHeight(fireroom) - FireHeight(1) - BurnerFlameLength)
            'Qnet = -4 * vf * RadiantLossFraction * Qburner / (BurnerWidth ^ 2 + 4 * BurnerFlameLength * BurnerWidth) - (h / 1000 * (gastemp - IgCeilingTemp(fireroom, stepcount))) + QCeiling(fireroom, i) - StefanBoltzmann / 1000 * Surface_Emissivity(1, fireroom) * (CeilingTemp(fireroom, i) ^ 4 - IgCeilingNode(fireroom, 1, i) ^ 4) 'kW/m2
        End If
        'Debug.Print stepcount; Qnet
        cx(1, 1) = -2 * Qnet * 1000 * CeilingFourier(fireroom) * CeilingDeltaX(fireroom) / CeilingConductivity(fireroom) + IgCeilingNode(fireroom, 1, i)

        For k = 2 To 2 * Ceilingnodes - 2
            cx(k, 1) = IgCeilingNode(fireroom, k, i)
        Next k

        'insulated rear face
        'exterior boundary conditions
        cx(2 * Ceilingnodes - 1, 1) = 2 * CeilingFourier2 * CeilingOutsideBiot2 * (ExteriorTemp - IgCeilingNode(fireroom, 2 * Ceilingnodes - 1, i)) + IgCeilingNode(fireroom, 2 * Ceilingnodes - 1, i)

        Dim ier As Short
        If frmOptions1.optLUdecom.Checked = True Then
            'find surface temperatures for the next timestep
            'using method of LU decomposition
            Call MatSol(UC, cx, 2 * Ceilingnodes - 1)
        Else
            'find surface temperatures for the next timestep
            'using method of Gauss-Jordan elimination
            Call LINEAR2(2 * Ceilingnodes - 1, UC, cx, ier)
            If ier = 1 Then MsgBox("singular matrix in subroutine ceiling_ignite2.")
        End If

        For j = 1 To 2 * Ceilingnodes - 1
            IgCeilingNode(fireroom, j, i + 1) = cx(j, 1)
        Next j

        'store surface temps in another array
        IgCeilingTemp(fireroom, i + 1) = IgCeilingNode(fireroom, 1, i + 1)

    End Sub
	
    Public Sub Ceiling_Ignite_Adj1(ByRef projection As Double, ByRef room As Integer, ByVal i As Integer, ByRef IgCeilingNode(,,) As Double, ByRef layer1 As Double, ByRef ventfire As Double, ByRef ceilingheight2 As Double, ByRef vwidth As Single)
        '*  ================================================================
        '*  This function updates the surface temperatures on the ceiling above the
        '*  the burner, using an implicit finite difference method.
        '*
        '*  IgCeilingNode = an array storing the temperatures within the ceiling.
        '*
        '*  Colleen Wade 24/7/98
        '*  ================================================================

        Dim qrad, flameheight As Double
        Dim VF As Single
        Dim k, j As Integer
        Dim h, Qnet, gastemp As Double
        Dim UC(Ceilingnodes, Ceilingnodes) As Double
        Dim cx(Ceilingnodes, 1) As Double

        UC(1, 1) = 1 + 2 * CeilingFourier(room)
        UC(1, 2) = -2 * CeilingFourier(room)

        k = 2
        For j = 2 To Ceilingnodes - 1
            UC(j, k - 1) = -CeilingFourier(room)
            UC(j, k) = 1 + 2 * CeilingFourier(room)
            UC(j, k + 1) = -CeilingFourier(room)
            k = k + 1
        Next j

        UC(Ceilingnodes, Ceilingnodes - 1) = -2 * CeilingFourier(room)
        UC(Ceilingnodes, Ceilingnodes) = 1 + 2 * CeilingFourier(room)

        'get ceiling jet temperature in adjacent room
        'UPGRADE_WARNING: Couldn't resolve default property of object CeilingJet_MaxTemp_Adj(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        gastemp = CeilingJet_MaxTemp_Adj(ventfire, room, layer1)
        'convective heat transfer coefficent
        h = Get_Convection_Coefficient2(1, gastemp, IgCeilingTemp(room, stepcount), CStr(HORIZONTAL))

        Call AdjacentRoom_Ignition(qrad, ventfire, vwidth, layer1, ceilingheight2, flameheight, projection)

        'calculate view factor between ceiling and top of flame volume
        If flameheight = 0 Then
            VF = 0
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object ViewFactor_Flame(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VF = ViewFactor_Flame(projection / 2, vwidth / 2, ceilingheight2 - layer1 - flameheight)
            If 4 * VF > 1 Then VF = 1 / 4
        End If

        Qnet = -4 * VF * qrad - (h / 1000 * (gastemp - IgCeilingTemp(room, stepcount))) + QCeiling(room, i) - StefanBoltzmann / 1000 * Surface_Emissivity(1, room) * (CeilingTemp(room, i) ^ 4 - IgCeilingNode(room, 1, i) ^ 4) 'kW/m2

        cx(1, 1) = -2 * Qnet * 1000 * CeilingFourier(room) * CeilingDeltaX(room) / CeilingConductivity(room) + IgCeilingNode(room, 1, i)

        For k = 2 To Ceilingnodes - 1
            cx(k, 1) = IgCeilingNode(room, k, i)
        Next k

        'insulated rear face
        'exterior boundary conditions
        cx(Ceilingnodes, 1) = 2 * CeilingFourier(room) * CeilingOutsideBiot(room) * (ExteriorTemp - IgCeilingNode(room, Ceilingnodes, i)) + IgCeilingNode(room, Ceilingnodes, i)

        Dim ier As Short
        If frmOptions1.optLUdecom.Checked = True Then
            'find surface temperatures for the next timestep
            'using method of LU decomposition
            Call MatSol(UC, cx, Ceilingnodes)
        Else
            'find surface temperatures for the next timestep
            'using method of Gauss-Jordan elimination
            Call LINEAR2(Ceilingnodes, UC, cx, ier)
            If ier = 1 Then MsgBox("singular matrix in subroutine ceiling_ignite2_adj1.")
        End If

        For j = 1 To Ceilingnodes
            IgCeilingNode(room, j, i + 1) = cx(j, 1)
        Next j

        'store surface temps in another array
        IgCeilingTemp(room, i + 1) = IgCeilingNode(room, 1, i + 1)
        'Debug.Print IgCeilingTemp(room, i + 1) - 273, CeilingTemp(room, i + 1) - 273
    End Sub
	
    Public Sub Ceiling_Ignite_Adj2(ByRef projection As Double, ByRef room As Integer, ByVal i As Integer, ByRef IgCeilingNode(,,) As Double, ByRef layer1 As Double, ByRef ventfire As Double, ByRef ceilingheight2 As Double, ByRef vwidth As Single)
        '*  ================================================================
        '*  This function updates the surface temperatures on the ceiling above
        '*  the burner, using an implicit finite difference method.
        '*  For use with Quintiere's room corner model and a two-layer ceiling
        '*
        '*  IgCeilingNode = an array storing the temperatures within the ceiling.
        '*
        '*  Colleen Wade 26/3/99
        '*  ================================================================

        Dim k, j As Integer
        Dim CeilingFourier2, Qnet, CeilingOutsideBiot2 As Double
        Dim flameheight, gastemp, h, qrad, VF As Double
        Dim UC(2 * Ceilingnodes - 1, 2 * Ceilingnodes - 1) As Double
        Dim cx(2 * Ceilingnodes - 1, 1) As Double

        CeilingFourier2 = CeilingSubConductivity(room) * Timestep / (((CeilingSubThickness(room) / 1000) / (Ceilingnodes - 1)) ^ 2 * CeilingSubDensity(room) * CeilingSubSpecificHeat(room))
        CeilingOutsideBiot2 = OutsideConvCoeff * ((CeilingSubThickness(room) / 1000) / (Ceilingnodes - 1)) / CeilingSubConductivity(room)

        UC(1, 1) = 1 + 2 * CeilingFourier(room)
        UC(1, 2) = -2 * CeilingFourier(room)

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

        UC(2 * Ceilingnodes - 1, 2 * Ceilingnodes - 2) = -2 * CeilingFourier2
        UC(2 * Ceilingnodes - 1, 2 * Ceilingnodes - 1) = 1 + 2 * CeilingFourier2

        'get ceiling jet temperature in adjacent room
        'UPGRADE_WARNING: Couldn't resolve default property of object CeilingJet_MaxTemp_Adj(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        gastemp = CeilingJet_MaxTemp_Adj(ventfire, room, layer1)

        'convective heat transfer coefficent
        h = Get_Convection_Coefficient2(1, gastemp, IgCeilingTemp(room, stepcount), CStr(HORIZONTAL))

        Call AdjacentRoom_Ignition(qrad, ventfire, vwidth, layer1, ceilingheight2, flameheight, projection)

        'calculate view factor between ceiling and top of flame volume
        If flameheight = 0 Then
            VF = 0
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object ViewFactor_Flame(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VF = ViewFactor_Flame(projection / 2, vwidth / 2, ceilingheight2 - layer1 - flameheight)
            If 4 * VF > 1 Then VF = 1 / 4
        End If

        Qnet = -4 * VF * qrad - (h / 1000 * (gastemp - IgCeilingTemp(room, stepcount))) + QCeiling(room, i) - StefanBoltzmann / 1000 * Surface_Emissivity(1, room) * (CeilingTemp(room, i) ^ 4 - IgCeilingNode(room, 1, i) ^ 4) 'kW/m2
        'Debug.Print stepcount; Qnet
        cx(1, 1) = -2 * Qnet * 1000 * CeilingFourier(room) * CeilingDeltaX(room) / CeilingConductivity(room) + IgCeilingNode(room, 1, i)

        For k = 2 To 2 * Ceilingnodes - 2
            cx(k, 1) = IgCeilingNode(room, k, i)
        Next k

        'insulated rear face
        'exterior boundary conditions
        cx(2 * Ceilingnodes - 1, 1) = 2 * CeilingFourier2 * CeilingOutsideBiot2 * (ExteriorTemp - IgCeilingNode(room, 2 * Ceilingnodes - 1, i)) + IgCeilingNode(room, 2 * Ceilingnodes - 1, i)

        Dim ier As Short
        If frmOptions1.optLUdecom.Checked = True Then
            'find surface temperatures for the next timestep
            'using method of LU decomposition
            Call MatSol(UC, cx, 2 * Ceilingnodes - 1)
        Else
            'find surface temperatures for the next timestep
            'using method of Gauss-Jordan elimination
            Call LINEAR2(2 * Ceilingnodes - 1, UC, cx, ier)
            If ier = 1 Then MsgBox("singular matrix in subroutine ceiling_ignite2_adj2.")
        End If

        For j = 1 To 2 * Ceilingnodes - 1
            IgCeilingNode(room, j, i + 1) = cx(j, 1)
        Next j

        'store surface temps in another array
        IgCeilingTemp(room, i + 1) = IgCeilingNode(room, 1, i + 1)
        'Debug.Print IgCeilingTemp(room, i + 1) - 273, CeilingTemp(room, i + 1) - 273
    End Sub
	
	Private Sub Pyrol_Area_NextRoom(ByVal vwidth As Single, ByVal tim As Single, ByRef CeilingArea As Double, ByRef WallArea As Double, ByRef FloorArea As Double, ByVal room As Integer)
		'=================================================================
		'  This function calculates the pyrolysis area
		'  for a room adjacent to the fireroom
		'
		'  Revised 29 April 2002 Colleen Wade
		'=================================================================
		
		Dim angle, X, wall_length, AP1 As Double
		Dim jetdepth As Single
		Dim z_pyrol, APJ1 As Double
		
		jetdepth = 0.12 'depth of ceiling jet 12% of fire-ceiling height
		
		If tim >= WallIgniteTime(room) Then
			'wall has ignited
			If Y_pyrolysis(room, stepcount) > RoomHeight(room) Then
				'wall has ignited and pyrolysis front has reached the ceiling
				If CeilingIgniteFlag(room) = 0 Then
					CeilingIgniteFlag(room) = 1
                    'MDIFrmMain.ToolStripStatusLabel2.Text = "Ceiling in Room " & CStr(room) & " has ignited at " & CStr(tim) & " seconds."

                    Dim Message As String = "Ceiling in Room " & CStr(room) & " has ignited at " & CStr(tim) & " seconds."
                    MDIFrmMain.ToolStripStatusLabel2.Text = Message.ToString
                    If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                    CeilingIgniteTime(room) = tim
					CeilingIgniteStep(room) = stepcount
				End If
			End If
			
			'the wall has ignited
			If Y_pyrolysis(room, stepcount) < RoomHeight(room) Then
				'Case 1a
				'The vent fire has caused ignition of wall
				'The pyrolysis height is less than the ceiling height
				WallArea = 2 * ((Y_pyrolysis(room, stepcount) - Y_pyrolysis(room, WallIgniteStep(room) - 1)) * X_pyrolysis(room, WallIgniteStep(room)) + (X_pyrolysis(room, stepcount) - X_pyrolysis(room, WallIgniteStep(room))) * (Y_pyrolysis(room, WallIgniteStep(room)) - Y_pyrolysis(room, WallIgniteStep(room) - 1)) + 0.5 * (Y_pyrolysis(room, stepcount) - Y_pyrolysis(room, WallIgniteStep(room))) * (X_pyrolysis(room, stepcount) - X_pyrolysis(room, WallIgniteStep(room))))
				CeilingArea = 0
				Exit Sub
			Else
				'Case 1b
				'The vent fire has caused ignition of wall
				'The pyrolysis height is greater than the ceiling height
				AP1 = 2 * ((RoomHeight(room) - Y_pyrolysis(room, WallIgniteStep(room) - 1)) * X_pyrolysis(room, WallIgniteStep(room)) + (X_pyrolysis(room, stepcount) - X_pyrolysis(room, WallIgniteStep(room))) * (Y_pyrolysis(room, WallIgniteStep(room)) - Y_pyrolysis(room, WallIgniteStep(room) - 1)) + 0.5 * (X_pyrolysis(room, stepcount) - X_pyrolysis(room, WallIgniteStep(room))) * (RoomHeight(room) - Y_pyrolysis(room, WallIgniteStep(room))))
				
				'The z_pyrolysis length cannot be greater than (1-0.12) * room height
				z_pyrol = (1 - jetdepth) * (RoomHeight(room) - Y_pyrolysis(room, WallIgniteStep(room) - 1))
				If z_pyrol > X_pyrolysis(room, stepcount) - X_pyrolysis(room, WallIgniteStep(room)) Then z_pyrol = X_pyrolysis(room, stepcount) - X_pyrolysis(room, WallIgniteStep(room))
				If tim = stepcount Then
					Z_pyrolysis(room, stepcount) = z_pyrol
				End If
				
				If z_pyrol > 0 Then
					APJ1 = 2 * ((Y_pyrolysis(room, stepcount) - (RoomHeight(room) - Y_pyrolysis(room, WallIgniteStep(room) - 1))) * (jetdepth * (RoomHeight(room) - Y_pyrolysis(room, WallIgniteStep(room) - 1))) + 0.5 * z_pyrol * (Y_pyrolysis(room, stepcount) - RoomHeight(room)) - 0.5 * (jetdepth * (RoomHeight(room) - Y_pyrolysis(room, WallIgniteStep(room) - 1)) + z_pyrol) ^ 2 * ((X_pyrolysis(room, stepcount) - X_pyrolysis(room, WallIgniteStep(room))) / (RoomHeight(room) - Y_pyrolysis(room, WallIgniteStep(room)))))
				Else
					APJ1 = 2 * (Y_pyrolysis(room, stepcount) - (RoomHeight(room) - Y_pyrolysis(room, WallIgniteStep(room) - 1))) * (jetdepth * (RoomHeight(room) - Y_pyrolysis(room, WallIgniteStep(room) - 1)))
				End If
				
				WallArea = AP1 + APJ1
			End If
		Else
			WallArea = 0
		End If
		
		If tim >= CeilingIgniteTime(room) Then
			'ceiling has ignited
			If room > fireroom Then
				wall_length = WallLength2(fireroom, room, 1)
			Else
				wall_length = WallLength2(room, fireroom, 1)
			End If
			
			If Y_pyrolysis(room, stepcount) - Y_pyrolysis(room, CeilingIgniteStep(room) - 1) > wall_length / 2 Then
                X = Sqrt((Y_pyrolysis(room, stepcount) - Y_pyrolysis(room, CeilingIgniteStep(room) - 1)) ^ 2 - (wall_length / 2) ^ 2)
                angle = Atan(2 * X / wall_length) * 180 / PI
				CeilingArea = PI / 2 * (Y_pyrolysis(room, stepcount) - Y_pyrolysis(room, CeilingIgniteStep(room) - 1)) ^ 2 - 2 * (PI * (Y_pyrolysis(room, stepcount) - Y_pyrolysis(room, CeilingIgniteStep(room))) ^ 2 * angle / 360 - wall_length * X / 4)
			Else
				CeilingArea = PI / 2 * (Y_pyrolysis(room, stepcount) - Y_pyrolysis(room, CeilingIgniteStep(room) - 1)) ^ 2 'area of semi-circle
			End If
			
			If CeilingArea > RoomWidth(room) * RoomLength(room) Then CeilingArea = RoomWidth(room) * RoomLength(room)
		Else
			CeilingArea = 0
		End If
		
		'If tim >= FloorIgniteTime(room) Then
		'    'floor has ignited
		'    If vwidth > 2 * Xf_pyrolysis(fireroom, FloorIgniteStep(room)) Then
		'       FloorArea = Yf_pyrolysis(room, stepcount) * (2 * Xf_pyrolysis(room, FloorIgniteStep(room)))
		'        FloorArea = FloorArea + 2 * (Xf_pyrolysis(room, stepcount) - Xf_pyrolysis(room, FloorIgniteStep(room))) * Yf_pyrolysis(room, stepcount)
		'        FloorArea = FloorArea + (Xf_pyrolysis(room, stepcount) - Xf_pyrolysis(room, FloorIgniteStep(room))) * (Yf_pyrolysis(room, stepcount) - Yf_pyrolysis(room, FloorIgniteStep(room)))
		'    Else
		'        FloorArea = Yf_pyrolysis(room, stepcount) * vwidth
		'        FloorArea = FloorArea + 2 * (Xf_pyrolysis(room, stepcount) - vwidth / 2) * Yf_pyrolysis(room, stepcount)
		'        FloorArea = FloorArea + (Xf_pyrolysis(room, stepcount) - vwidth / 2) * (Yf_pyrolysis(room, stepcount) - Yf_pyrolysis(room, FloorIgniteStep(room)))
		'    End If
		'Else
		'    FloorArea = 0
		'End If
		
		If tim >= FloorIgniteTime(room) Then
			'floor has ignited
			If room > fireroom Then
				wall_length = WallLength2(fireroom, room, 1)
			Else
				wall_length = WallLength2(room, fireroom, 1)
			End If
			
			If Yf_pyrolysis(room, stepcount) - Yf_pyrolysis(room, FloorIgniteStep(room) - 1) > wall_length / 2 Then
                X = Sqrt((Yf_pyrolysis(room, stepcount) - Yf_pyrolysis(room, FloorIgniteStep(room) - 1)) ^ 2 - (wall_length / 2) ^ 2)
                angle = Atan(2 * X / wall_length) * 180 / PI
				FloorArea = PI / 2 * (Yf_pyrolysis(room, stepcount) - Yf_pyrolysis(room, FloorIgniteStep(room) - 1)) ^ 2 - 2 * (PI * (Yf_pyrolysis(room, stepcount) - Yf_pyrolysis(room, FloorIgniteStep(room))) ^ 2 * angle / 360 - wall_length * X / 4)
			Else
				FloorArea = PI / 2 * (Yf_pyrolysis(room, stepcount) - Yf_pyrolysis(room, FloorIgniteStep(room) - 1)) ^ 2 'area of semi-circle
			End If
			
			If FloorArea > RoomWidth(room) * RoomLength(room) Then FloorArea = RoomWidth(room) * RoomLength(room)
		Else
			FloorArea = 0
		End If
		
	End Sub
	
    Public Function HRR_nextroom(ByVal vwidth As Single, ByVal room As Integer, ByVal tim As Double, ByVal flux_wall_AR As Double) As Double
        '******************************************************
        '*  Find the total heat release rate for use by Quintiere's Room Corner Model
        '*
        '*  tim = current time
        '*  Q = hrr to use in the plume correlation kW/m2
        '*
        '*  5 April 1998 Colleen Wade
        '******************************************************

        Static HRRlast, timlast, Qlast As Double
        Dim area1, hrr, Area, area2 As Double
        Dim QNormalF, QNormalW, burnouttime, QNormalC, QNormalFB As Double
        Dim QP_ceiling, FloorArea, WallArea, CeilingArea, QP_wall, QP_floor As Double
        Dim QNormalWB, QP_wall_B, QP_floor_B, h As Double
        Dim QCeiling, QWall, QFloor As Single
        Dim qstar, deltaZ As Double
        Dim floorflux_f, NetFlux_F, NetFLux_B, floorflux_b As Single
        On Error GoTo HRRerrorhandler

        If tim >= CeilingIgniteTime(room) Then IgCeilingTemp(room, stepcount) = IgTempC(room)
        If tim >= WallIgniteTime(room) Then IgWallTemp(room, stepcount) = IgTempW(room)
        If tim >= FloorIgniteTime(room) Then IgFloorTemp(room, stepcount) = IgTempF(room)

        'energy release rate for material (peak)

        If frmQuintiere.optUseOneCurve.Checked = True Then
            Call NewPeak(flux_wall_AR, room, QP_wall, QP_ceiling, QP_wall_B, QP_floor, QP_floor_B)
        Else
            Call Get_NetFlux(flux_wall_AR, room, NetFlux_F, NetFLux_B, floorflux_f, floorflux_b)
        End If

        'find the pyrolysis area
        Call Pyrol_Area_NextRoom(vwidth, tim, CeilingArea, WallArea, FloorArea, room)
        If stepcount > 1 Then
            'keep track of the total pyrolysis area and the incremental area at each time step
            'pyrolarea(1, stepcount) = WallArea
            'delta_area(1, stepcount) = pyrolarea(1, stepcount) - pyrolarea(1, stepcount - 1)
            If CeilingEffectiveHeatofCombustion(room) > 0 Then
                pyrolarea(room, 2, stepcount) = CeilingArea
                If pyrolarea(room, 2, stepcount) > pyrolarea(room, 2, stepcount - 1) Then
                    delta_area(room, 2, stepcount) = pyrolarea(room, 2, stepcount) - pyrolarea(room, 2, stepcount - 1)
                Else
                    delta_area(room, 2, stepcount) = 0
                End If
            End If
            If WallEffectiveHeatofCombustion(room) > 0 Then
                pyrolarea(room, 1, stepcount) = WallArea
                If pyrolarea(room, 1, stepcount) > pyrolarea(room, 1, stepcount - 1) Then
                    delta_area(room, 1, stepcount) = pyrolarea(room, 1, stepcount) - pyrolarea(room, 1, stepcount - 1)
                Else
                    delta_area(room, 1, stepcount) = 0
                End If
            End If
            If FloorEffectiveHeatofCombustion(room) > 0 Then
                pyrolarea(room, 3, stepcount) = FloorArea
                If pyrolarea(room, 3, stepcount) > pyrolarea(room, 3, stepcount - 1) Then
                    delta_area(room, 3, stepcount) = pyrolarea(room, 3, stepcount) - pyrolarea(room, 3, stepcount - 1)
                Else
                    delta_area(room, 3, stepcount) = 0
                End If
            End If
        End If

        If frmQuintiere.optUseOneCurve.Checked = True Then
            'UPGRADE_WARNING: Couldn't resolve default property of object hrr_dave_nextroom(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            hrr = hrr_dave_nextroom(room, tim, QP_wall, QP_wall_B, QP_floor, QP_floor_B, QP_ceiling, QNormalW, QNormalWB, QNormalC, QNormalF, QNormalFB)
            'Debug.Print tim, hrr; "old"
            If flagstop = 1 Then Exit Function
        Else
            'hrr = hrr_linings_nextroom(ByVal room, ByVal NetFlux_F, ByVal NetFLux_B, ByVal FloorFlux_F, QWall, QCeiling, QFloor)
            'UPGRADE_WARNING: Couldn't resolve default property of object hrr_linings2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            hrr = hrr_linings2(room, tim, floorflux_f, floorflux_b, NetFlux_F, NetFLux_B, QFloor, QWall, QCeiling, 0, 0, 0)
            'Debug.Print tim, hrr; "new"
        End If

        'function returns the required heat release rate
        'UPGRADE_WARNING: Couldn't resolve default property of object HRR_nextroom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        HRR_nextroom = hrr

        Exit Function

HRRerrorhandler:
        MsgBox(ErrorToString(Err.Number) & " HRR_nextroom")
        ERRNO = Err.Number
        'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'System.Windows.Forms.Cursor.Current = Default_Renamed
    End Function
	
	Public Sub Get_NetFlux(ByVal flux_wall_AR As Double, ByVal room As Integer, ByRef NetFlux_F As Single, ByRef NetFLux_B As Single, ByRef floorflux_f As Single, ByRef floorflux_b As Single)
		'*********************************************************
		'*  This function calculates the  heat flux exposing the lining
		'*  For use in Quintiere's room corner model.
		'********************************************************
		
		Dim max As Double
		
		If room = fireroom Then
			max = 0
			
			'the Q term includes a rerad loss so we add it back on again to get the incident flux
			NetFlux_F = (QPyrol - QUpperWall(room, stepcount) + StefanBoltzmann / 1000 * Surface_Emissivity(2, room) * (Upperwalltemp(room, stepcount) ^ 4 - max ^ 4))
			NetFLux_B = (QFB - QUpperWall(room, stepcount) + StefanBoltzmann / 1000 * Surface_Emissivity(2, room) * (Upperwalltemp(room, stepcount) ^ 4 - max ^ 4))
			'floorflux = (QFB - QFloor((room), stepcount) + StefanBoltzmann / 1000 * Surface_Emissivity(4, (room)) * (FloorTemp((room), stepcount) ^ 4 - max ^ 4))
			
			'view factor expression in here
			
			floorflux_f = (QPyrol - QFloor(room, stepcount) + StefanBoltzmann / 1000 * Surface_Emissivity(4, room) * (FloorTemp(room, stepcount) ^ 4 - max ^ 4))
			floorflux_b = (QFB - QFloor(room, stepcount) + StefanBoltzmann / 1000 * Surface_Emissivity(4, room) * (FloorTemp(room, stepcount) ^ 4 - max ^ 4))
		Else
			NetFlux_F = (QPyrol - QUpperWall(room, stepcount) + StefanBoltzmann / 1000 * Surface_Emissivity(2, room) * (Upperwalltemp(room, stepcount) ^ 4 - max ^ 4))
			NetFLux_B = (flux_wall_AR - QUpperWall(room, stepcount) + StefanBoltzmann / 1000 * Surface_Emissivity(2, room) * (Upperwalltemp(room, stepcount) ^ 4 - max ^ 4))
			
			floorflux_f = (QPyrol - QFloor(room, stepcount) + StefanBoltzmann / 1000 * Surface_Emissivity(4, room) * (FloorTemp(room, stepcount) ^ 4 - max ^ 4))
			floorflux_b = (QPyrol - QFloor(room, stepcount) + StefanBoltzmann / 1000 * Surface_Emissivity(4, room) * (FloorTemp(room, stepcount) ^ 4 - max ^ 4))
		End If
	End Sub
	
    Public Function hrr_linings(ByVal floorflux As Single, ByVal NetFlux_F As Single, ByVal NetFLux_B As Single, ByRef QFloor As Single, ByRef QWall As Single, ByRef QCeiling As Single, ByRef QDotW As Double, ByRef QDotC As Double, ByRef QDotF As Double) As Double
        '====================================================================
        ' Determines the heat release rate from the compartment lining materials
        ' by storing the change in pyrolysis area for each time step and
        ' using the elapsed time to determine the heat release rate for each
        ' burning segment from the cone calorimeter input
        '
        '====================================================================

        Dim i, j As Integer
        Dim hrr As Double
        Dim s() As Double
        Dim index() As Integer
        Dim x2, x1, x4 As Single
        Dim burnarea As Double
        Dim curve As Integer
        Dim YW(1) As Double
        Dim YC(1) As Double
        Dim YF(1) As Double
        Dim T, x3, elapsed As Single
        'UPGRADE_NOTE: step was upgraded to step_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim QDotWB As Double
        Dim step_Renamed, lastpoint As Integer
        Dim ExtFlux() As Double

        If CurveNumber_W(fireroom) > 0 Then ReDim YW(CurveNumber_W(fireroom))
        If CurveNumber_C(fireroom) > 0 Then ReDim YC(CurveNumber_C(fireroom))
        If CurveNumber_F(fireroom) > 0 Then ReDim YF(CurveNumber_F(fireroom))
        On Error GoTo errorhandler

        'start at the timestep after when the a lining is first ignited

        i = stepcount
        If WallIgniteStep(fireroom) > 0 Then i = WallIgniteStep(fireroom) + 1 'exclude the area first ignited, see below
        If CeilingIgniteStep(fireroom) > 0 Then
            If CeilingIgniteStep(fireroom) < i Then i = CeilingIgniteStep(fireroom)
        End If
        If FloorIgniteStep(fireroom) > 0 Then
            If FloorIgniteStep(fireroom) < i Then i = FloorIgniteStep(fireroom)
        End If
        'If CeilingIgniteStep(fireroom) < WallIgniteStep(fireroom) Then i = CeilingIgniteStep(fireroom) + 1
        'If FloorIgniteStep(fireroom) < i Then i = FloorIgniteStep(fireroom)

        hrr = 0
        QWall = 0
        QCeiling = 0
        QFloor = 0

        'the external heat flux has to be within the range of the cone data available
        If CurveNumber_W(fireroom) > 0 Then
            x1 = NetFLux_B
            If x1 < ExtFlux_W(1, fireroom) Then x1 = ExtFlux_W(1, fireroom)
            If x1 > ExtFlux_W(CurveNumber_W(fireroom), fireroom) Then x1 = ExtFlux_W(CurveNumber_W(fireroom), fireroom)
            x2 = NetFlux_F
            If x2 < ExtFlux_W(1, fireroom) Then x2 = ExtFlux_W(1, fireroom)
            If x2 > ExtFlux_W(CurveNumber_W(fireroom), fireroom) Then x2 = ExtFlux_W(CurveNumber_W(fireroom), fireroom)
        End If
        If CurveNumber_C(fireroom) > 0 Then
            x3 = NetFlux_F
            If x3 < ExtFlux_C(1, fireroom) Then x3 = ExtFlux_C(1, fireroom)
            If x3 > ExtFlux_C(CurveNumber_C(fireroom), fireroom) Then x3 = ExtFlux_C(CurveNumber_C(fireroom), fireroom)
        End If
        If CurveNumber_F(fireroom) > 0 Then
            x4 = floorflux
            If x4 < ExtFlux_F(1, fireroom) Then x4 = ExtFlux_F(1, fireroom)
            If x4 > ExtFlux_F(CurveNumber_F(fireroom), fireroom) Then x4 = ExtFlux_F(CurveNumber_F(fireroom), fireroom)
        End If


        Do While i < stepcount
            'burnarea = delta_area(fireroom, 1, i) + delta_area(fireroom, 2, i) + delta_area(fireroom, 3, i)
            'If burnarea > 0 Then
            'something is burning
            'wall
            'Call Interpolate_hrr(ByVal i - WallIgniteStep, ExtFlux(), ByVal "wall", ByVal NetFlux_F, QDotW)
            'If delta_area(fireroom, 1, i) > 0 Then
            If WallIgniteFlag(fireroom) > 0 Then
                'elapsed = tim(i, 1) - WallIgniteTime(fireroom)
                elapsed = tim(stepcount, 1) - tim(i, 1) 'section i started burning at tim(i,1) and tim(stepcount,1) is the curent time
                'wall is burning -
                For curve = 1 To CurveNumber_W(fireroom)

                    lastpoint = ConeNumber_W(curve, fireroom)
                    Do While ConeYW(fireroom, curve, lastpoint) = 0
                        lastpoint = lastpoint - 1
                    Loop
                    'step = i - WallIgniteStep(fireroom)
                    'If step > ConeNumber_W(curve) Then step = ConeNumber_W(curve)
                    If elapsed > ConeYW(fireroom, curve, lastpoint) Then
                        YW(curve) = ConeXW(fireroom, curve, lastpoint) 'hrr
                    Else
                        Call Cone_Data(fireroom, 1, curve, elapsed, YW(curve))
                        'For T = 1 To ConeNumber_W(fireroom, curve)
                        '    If ConeYW(fireroom, curve, T) >= elapsed Then
                        '        YW(curve) = ConeXW(fireroom, curve, T)
                        '        Exit For
                        '    End If
                        'Next T
                    End If
                Next curve

                ReDim ExtFlux(CurveNumber_W(fireroom))
                For j = 1 To CurveNumber_W(fireroom)
                    ExtFlux(j) = ExtFlux_W(j, fireroom)
                Next j

                ReDim s(CurveNumber_W(fireroom))
                ReDim index(CurveNumber_W(fireroom))
                If CurveNumber_W(fireroom) > 1 Then
                    Call CSCOEF(CurveNumber_W(fireroom), ExtFlux, YW, s, index)
                    Call CSFIT1(CurveNumber_W(fireroom), ExtFlux, YW, s, index, x2, QDotW)
                Else
                    QDotW = YW(1)
                End If
            End If

            'if the ceiling material is the same as the wall, then we can treat then as a continuous surface
            If (WallConeDataFile(fireroom) <> CeilingConeDataFile(fireroom)) Or FireLocation(1) = 0 Or CeilingIgniteTime(fireroom) < WallIgniteTime(fireroom) Then
                'if they are different materials
                'If delta_area(fireroom, 2, i) > 0 Then
                If CeilingIgniteFlag(fireroom) = True Then
                    elapsed = tim(i, 1) - CeilingIgniteTime(fireroom)

                    'ceiling
                    For curve = 1 To CurveNumber_C(fireroom)

                        lastpoint = ConeNumber_C(curve, fireroom)
                        Do While ConeYC(fireroom, curve, lastpoint) = 0
                            lastpoint = lastpoint - 1
                        Loop

                        If elapsed > ConeYC(fireroom, curve, lastpoint) Then
                            YC(curve) = ConeXC(fireroom, curve, lastpoint) 'hrr
                        Else
                            Call Cone_Data(fireroom, 2, curve, elapsed, YC(curve))
                        End If
                    Next curve

                    ReDim ExtFlux(CurveNumber_C(fireroom))
                    For j = 1 To CurveNumber_C(fireroom)
                        ExtFlux(j) = ExtFlux_C(j, fireroom)
                    Next j

                    ReDim s(CurveNumber_C(fireroom))
                    ReDim index(CurveNumber_C(fireroom))
                    If CurveNumber_C(fireroom) > 1 Then
                        Call CSCOEF(CurveNumber_C(fireroom), ExtFlux, YC, s, index)
                        Call CSFIT1(CurveNumber_C(fireroom), ExtFlux, YC, s, index, x3, QDotC)
                    Else
                        QDotC = YC(1)
                    End If
                Else
                    QDotC = 0
                End If
            Else
                QDotC = QDotW
            End If

            'If delta_area(fireroom, 3, i) > 0 Then
            'floor
            If FloorIgniteFlag(fireroom) = 1 Then
                'floor
                elapsed = tim(i, 1) - FloorIgniteTime(fireroom)
                For curve = 1 To CurveNumber_F(fireroom)

                    lastpoint = ConeNumber_F(curve, fireroom)
                    Do While ConeYF(fireroom, curve, lastpoint) = 0
                        lastpoint = lastpoint - 1
                    Loop

                    If elapsed > ConeYF(fireroom, curve, lastpoint) Then
                        YF(curve) = ConeXF(fireroom, curve, lastpoint) 'hrr
                    Else
                        'get the hrr values for each heat flux after the elapsed time so they can be interpolated
                        Call Cone_Data(fireroom, 3, curve, elapsed, YF(curve))
                    End If
                Next curve

                ReDim ExtFlux(CurveNumber_F(fireroom))
                For j = 1 To CurveNumber_F(fireroom)
                    ExtFlux(j) = ExtFlux_F(j, fireroom)
                Next j

                ReDim s(CurveNumber_F(fireroom))
                ReDim index(CurveNumber_F(fireroom))
                If CurveNumber_F(fireroom) > 1 Then
                    Call CSCOEF(CurveNumber_F(fireroom), ExtFlux, YF, s, index)
                    Call CSFIT1(CurveNumber_F(fireroom), ExtFlux, YF, s, index, x4, QDotF)
                Else
                    QDotF = YF(1)
                End If
            Else
                QDotF = 0
            End If

            'QFloor = QDotF * RoomFloorArea(fireroom)
            'Else
            'QDotF = 0
            'End If
            'End If
            'Debug.Print i, elapsed, QDotF, delta_area(fireroom, 3, stepcount + FloorIgniteTime(fireroom) - i), stepcount
            'Debug.Print i, elapsed, QDotW, delta_area(fireroom, 1, stepcount + WallIgniteStep(fireroom) - i), stepcount
            'Debug.Print i, elapsed, QDotC, delta_area(fireroom, 2, i), stepcount
            QWall = QWall + QDotW * delta_area(fireroom, 1, i) 'excludes the area first ignited by the burner
            'QCeiling = QCeiling + QDotC * delta_area(fireroom, 2, stepcount + CeilingIgniteStep(fireroom) - i)
            'QFloor = QFloor + QDotF * delta_area(fireroom, 3, stepcount + FloorIgniteStep(fireroom) - i)
            i = i + 1
        Loop



        If WallIgniteStep(fireroom) > 0 And (WallConeDataFile(fireroom) <> "" And WallConeDataFile(fireroom) <> "null.txt") Then
            'call routine to get the qdot from the wall in the burner region = QDotWB
            If stepcount > WallIgniteStep(fireroom) Then
                'Call Interpolate_hrr(stepcount - WallIgniteStep, ExtFlux(), "wall", NetFlux_B, QDotWB)
                For curve = 1 To CurveNumber_W(fireroom)
                    'step = i - WallIgniteStep(fireroom)
                    'If step > ConeNumber_W(curve) Then step = ConeNumber_W(curve)
                    'YW(curve) = ConeXW(curve, step)
                    elapsed = tim(stepcount, 1) - WallIgniteTime(fireroom)

                    If elapsed > ConeYW(fireroom, curve, ConeNumber_W(curve, fireroom)) Then
                        YW(curve) = ConeXW(fireroom, curve, ConeNumber_W(curve, fireroom)) 'hrr
                    Else
                        Call Cone_Data(fireroom, 1, curve, elapsed, YW(curve))
                        'For T = 1 To ConeNumber_W(fireroom, curve)
                        '    If ConeYW(fireroom, curve, T) >= elapsed Then
                        '        YW(curve) = ConeXW(fireroom, curve, T)
                        '        Exit For
                        '    End If
                        'Next T
                    End If
                Next curve

                ReDim ExtFlux(CurveNumber_W(fireroom))
                For j = 1 To CurveNumber_W(fireroom)
                    ExtFlux(j) = ExtFlux_W(j, fireroom)
                Next j

                ReDim s(CurveNumber_W(fireroom))
                ReDim index(CurveNumber_W(fireroom))
                'If CurveNumber_C(fireroom) > 1 Then
                If CurveNumber_W(fireroom) > 1 Then
                    Call CSCOEF(CurveNumber_W(fireroom), ExtFlux, YW, s, index)
                    Call CSFIT1(CurveNumber_W(fireroom), ExtFlux, YW, s, index, x1, QDotWB)
                Else
                    QDotWB = YW(1)
                End If
            Else
                QDotWB = 0
            End If
            QWall = QWall + QDotWB * delta_area(fireroom, 1, WallIgniteStep(fireroom))
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object hrr_linings. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        hrr_linings = QWall + QCeiling + QFloor 'kw
        Erase YW
        Erase YC
        Erase YF
        Erase ExtFlux
        Exit Function

errorhandler:

HRRerrorhandler:
        MsgBox(ErrorToString(Err.Number))
        ERRNO = Err.Number
        'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'System.Windows.Forms.Cursor.Current = Default_Renamed

    End Function
	
    Public Sub Wall_Ignite_Adj1(ByVal layerz As Single, ByVal vwidth As Single, ByVal fheight As Single, ByVal Tflame As Double, ByVal eFLame As Double, ByVal room As Integer, ByVal i As Integer, ByRef IgWallNode(,,) As Double, ByRef flux_wall_AR As Double, ByRef awallflux_r As Double)
        '*  ================================================================
        '*  This function updates the surface temperatures on the ceiling above the
        '*  the burner, using an implicit finite difference method.
        '*
        '*  igwallnode = an array storing the temperatures within the ceiling.
        '*
        '*  Colleen Wade 24/7/98
        '*  ================================================================

        Dim k, j As Integer
        Dim h, Qnet As Double
        Dim UW(Wallnodes, Wallnodes) As Double
        Dim wx(Wallnodes, 1) As Double

        UW(1, 1) = 1 + 2 * WallFourier(room)
        UW(1, 2) = -2 * WallFourier(room)

        k = 2
        For j = 2 To Wallnodes - 1
            UW(j, k - 1) = -WallFourier(room)
            UW(j, k) = 1 + 2 * WallFourier(room)
            UW(j, k + 1) = -WallFourier(room)
            k = k + 1
        Next j

        UW(Wallnodes, Wallnodes - 1) = -2 * WallFourier(room)
        UW(Wallnodes, Wallnodes) = 1 + 2 * WallFourier(room)

        'get ceiling jet temperature in adjacent room
        'convective heat transfer coefficent
        If Tflame > uppertemp(room, i) Then
            'h = Get_Convection_Coefficient2(1, Tflame, IgWallTemp(room, i-1), VERTICAL)
            'UPGRADE_WARNING: Couldn't resolve default property of object forced_h(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            h = forced_h(IgWallTemp(room, i - 1), tim(i, 1), fheight, Tflame, stepcount, vwidth, layerz, uppertemp(fireroom, i))
        Else
            h = Get_Convection_Coefficient2(1, uppertemp(room, i), IgWallTemp(room, i), CStr(VERTICAL))
        End If

        If Tflame > uppertemp(room, i) Then
            'Qnet = -eFLame * StefanBoltzmann / 1000 * Tflame ^ 4 - h / 1000 * (Tflame - IgWallTemp(room, i)) + QUpperWall(room, i) - StefanBoltzmann / 1000 * Surface_Emissivity(2, room) * (Upperwalltemp(room, i) ^ 4 - IgWallTemp(room, i) ^ 4)
            Qnet = -eFLame * awallflux_r - h / 1000 * (Tflame - IgWallTemp(room, i)) + QUpperWall(room, i) - StefanBoltzmann / 1000 * Surface_Emissivity(2, room) * (Upperwalltemp(room, i) ^ 4 - (IgWallTemp(room, i)) ^ 4)
        Else
            Qnet = -h / 1000 * (uppertemp(room, i) - IgWallTemp(room, i)) + QUpperWall(room, i) - StefanBoltzmann / 1000 * Surface_Emissivity(2, room) * (Upperwalltemp(room, i) ^ 4 - (IgWallTemp(room, i)) ^ 4)
        End If

        'flux to the wall excluding the radiation term
        flux_wall_AR = -(Qnet - StefanBoltzmann / 1000 * Surface_Emissivity(2, room) * (IgWallTemp(room, i) ^ 4))

        If WallIgniteFlag(room) = 1 Then
            IgWallTemp(room, i + 1) = IgWallTemp(room, i)
            Exit Sub 'so we can get the flux
        End If

        wx(1, 1) = -2 * Qnet * 1000 * WallFourier(room) * WallDeltaX(room) / WallConductivity(room) + IgWallNode(room, 1, i)

        For k = 2 To Wallnodes - 1
            wx(k, 1) = IgWallNode(room, k, i)
        Next k

        'rear face is the wall in the fire compartment
        'exterior boundary conditions
        wx(Wallnodes, 1) = Upperwalltemp(fireroom, i)

        Dim ier As Short
        If frmOptions1.optLUdecom.Checked = True Then
            'find surface temperatures for the next timestep
            'using method of LU decomposition
            Call MatSol(UW, wx, Wallnodes)
        Else
            'find surface temperatures for the next timestep
            'using method of Gauss-Jordan elimination
            Call LINEAR2(Wallnodes, UW, wx, ier)
            If ier = 1 Then MsgBox("singular matrix in subroutine wall_ignite_adj1.")
        End If

        For j = 1 To Wallnodes
            IgWallNode(room, j, i + 1) = wx(j, 1)
        Next j

        'store surface temps in another array
        IgWallTemp(room, i + 1) = IgWallNode(room, 1, i + 1)
        'Debug.Print IgWallTemp(room, i + 1) - 273

    End Sub
	
    Public Sub Wall_Ignite_Adj2(ByVal layerz As Single, ByVal vwidth As Single, ByVal fheight As Single, ByVal Tflame As Double, ByVal eFLame As Double, ByVal room As Integer, ByVal i As Integer, ByRef IgWallNode(,,) As Double, ByRef flux_wall_AR As Double, ByRef awallflux_r As Object)
        '*  ================================================================
        '*  This function updates the surface temperatures on the wall next to
        '*  the burner, using an implicit finite difference method.
        '*  For use with Quintiere's room corner model.
        '*
        '*  IgWallNode = an array storing the temperatures within the wall adjacent to the burner
        '*
        '*  Colleen Wade 24/7/98
        '*  ================================================================

        Dim k, j As Integer
        Dim WallFourier2, Qnet, WallOutsideBiot2 As Double
        Dim h As Double
        Dim UW(2 * Wallnodes - 1, 2 * Wallnodes - 1) As Double
        Dim wx(2 * Wallnodes - 1, 1) As Double

        WallFourier2 = WallSubConductivity(room) * Timestep / (((WallSubThickness(room) / 1000) / (Wallnodes - 1)) ^ 2 * WallSubDensity(room) * WallSubSpecificHeat(room))
        WallOutsideBiot2 = OutsideConvCoeff * ((WallSubThickness(room) / 1000) / (Wallnodes - 1)) / WallSubConductivity(room)

        UW(1, 1) = 1 + 2 * WallFourier(room)
        UW(1, 2) = -2 * WallFourier(room)

        k = 2
        For j = 2 To Wallnodes - 1
            UW(j, k - 1) = -WallFourier(room)
            UW(j, k) = 1 + 2 * WallFourier(room)
            UW(j, k + 1) = -WallFourier(room)
            k = k + 1
        Next j

        j = Wallnodes
        UW(j, k - 1) = -WallFourier(room)
        UW(j, k) = 1 + WallFourier(room) + WallFourier2
        UW(j, k + 1) = -WallFourier2
        k = k + 1

        For j = Wallnodes + 1 To 2 * Wallnodes - 2
            UW(j, k - 1) = -WallFourier2
            UW(j, k) = 1 + 2 * WallFourier2
            UW(j, k + 1) = -WallFourier2
            k = k + 1
        Next j

        UW(2 * Wallnodes - 1, 2 * Wallnodes - 2) = -2 * WallFourier2
        UW(2 * Wallnodes - 1, 2 * Wallnodes - 1) = 1 + 2 * WallFourier2

        'get ceiling jet temperature in adjacent room
        'convective heat transfer coefficent
        'If i > 1 Then
        If Tflame > uppertemp(room, i) Then
            'h = Get_Convection_Coefficient2(1, Tflame, IgWallTemp(room, i-1), VERTICAL)
            'UPGRADE_WARNING: Couldn't resolve default property of object forced_h(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            h = forced_h(IgWallTemp(room, i - 1), tim(i, 1), fheight, Tflame, stepcount, vwidth, layerz, uppertemp(fireroom, i))
        Else
            h = Get_Convection_Coefficient2(1, uppertemp(room, i), IgWallTemp(room, i), CStr(VERTICAL))
        End If

        If Tflame > uppertemp(room, i) Then
            'Qnet = -eFLame * StefanBoltzmann / 1000 * Tflame ^ 4 - h / 1000 * (Tflame - IgWallTemp(room, i)) + QUpperWall(room, i) - StefanBoltzmann / 1000 * Surface_Emissivity(2, room) * (Upperwalltemp(room, i) ^ 4 - IgWallTemp(room, i) ^ 4)
            'UPGRADE_WARNING: Couldn't resolve default property of object awallflux_r. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Qnet = -eFLame * awallflux_r - h / 1000 * (Tflame - IgWallTemp(room, i)) + QUpperWall(room, i) - StefanBoltzmann / 1000 * Surface_Emissivity(2, room) * (Upperwalltemp(room, i) ^ 4 - IgWallTemp(room, i) ^ 4)
        Else
            '-ve is into the wall
            Qnet = -h / 1000 * (uppertemp(room, i) - IgWallTemp(room, i)) + QUpperWall(room, i) - StefanBoltzmann / 1000 * Surface_Emissivity(2, room) * (Upperwalltemp(room, i) ^ 4 - IgWallTemp(room, i) ^ 4)
        End If

        'flux to the wall excluding the radiation term
        flux_wall_AR = -(Qnet - StefanBoltzmann / 1000 * Surface_Emissivity(2, room) * (IgWallTemp(room, i) ^ 4))
        'flux_wall_AR = h / 1000 * (uppertemp(room, i) - IgWallTemp(room, i)) - QUpperWall(room, i) + StefanBoltzmann / 1000 * Surface_Emissivity(2, room) * (Upperwalltemp(room, i) ^ 4 + uppertemp(room, i) ^ 4)

        'Else
        '   h = 0
        '   Qnet = 0
        '    flux_wall_ar = 0
        'End If

        If WallIgniteFlag(room) = 1 Then
            IgWallTemp(room, i + 1) = IgWallTemp(room, i)
            Exit Sub
        End If

        wx(1, 1) = -2 * Qnet * 1000 * WallFourier(room) * WallDeltaX(room) / WallConductivity(room) + IgWallTemp(room, i)

        For k = 2 To 2 * Wallnodes - 2
            wx(k, 1) = IgWallNode(room, k, i)
        Next k

        'rear face is the wall of the fireroom
        'exterior boundary conditions
        'wx(2 * Wallnodes - 1, 1) = 2 * WallFourier2 * WallOutsideBiot2 * (ExteriorTemp - IgWallNode(room, 2 * Wallnodes - 1, i)) + IgWallNode(room, 2 * Wallnodes - 1, i)
        wx(2 * Wallnodes - 1, 1) = Upperwalltemp(fireroom, i)

        Dim ier As Short
        If frmOptions1.optLUdecom.Checked = True Then
            'find surface temperatures for the next timestep
            'using method of LU decomposition
            Call MatSol(UW, wx, 2 * Wallnodes - 1) 'wall
        Else
            'find surface temperatures for the next timestep
            'using method of Gauss-Jordan elimination
            Call LINEAR2(2 * Wallnodes - 1, UW, wx, ier)
            If ier = 1 Then MsgBox("singular matrix in implicit_surface_temps")
        End If

        For j = 1 To 2 * Wallnodes - 1
            IgWallNode(room, j, i + 1) = wx(j, 1)
        Next j

        'store surface temps in another array
        IgWallTemp(room, i + 1) = IgWallNode(room, 1, i + 1)
        'Debug.Print IgWallTemp(room, i + 1) - 273

    End Sub
	
	Public Function hrr_linings_nextroom(ByVal room As Integer, ByVal NetFlux_F As Single, ByVal NetFLux_B As Single, ByVal floorflux_f As Single, ByRef QWall As Single, ByRef QCeiling As Single, ByRef QFloor As Single) As Object
		'====================================================================
		' Determines the heat release rate from the compartment lining materials
		' by storing the change in pyrolysis area for each time step and
		' using the elapsed time to determine the heat release rate for each
		' burning segment from the cone calorimeter input
		'
		'====================================================================
		
		Dim i, j As Integer
		Dim hrr As Double
		Dim s() As Double
		Dim index() As Integer
		Dim x3, x1, x2 As Single
		Dim burnarea As Double
		Dim curve As Integer
		Dim YW() As Double
		Dim YC() As Double
		Dim YF() As Double
		Dim ExtFlux() As Double
		'UPGRADE_NOTE: step was upgraded to step_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim QDotF, QDotW, QDotC, QDotWB As Double
		Dim step_Renamed As Integer
		Dim elapsed, T As Single
		ReDim YW(CurveNumber_W(room))
		ReDim YC(CurveNumber_C(room))
		ReDim YF(CurveNumber_C(room))
		On Error GoTo errorhandler
		
		'start at the timestep after when the wall first ignited
		i = WallIgniteStep(room) + 1
		hrr = 0
		QWall = 0
		QCeiling = 0
		
		x1 = NetFLux_B
		If x1 < ExtFlux_W(1, room) Then x1 = ExtFlux_W(1, room)
		If x1 > ExtFlux_W(CurveNumber_W(room), room) Then x1 = ExtFlux_W(CurveNumber_W(room), room)
		x2 = NetFlux_F
		If x2 < ExtFlux_W(1, room) Then x2 = ExtFlux_W(1, room)
		If x2 > ExtFlux_W(CurveNumber_W(room), room) Then x2 = ExtFlux_W(CurveNumber_W(room), room)
		x3 = NetFlux_F
		If x3 < ExtFlux_C(1, room) Then x3 = ExtFlux_C(1, room)
		If x3 > ExtFlux_C(CurveNumber_C(room), room) Then x3 = ExtFlux_C(CurveNumber_C(room), room)
		
		Do While i < stepcount
			'elapsed = timenow - tim(i, 1)
			burnarea = delta_area(room, 1, i) + delta_area(room, 2, i)
			If burnarea > 0 Then
				'wall
				'Call Interpolate_hrr(ByVal i - WallIgniteStep, ExtFlux(), ByVal "wall", ByVal NetFlux_F, QDotW)
				For curve = 1 To CurveNumber_W(room)
					'step = i - WallIgniteStep(room)
					'If step > ConeNumber_W(room, curve) Then step = ConeNumber_W(room, curve)
					'YW(curve) = ConeXW(room, curve, step)
					elapsed = tim(i, 1) - WallIgniteTime(room)
					If elapsed > ConeYW(room, curve, ConeNumber_W(curve, room)) Then
						YW(curve) = ConeXW(room, curve, ConeNumber_W(curve, room)) 'hrr
					Else
						Call Cone_Data(room, 1, curve, elapsed, YW(curve))
						'For T = 1 To ConeNumber_W(room, curve)
						'    If ConeYW(room, curve, T) >= elapsed Then
						'        YW(curve) = ConeXW(room, curve, T)
						'        Exit For
						'    End If
						'Next T
					End If
				Next curve
				
				ReDim ExtFlux(CurveNumber_W(room))
				For j = 1 To CurveNumber_W(room)
					ExtFlux(j) = ExtFlux_W(j, room)
				Next j
				
				ReDim s(CurveNumber_W(room))
				ReDim index(CurveNumber_W(room))
				If CurveNumber_W(room) > 1 Then
					Call CSCOEF(CurveNumber_W(room), ExtFlux, YW, s, index)
					Call CSFIT1(CurveNumber_W(room), ExtFlux, YW, s, index, x2, QDotW)
				Else
					QDotW = YW(1)
				End If
				
				If WallConeDataFile(room) <> CeilingConeDataFile(room) Then
					If delta_area(room, 2, i) > 0 Then
						'ceiling
						'Call Interpolate_hrr(ByVal i - WallIgniteStep, ExtFlux(), "ceiling", X2, QDotC)
						For curve = 1 To CurveNumber_C(room)
							'step = i - WallIgniteStep(room)
							'If step > ConeNumber_C(room, curve) Then step = ConeNumber_C(room, curve)
							'YC(curve) = ConeXC(curve, step)
							elapsed = tim(i, 1) - CeilingIgniteTime(room)
							'If step > ConeNumber_C(curve) Then step = ConeNumber_C(curve)
							'YC(curve) = ConeXC(curve, step)
							If elapsed > ConeYC(room, curve, ConeNumber_C(curve, room)) Then
								YC(curve) = ConeXC(room, curve, ConeNumber_C(curve, room)) 'hrr
							Else
								Call Cone_Data(room, 2, curve, elapsed, YC(curve))
								'For T = 1 To ConeNumber_C(room, curve)
								'    If ConeYC(room, curve, T) >= elapsed Then
								'        YC(curve) = ConeXC(room, curve, T)
								'        Exit For
								'    End If
								'Next T
							End If
						Next curve
						
						ReDim ExtFlux(CurveNumber_C(room))
						For j = 1 To CurveNumber_C(room)
							ExtFlux(j) = ExtFlux_C(j, room)
						Next j
						If CurveNumber_C(room) > 1 Then
							Call CSCOEF(CurveNumber_C(room), ExtFlux, YC, s, index)
							Call CSFIT1(CurveNumber_C(room), ExtFlux, YC, s, index, x2, QDotC)
						Else
							QDotC = YC(1)
						End If
					Else
						QDotC = 0
					End If
				Else
					QDotC = QDotW
				End If
				QWall = QWall + QDotW * delta_area(room, 1, i)
				QCeiling = QCeiling + QDotC * delta_area(room, 2, i)
			End If
			i = i + 1
		Loop 
		
		If WallIgniteStep(room) > 0 And (WallConeDataFile(room) <> "" And WallConeDataFile(room) <> "null.txt") Then
			'call routine to get the qdot from the wall in the burner region = QDotWB
			If stepcount > WallIgniteStep(room) Then
				'Call Interpolate_hrr(stepcount - WallIgniteStep, ExtFlux(), "wall", NetFlux_B, QDotWB)
				For curve = 1 To CurveNumber_W(room)
					'step = i - WallIgniteStep(room)
					'If step > ConeNumber_W(room, curve) Then step = ConeNumber_W(room, curve)
					'YW(curve) = ConeXW(room, curve, step)
					elapsed = tim(stepcount, 1) - WallIgniteTime(room)
					If elapsed > ConeYW(room, curve, ConeNumber_W(curve, room)) Then
						YW(curve) = ConeXW(room, curve, ConeNumber_W(curve, room)) 'hrr
					Else
						Call Cone_Data(room, 1, curve, elapsed, YW(curve))
						'For T = 1 To ConeNumber_W(room, curve)
						'    If ConeYW(room, curve, T) >= elapsed Then
						'        YW(curve) = ConeXW(room, curve, T)
						'        Exit For
						'    End If
						'Next T
					End If
				Next curve
				'If X1 > 50 Then Stop
				
				ReDim ExtFlux(CurveNumber_W(room))
				For j = 1 To CurveNumber_W(room)
					ExtFlux(j) = ExtFlux_W(j, room)
				Next j
				ReDim s(CurveNumber_W(room))
				ReDim index(CurveNumber_W(room))
				If CurveNumber_W(room) > 1 Then
					Call CSCOEF(CurveNumber_W(room), ExtFlux, YW, s, index)
					Call CSFIT1(CurveNumber_W(room), ExtFlux, YW, s, index, x1, QDotWB)
				Else
					QDotWB = YW(1)
				End If
			Else
				QDotWB = 0
			End If
			QWall = QWall + QDotWB * delta_area(room, 1, WallIgniteStep(room))
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object hrr_linings_nextroom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hrr_linings_nextroom = QWall + QCeiling
		Exit Function
		
errorhandler: 
		
HRRerrorhandler: 
		MsgBox(ErrorToString(Err.Number))
		ERRNO = Err.Number
		'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'System.Windows.Forms.Cursor.Current = Default_Renamed
		
	End Function
	
	Public Function forced_h(ByVal surfacetemp As Single, ByVal X As Double, ByVal flameheight As Single, ByVal Flametemp As Single, ByVal i As Short, ByVal vwidth As Single, ByVal layerz As Double, ByVal UT As Double) As Object
		
        Dim denu, Re As Double
        Dim V, Pr As Single
		Dim Nu, VentAreaUpper, k As Double
		
		'get vent areas
		Call Get_Vent_Area(fireroom, X, VentAreaUpper, 0, layerz)
		
		'kinematic viscosity
		V = 7.18 * 10 ^ (-10) * ((Flametemp + surfacetemp) / 2) ^ (7 / 4)
		
		'upper layer density
		denu = ReferenceTemp * ReferenceDensity / UT
		
		If VentAreaUpper > 0 And FlowToUpper(fireroom, i - 1) < 0 Then
			Re = -FlowToUpper(fireroom, i - 1) / VentAreaUpper / denu * flameheight / V
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object forced_h. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			forced_h = 0
			Exit Function
		End If
		
		Pr = 0.72
		
		If Re < 500000# Then
            Nu = 0.332 * Sqrt(Re) * Pr ^ (1 / 3)
		Else
			Nu = 0.0296 * Re ^ (4 / 5) * Pr ^ (1 / 3)
		End If
		
		'gas conductivity
		k = 2.72 * 10 ^ (-4) * ((Flametemp + surfacetemp) / 2) ^ 0.8
		
		'UPGRADE_WARNING: Couldn't resolve default property of object forced_h. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		forced_h = Nu * k / flameheight
		
	End Function
	
	Public Sub HRR_peak(ByVal flux_wall_AR As Single, ByVal room As Integer, ByVal floorflux As Single, ByVal NetFLux_B As Single, ByVal NetFlux_F As Single, ByRef QDotW As Double, ByRef QDotC As Double, ByRef QDotWB As Double, ByRef QDotF As Double)
		'====================================================================
		' Determines the peak heat release rate from the compartment lining materials
		'
		'====================================================================
		Dim s() As Double
		Dim index() As Integer
		Dim x1 As Single
		Dim x2 As Single
		Dim x4 As Single
		Dim curve As Integer
        Dim YW(1) As Double
        Dim YC(1) As Double
		Dim x3 As Single
        Dim YF(1) As Double
		Dim ExtFlux() As Double
		Dim j As Integer
		
		If CurveNumber_W(room) > 0 Then ReDim YW(CurveNumber_W(room))
		If CurveNumber_C(room) > 0 Then ReDim YC(CurveNumber_C(room))
		If CurveNumber_F(room) > 0 Then ReDim YF(CurveNumber_F(room))
		
		If CurveNumber_W(room) > 0 Then
			If room = fireroom Then
				x1 = NetFLux_B
			Else
				x1 = flux_wall_AR
			End If
			If x1 < ExtFlux_W(1, room) Then x1 = ExtFlux_W(1, room)
			If x1 > ExtFlux_W(CurveNumber_W(room), room) Then x1 = ExtFlux_W(CurveNumber_W(room), room)
			x2 = NetFlux_F
			If x2 < ExtFlux_W(1, room) Then x2 = ExtFlux_W(1, room)
			If x2 > ExtFlux_W(CurveNumber_W(room), room) Then x2 = ExtFlux_W(CurveNumber_W(room), room)
		End If
		
		If CurveNumber_C(room) > 0 Then
			x3 = NetFlux_F
			If x3 < ExtFlux_C(1, room) Then x3 = ExtFlux_C(1, room)
			If x3 > ExtFlux_C(CurveNumber_C(room), room) Then x3 = ExtFlux_C(CurveNumber_C(room), room)
		End If
		
		If CurveNumber_F(room) > 0 Then
			x4 = floorflux
			If x4 < ExtFlux_F(1, room) Then x4 = ExtFlux_F(1, room)
			If x4 > ExtFlux_F(CurveNumber_F(room), room) Then x4 = ExtFlux_F(CurveNumber_F(room), room)
		End If
		
		'get the peak hrr for each cone input curve
		For curve = 1 To CurveNumber_W(room)
			YW(curve) = PeakWallHRR(curve, room)
		Next curve
		
		If CurveNumber_W(room) > 0 Then
			ReDim ExtFlux(CurveNumber_W(room))
			For j = 1 To CurveNumber_W(room)
				ExtFlux(j) = ExtFlux_W(j, room)
			Next j
			
			ReDim s(CurveNumber_W(room))
			ReDim index(CurveNumber_W(room))
			
			If CurveNumber_W(room) > 1 Then
				Call CSCOEF(CurveNumber_W(room), ExtFlux, YW, s, index)
				Call CSFIT1(CurveNumber_W(room), ExtFlux, YW, s, index, x2, QDotW)
			Else
				QDotW = YW(1)
			End If
			If x2 > NetFLux_B Then
				QDotW = QDotW - QDotW * (x2 - NetFlux_F) / x2
			End If
			
			If CurveNumber_W(room) > 1 Then
				Call CSCOEF(CurveNumber_W(room), ExtFlux, YW, s, index)
				Call CSFIT1(CurveNumber_W(room), ExtFlux, YW, s, index, x1, QDotWB)
			Else
				QDotWB = YW(1)
			End If
			If x1 > NetFLux_B Then
				QDotWB = QDotWB - QDotWB * (x1 - NetFLux_B) / x1
			End If
		End If
		
		For curve = 1 To CurveNumber_C(room)
			YC(curve) = PeakCeilingHRR(curve, room)
		Next curve
		
		If CurveNumber_C(room) > 0 Then
			ReDim ExtFlux(CurveNumber_C(room))
			For j = 1 To CurveNumber_C(room)
				ExtFlux(j) = ExtFlux_C(j, room)
			Next j
			
			ReDim s(CurveNumber_C(room))
			ReDim index(CurveNumber_C(room))
			If CurveNumber_C(room) > 1 Then
				Call CSCOEF(CurveNumber_C(room), ExtFlux, YC, s, index)
				Call CSFIT1(CurveNumber_C(room), ExtFlux, YC, s, index, x3, QDotC)
			Else
				QDotC = YC(1)
			End If
			If x3 > NetFlux_F Then
				QDotC = QDotC - QDotC * (x3 - NetFlux_F) / x3
			End If
		End If
		
		For curve = 1 To CurveNumber_F(room)
			YF(curve) = PeakFloorHRR(curve, room)
		Next curve
		
		
		If CurveNumber_F(room) > 0 Then
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
			If x4 > floorflux Then
				QDotF = QDotF - QDotF * (x4 - floorflux) / x4
			End If
		End If
	End Sub
	
    Function HRR_Quintiere2(ByRef id As Short, ByRef tim As Double, ByRef Q As Double, ByRef Qburner As Double, ByRef QFloor As Double, ByRef QWall As Double, ByRef QCeiling As Double) As Double
        '******************************************************
        '*  Find the total heat release rate for use by Quintiere's Room Corner Model
        '*
        '*  tim = current time
        '*  Q = hrr to use in the plume correlation kW/m2
        '*
        ' 13/12/2000 version 2001.0
        '******************************************************

        Static HRRlast, timlast, Qlast As Double
        Dim area1, hrr, Area, area2 As Double
        Dim QNormalW, burnouttime, QNormalC As Double
        Dim floorflux_f, floorflux_b As Single
        Dim QP_ceiling, FloorArea, WallArea, CeilingArea, QP_wall, QP_floor As Double
        Dim QNormalFB, QP_floor_B, QP_wall_B, QNormalWB, QNormalF As Double
        Dim qstar, deltaZ As Double
        Dim NetFlux_F, NetFLux_B As Single
        On Error GoTo HRRerrorhandler

        'if we have already called this function at this time, then we don't need to do the calculation again.
        If tim = timlast And stepcount <> 1 Then
            HRR_Quintiere2 = HRRlast
            Q = Qlast
            Exit Function
        End If

        'get the heat release rate for the burner
        'the heat release rate is entered as the first fire object by the user
        'UPGRADE_WARNING: Couldn't resolve default property of object Get_Burner(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Qburner = Get_Burner(id, tim)

        BurnerFlameLength = Get_BurnerFlameLength(id, Qburner) 'updates the global variable

        'If frmQuintiere.chkFelix.Value = vbChecked Then BurnerFlameLength = CSng(frmQuintiere.txtFlameHeight.Text)

        If tim >= WallIgniteTime(fireroom) Then
            'after ignition, the temperature of the wall next to the burner is the higher of the global wall temp
            'or the ignition temperature of the wall material
            If IgTempW(fireroom) >= Upperwalltemp(fireroom, stepcount) Then
                IgWallTemp(fireroom, stepcount) = IgTempW(fireroom)
            Else
                IgWallTemp(fireroom, stepcount) = Upperwalltemp(fireroom, stepcount)
            End If
        End If

        If tim >= CeilingIgniteTime(fireroom) Then
            If IgTempC(fireroom) >= CeilingTemp(fireroom, stepcount) Then
                IgCeilingTemp(fireroom, stepcount) = IgTempC(fireroom)
            Else
                IgCeilingTemp(fireroom, stepcount) = CeilingTemp(fireroom, stepcount)
            End If
        End If

        If tim >= FloorIgniteTime(fireroom) Then
            If IgTempF(fireroom) >= FloorTemp(fireroom, stepcount) Then
                IgFloorTemp(fireroom, stepcount) = IgTempF(fireroom)
            Else
                IgFloorTemp(fireroom, stepcount) = FloorTemp(fireroom, stepcount)
            End If
        End If

        'radiant flame flux to the wall surfaces
        If FireLocation(id) = 0 Then
            'assumes burner in centre of room
            QFB = 0 'burner flame flux
            'QPyrol = 35

            'incident heat flux from point source to nearest vertical surface on target
            QFB = ObjectRLF(id) * Qburner / (4 * PI * itemtowalldistance(id) ^ 2) 'kW/m2 first item ignited


        ElseIf FireLocation(id) = 1 Then
            ' against wall
            'assumes wall corner surfaces are in contact with burner flame
            QFB = 200 * (1 - Exp(-0.09 * Qburner ^ (1 / 3)))
        ElseIf FireLocation(id) = 2 Then
            'sfpe 3rd ed 2-272
            QFB = 120 * (1 - Exp(-4 * BurnerWidth)) '=60kw/m2 for 0.17 m burner
        End If

        QPyrol = 35
        'If frmQuintiere.chkFelix.Value = vbChecked Then
        'QFB = CDbl(frmOptions1.txtWallHeatFlux.Text)
        'QPyrol = 35
        'End If

        'determine if the wall adjacent to the burner has ignited
        If tim < WallIgniteTime(fireroom) And tim < CeilingIgniteTime(fireroom) And tim < FloorIgniteTime(fireroom) Then
            'heat release is only from the burner, no linings have ignited
            hrr = Qburner
            Q = Qburner
        Else

            'energy release rate for material (peak)
            If frmQuintiere.optUseOneCurve.Checked = True Then
                Call NewPeak(0, fireroom, QP_wall, QP_ceiling, QP_wall_B, QP_floor, QP_floor_B)
            Else
                'use all cone data and interpolate
                'Call Get_NetFlux(0, ByVal fireroom, NetFlux_F, NetFLux_B, floorflux_f, floorflux_b)
            End If
            Call Get_NetFlux(0, fireroom, NetFlux_F, NetFLux_B, floorflux_f, floorflux_b)

            If WallIgniteFlag(fireroom) = 0 Then
                'wall not yet ignited
                QP_wall_B = 0
                QP_wall = 0
            End If

            'find the pyrolysis area

            'area = Pyrol_Area(tim) 'with burnout

            'Call Pyrol_Area_NB2(id, tim, WallArea, CeilingArea, FloorArea)
            Call Pyrol_Area_NB3(id, tim, WallArea, CeilingArea, FloorArea) 'testing
            'If stepcount = 600 Then Stop
            If stepcount > 1 Then
                'keep track of the total pyrolysis area and the incremental area at each time step
                If WallEffectiveHeatofCombustion(fireroom) > 0 Then
                    pyrolarea(fireroom, 1, stepcount) = WallArea
                    'delta_area(fireroom, 1, stepcount) = pyrolarea(fireroom, 1, stepcount) - pyrolarea(fireroom, 1, stepcount - 1)
                    If pyrolarea(fireroom, 1, stepcount) > pyrolarea(fireroom, 1, stepcount - 1) Then
                        delta_area(fireroom, 1, stepcount) = pyrolarea(fireroom, 1, stepcount) - pyrolarea(fireroom, 1, stepcount - 1)
                    Else
                        delta_area(fireroom, 1, stepcount) = 0
                    End If
                End If
                If CeilingEffectiveHeatofCombustion(fireroom) > 0 Then
                    pyrolarea(fireroom, 2, stepcount) = CeilingArea
                    'delta_area(fireroom, 2, stepcount) = pyrolarea(fireroom, 2, stepcount) - pyrolarea(fireroom, 2, stepcount - 1)
                    If pyrolarea(fireroom, 2, stepcount) > pyrolarea(fireroom, 2, stepcount - 1) Then
                        delta_area(fireroom, 2, stepcount) = pyrolarea(fireroom, 2, stepcount) - pyrolarea(fireroom, 2, stepcount - 1)
                    Else
                        delta_area(fireroom, 2, stepcount) = 0
                    End If
                End If
                If FloorEffectiveHeatofCombustion(fireroom) > 0 Then
                    pyrolarea(fireroom, 3, stepcount) = FloorArea
                    'delta_area(fireroom, 3, stepcount) = pyrolarea(fireroom, 3, stepcount) - pyrolarea(fireroom, 3, stepcount - 1)
                    If pyrolarea(fireroom, 3, stepcount) > pyrolarea(fireroom, 3, stepcount - 1) Then
                        delta_area(fireroom, 3, stepcount) = pyrolarea(fireroom, 3, stepcount) - pyrolarea(fireroom, 3, stepcount - 1)
                    Else
                        delta_area(fireroom, 3, stepcount) = 0
                    End If
                End If
            End If

            If frmQuintiere.optUseOneCurve.Checked = True Then
                hrr = Qburner + hrr_dave(tim, floorflux_f, floorflux_b, QP_wall, QP_wall_B, QP_ceiling, QNormalFB, QNormalF, QNormalW, QNormalWB, QNormalC, QWall, QCeiling, QFloor)
                If flagstop = 1 Then Exit Function
            Else
                'floorflux = Target(fireroom, stepcount)
                hrr = Qburner + hrr_linings3(id, fireroom, tim, floorflux_f, floorflux_b, NetFlux_F, NetFLux_B, QFloor, QWall, QCeiling, QP_wall, QP_ceiling, QP_floor)
                'If hrr < Qburner Then Stop
            End If

            'max hrr to use in the plume correlation
            If stepcount > 1 Then
                If FireLocation(id) = 0 Or WallIgniteFlag(fireroom) = 0 Then
                    'centre of room, assume walls not burning
                    Q = Qburner
                ElseIf FireLocation(id) = 1 Then
                    'against wall
                    If frmQuintiere.optUseOneCurve.Checked = True Then
                        'Q = Qburner + BurnerWidth * (layerheight(fireroom, stepcount) - FireHeight(1)) * QP_wall * QNormalW
                        If Y_pyrolysis(fireroom, stepcount) > layerheight(fireroom, stepcount) Then
                            Q = Qburner + ((layerheight(fireroom, stepcount) - FireHeight(id)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (layerheight(fireroom, WallIgniteStep(fireroom)) - FireHeight(id)) + 0.5 * (layerheight(fireroom, stepcount) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom)))) * QP_wall * QNormalW
                        Else
                            Q = Qburner + ((Y_pyrolysis(fireroom, stepcount) - FireHeight(id)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (Y_pyrolysis(fireroom, WallIgniteStep(fireroom)) - FireHeight(id)) + 0.5 * (Y_pyrolysis(fireroom, stepcount) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom)))) * QP_wall * QNormalW
                        End If
                    Else
                        'Q = Qburner + BurnerWidth * (layerheight(fireroom, stepcount) - FireHeight(1)) * QWall
                        If Y_pyrolysis(fireroom, stepcount) > layerheight(fireroom, stepcount) Then
                            Q = Qburner + ((layerheight(fireroom, stepcount) - FireHeight(id)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (layerheight(fireroom, WallIgniteStep(fireroom)) - FireHeight(id)) + 0.5 * (layerheight(fireroom, stepcount) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom)))) * QP_wall
                        Else
                            Q = Qburner + ((Y_pyrolysis(fireroom, stepcount) - FireHeight(id)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (Y_pyrolysis(fireroom, WallIgniteStep(fireroom)) - FireHeight(id)) + 0.5 * (Y_pyrolysis(fireroom, stepcount) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom)))) * QP_wall
                        End If
                    End If
                Else
                    'corner
                    If frmQuintiere.optUseOneCurve.Checked = True Then
                        'Q = Qburner + 2 * BurnerWidth * (layerheight(fireroom, stepcount) - FireHeight(1)) * QP_wall * QNormalW
                        If Y_pyrolysis(fireroom, stepcount) > layerheight(fireroom, stepcount) Then
                            Q = Qburner + 2 * ((layerheight(fireroom, stepcount) - FireHeight(id)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (layerheight(fireroom, WallIgniteStep(fireroom)) - FireHeight(id)) + 0.5 * (layerheight(fireroom, stepcount) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom)))) * QP_wall * QNormalW
                        Else
                            Q = Qburner + 2 * ((Y_pyrolysis(fireroom, stepcount) - FireHeight(id)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (Y_pyrolysis(fireroom, WallIgniteStep(fireroom)) - FireHeight(id)) + 0.5 * (Y_pyrolysis(fireroom, stepcount) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom)))) * QP_wall * QNormalW
                        End If
                    Else
                        'Q = Qburner + 2 * BurnerWidth * (layerheight(fireroom, stepcount) - FireHeight(1)) * QWall
                        'Q = Qburner + 2 * BurnerWidth * (layerheight(fireroom, stepcount) - FireHeight(1)) * QP_wall
                        If Y_pyrolysis(fireroom, stepcount) > layerheight(fireroom, stepcount) Then
                            Q = Qburner + 2 * ((layerheight(fireroom, stepcount) - FireHeight(id)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (layerheight(fireroom, WallIgniteStep(fireroom)) - FireHeight(id)) + 0.5 * (layerheight(fireroom, stepcount) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom)))) * QP_wall
                        Else
                            Q = Qburner + 2 * ((Y_pyrolysis(fireroom, stepcount) - FireHeight(id)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (Y_pyrolysis(fireroom, WallIgniteStep(fireroom)) - FireHeight(id)) + 0.5 * (Y_pyrolysis(fireroom, stepcount) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom)))) * QP_wall
                        End If
                    End If
                End If
                Q = Q + QFloor
                If Q < Qburner + QFloor Then Q = Qburner + QFloor
            Else
                Q = Qburner + QFloor
            End If

            'Q can't be greater than the total hrr
            If Q > hrr Then Q = hrr

            'if the layer drops to the base of the fire, no entrainment
            If stepcount > 1 Then
                If layerheight(fireroom, stepcount) - FireHeight(id) < 0 Then
                    Q = 0
                End If
            End If
        End If

        'function returns the required heat release rate
        HRR_Quintiere2 = hrr
        HRR_total = hrr 'a global variable to match with the burner flame height

        'store the time and heat release values so we know what they are next time this function is called.
        timlast = tim
        HRRlast = hrr
        Qlast = Q

        Exit Function

HRRerrorhandler:
        MsgBox(ErrorToString(Err.Number) & " HRR_Quintiere2")
        ERRNO = Err.Number
    End Function
	
	Public Function RK_YBfront(ByVal fheight As Double, ByVal room As Integer, ByVal YP As Double, ByVal YB As Double, ByVal QW As Double, ByVal QC As Double) As Object
		'**********************************************************************
		'  first order diferential equation to solve for the position of the
		'  burn out front
		'   created 23/8/2001   C Wade
		'**********************************************************************
		Dim tb As Double
		tb = 0
		If CSng(YP) < RoomHeight(room) Then
			If QW > 0 Then tb = AreaUnderWallCurve(room) / QW 'burnout time applicable to wall material
		Else
			If QC > 0 Then tb = AreaUnderCeilingCurve(room) / QC 'burnout time applicable to the ceiling material
		End If
		
		If tb > 0 Then
            RK_YBfront = (YP - YB) / tb
		Else
            RK_YBfront = 0
        End If

        If RK_YBfront > 0.05 Then
            RK_YBfront = 0.05
        End If
	End Function
	
    Public Function RK_YPfront2(ByVal id As Integer, ByVal fheight As Double, ByVal room As Integer, ByVal YP As Double, ByVal YB As Single, ByVal QW As Double, ByVal QC As Double) As Double
        '*****************************************************************
        '*  Runge-Kutta function for the position of the Y pyrolysis front
        '   created 23/08/2001  C Wade
        '*****************************************************************

        Dim Qig, tig, qstar As Double
        Dim QM, YFlame, Lf As Double

        If room = fireroom Then
            'use the shorter tig of the wall or ceiling(worse case)
            tig = 100000
            If FireLocation(id) <> 0 Then
                If IgnCorrelation = vbJanssens Then
                    If WallConeDataFile(room) <> "null.txt" Then
                        tig = PI / 4 * ThermalInertiaWall(room) * ((IgTempW(room) - Upperwalltemp(room, stepcount)) / (QSpread)) ^ 2
                    End If
                    If CeilingConeDataFile(room) <> "null.txt" Then
                        If tig > PI / 4 * ThermalInertiaCeiling(room) * ((IgTempC(room) - CeilingTemp(room, stepcount)) / (QSpread)) ^ 2 Then tig = PI / 4 * ThermalInertiaCeiling(room) * ((IgTempC(room) - CeilingTemp(room, stepcount)) / (QSpread)) ^ 2
                    End If
                Else 'FTP method
                    If WallFTP(room) > 0 Then
                        If QSpread >= WallQCrit(room) Then
                            If WallConeDataFile(room) <> "null.txt" Then tig = WallFTP(room) / (QSpread - WallQCrit(room)) ^ Walln(room)
                        Else
                            'tig stays at 100000
                        End If
                    Else
                        If WallConeDataFile(room) <> "null.txt" Then tig = PI / 4 * ThermalInertiaWall(room) * ((IgTempW(room) - Upperwalltemp(room, stepcount)) / (QSpread)) ^ 2
                    End If
                    If CeilingFTP(room) > 0 Then
                        If CeilingConeDataFile(room) <> "null.txt" Then
                            If QSpread >= CeilingQCrit(room) Then
                                If tig > CeilingFTP(room) / (QSpread - CeilingQCrit(room)) ^ Ceilingn(room) Then tig = CeilingFTP(room) / (QSpread - CeilingQCrit(room)) ^ Ceilingn(room)
                            End If
                        End If
                    Else
                        If CeilingConeDataFile(room) <> "null.txt" Then
                            If tig > PI / 4 * ThermalInertiaCeiling(room) * ((IgTempC(room) - CeilingTemp(room, stepcount)) / (QSpread)) ^ 2 Then tig = PI / 4 * ThermalInertiaCeiling(room) * ((IgTempC(room) - CeilingTemp(room, stepcount)) / (QSpread)) ^ 2
                        End If
                    End If
                End If
            Else
                'fire in centre of room
                If IgnCorrelation = vbJanssens Then
                    If CeilingConeDataFile(room) <> "null.txt" Then
                        tig = PI / 4 * ThermalInertiaCeiling(room) * ((IgTempC(room) - CeilingTemp(room, stepcount)) / (QSpread)) ^ 2
                    End If
                Else 'FTP method
                    If CeilingFTP(room) > 0 Then
                        If CeilingConeDataFile(room) <> "null.txt" Then
                            If QSpread >= CeilingQCrit(room) Then
                                tig = CeilingFTP(room) / (QSpread - CeilingQCrit(room)) ^ Ceilingn(room)
                            End If
                        End If
                    Else
                        If CeilingConeDataFile(room) <> "null.txt" Then
                            tig = PI / 4 * ThermalInertiaCeiling(room) * ((IgTempC(room) - CeilingTemp(room, stepcount)) / (QSpread)) ^ 2
                        End If
                    End If
                End If
            End If

            Qig = (BurnerFlameLength / FlameAreaConstant) ^ (1 / FlameLengthPower)


            If CSng(YP) < RoomHeight(fireroom) Then
                QM = QW
            Else
                QM = QC
            End If

            'work out Yflame assuming flames are contiguous (Yflame should be the position of the flame tip relative to floor level)
            qstar = (HRR_total) / (1110 * BurnerWidth ^ (5 / 2))

            If FireLocation(id) = 0 Or FireLocation(id) = 1 Then
                'assumes burner in centre of room
                'where YB is measured relative to floor level
                'YFlame = YB + FlameAreaConstant * (QM * (YP - YB)) ^ FlameLengthPower
                Lf = -1.02 * BurnerDiameter + 0.235 * (HRR_total) ^ (2 / 5)
                If Lf + FireHeight(id) > RoomHeight(room) Then
                    'Heskestad and Hamada, pg 77 Karlsson and Quintiere book
                    YFlame = RoomHeight(room) + 0.95 * (Lf + FireHeight(id) - RoomHeight(room))
                Else
                    'YFlame = RoomHeight(room)
                    YFlame = Lf + FireHeight(id)
                End If
                If YFlame < 0 Then YFlame = 0
                'ElseIf FireLocation(1) = 1 Then
                'for wall configuration
                'YFlame = BurnerWidth * 3 * (qstar) ^ (2 / 3)
            Else
                'for corner configuration
                'kokkala/heskestad use - when flame height is below ceiling
                'YFlame = FireHeight(1) + BurnerWidth * (-2.04 + 6.62 * (qstar) ^ (2 / 5))

                '10june2003
                YFlame = FireHeight(id) + BurnerWidth * 5.9 * Sqrt(qstar)

                If YFlame > RoomHeight(fireroom) Then
                    'Use Karlsson-Thomas correlation for flame extension beneath ceiling
                    qstar = (HRR_total) / (ReferenceDensity * InteriorTemp * SpecificHeat_air * (RoomHeight(fireroom) - FireHeight(id) + 3 * BurnerWidth) ^ 2 * Sqrt(G * (RoomHeight(fireroom) - FireHeight(id) + 3 * BurnerWidth)))
                    'flame extension beneath ceiling
                    Lf = (25 * qstar - 0.15) * (RoomHeight(fireroom) - FireHeight(id) + 3 * BurnerWidth)
                    YFlame = Lf + RoomHeight(fireroom) 'distance of flame tip from the floor
                End If


                If YFlame < 0 Then YFlame = 0
            End If

            If YB - FireHeight(id) < BurnerFlameLength Then
                'If YB < YFlame Then
                'burner and wall flame are contiguous

                '    If FireLocation(1) = 0 Then
                'assumes burner in centre of room
                '        YFlame = YB + FlameAreaConstant * (QM * (YP - YB)) ^ FlameLengthPower

                '        If YFlame < 0 Then YFlame = 0
                '    ElseIf FireLocation(1) = 1 Then
                'for wall configuration
                'YFlame = BurnerDiameter * 3 * (HeatRelease(room, stepcount, 2) / (1110 * BurnerDiameter ^ (5 / 2))) ^ (2 / 3)
                'YFlame = BurnerDiameter * 3 * (HRR_total / (1110 * BurnerDiameter ^ (5 / 2))) ^ (2 / 3)
                '        YFlame = BurnerDiameter * 3 * (HRR_total / (1110 * BurnerWidth ^ (5 / 2))) ^ (2 / 3)
                '    Else
                'for corner configuration
                'qstar = (HeatRelease(room, stepcount, 2)) / (1110 * BurnerDiameter ^ (5 / 2))
                'qstar = (HRR_total) / (1110 * BurnerDiameter ^ (5 / 2))
                '        qstar = (HRR_total) / (1110 * BurnerWidth ^ (5 / 2))
                '        If qstar < 0 Then qstar = 0
                'kokkala/heskestad use - when flame height is below ceiling
                '        YFlame = FireHeight(1) + BurnerWidth * (-2.04 + 6.62 * (qstar) ^ (2 / 5))
                'YFlame = YB + BurnerWidth * (-2.04 + 6.62 * (qstar) ^ (2 / 5))
                '        If YFlame > RoomHeight(fireroom) Then
                'qstar = (HeatRelease(room, stepcount, 2)) / (ReferenceDensity * InteriorTemp * SpecificHeat_air * (RoomHeight(fireroom) - FireHeight(1) + 3 * BurnerWidth) ^ 2 * Sqr(G * (RoomHeight(fireroom) - FireHeight(1) + 3 * BurnerWidth)))
                '            qstar = (HRR_total) / (ReferenceDensity * InteriorTemp * SpecificHeat_air * (RoomHeight(fireroom) - FireHeight(1) + 3 * BurnerWidth) ^ 2 * Sqr(G * (RoomHeight(fireroom) - FireHeight(1) + 3 * BurnerWidth)))

                '            YFlame = (25 * qstar - 0.15) * (RoomHeight(fireroom) - FireHeight(1) + 3 * BurnerWidth) + RoomHeight(fireroom)
                '        End If
                '         If YFlame < 0 Then YFlame = 0
                'If YFlame > RoomHeight(fireroom) + RoomLength(fireroom) + RoomWidth(fireroom) Then Stop
                '     End If
                'Debug.Print stepcount; YP; YB; YFlame
                'YFlame = FireHeight(1) + FlameAreaConstant * (Qig + QM * (YP - YB)) ^ FlameLengthPower

            Else
                'burner and wall flame are separated
                'qstar = (HeatRelease(room, stepcount, 2) - Qburner) / (1110 * BurnerDiameter ^ (5 / 2))
                'If qstar < 0 Then qstar = 0
                'YFlame = YB + BurnerWidth * (-2.04 + 6.62 * (qstar) ^ (2 / 5))

                If YP > YB Then
                    YFlame = YB + FlameAreaConstant * (QM * (YP - YB)) ^ FlameLengthPower
                    'YFlame = YB + 0.043 * (QM * (YP - YB)) ^ FlameLengthPower - 0.27 'from kokkala, baroudi and parker 5th IAFSS
                    'YFlame = YB + 0.0065 * (QM * (YP - YB)) 'from kokkala, baroudi and parker 5th IAFSS
                Else
                    YFlame = YB
                End If
            End If

            If tig > 0 Then
                If YFlame > 2 * RoomHeight(room) + RoomLength(room) + RoomWidth(room) Then YFlame = 2 * RoomHeight(room) + RoomLength(room) + RoomWidth(room)
                RK_YPfront2 = (YFlame - YP) / tig 'm/s


                'Exit Function
            Else
                RK_YPfront2 = 0.1
            End If

        Else
            'not the fireroom
            tig = 100000
            If IgnCorrelation = vbJanssens Then
                If WallConeDataFile(room) <> "null.txt" Then tig = PI / 4 * ThermalInertiaWall(room) * ((IgTempW(room) - Upperwalltemp(room, stepcount)) / (QSpread)) ^ 2
                If CeilingConeDataFile(room) <> "null.txt" Then
                    If tig > PI / 4 * ThermalInertiaCeiling(room) * ((IgTempC(room) - CeilingTemp(room, stepcount)) / (QSpread)) ^ 2 Then tig = PI / 4 * ThermalInertiaCeiling(room) * ((IgTempC(room) - CeilingTemp(room, stepcount)) / (QSpread)) ^ 2
                End If
            Else
                If WallConeDataFile(room) <> "null.txt" Then tig = WallFTP(room) / (QSpread - WallQCrit(room)) ^ Walln(room)
                'If WallConeDataFile(room) <> "null.txt" Then tig = WallFTP(room) / (QSpread) ^ Walln(room)
                If CeilingConeDataFile(room) <> "null.txt" Then
                    If tig > CeilingFTP(room) / (QSpread - CeilingQCrit(room)) ^ Ceilingn(room) Then tig = CeilingFTP(room) / (QSpread - CeilingQCrit(room)) ^ Ceilingn(room)
                    'If tig > CeilingFTP(room) / (QSpread) ^ Ceilingn(room) Then tig = CeilingFTP(room) / (QSpread) ^ Ceilingn(room)
                End If
            End If

            'estimate the flame height of a vent fire using a hasemi wall correlation
            'qstar = ventfire(room, stepcount) / (1110 * VentWidth(fireroom, room, 1) ^ (5 / 2))
            'flameheight = 2.8 * VentWidth(fireroom, room, 1) * (qstar / VentWidth(fireroom, room, 1)) ^ (2 / 3)
            Qig = (fheight / FlameAreaConstant) ^ (1 / FlameLengthPower)

            If WallIgniteStep(room) > 1 Then
                If QW > QC Then
                    YFlame = FlameAreaConstant * (Qig + QW * (YP - Y_pyrolysis(room, WallIgniteStep(room) - 1))) ^ FlameLengthPower
                Else
                    YFlame = FlameAreaConstant * (Qig + QC * (YP - Y_pyrolysis(room, WallIgniteStep(room) - 1))) ^ FlameLengthPower
                End If
                If tig > 0 Then
                    RK_YPfront2 = (YFlame - (YP - Y_pyrolysis(room, WallIgniteStep(room) - 1))) / tig
                    'Exit Function
                Else
                    RK_YPfront2 = 1000 'very fast
                End If
            Else
                'only ceiling is burning
                If CeilingIgniteStep(room) > 1 Then
                    YFlame = FlameAreaConstant * (Qig + QC * (YP - Y_pyrolysis(room, CeilingIgniteStep(room) - 1))) ^ FlameLengthPower
                    If tig > 0 Then
                        RK_YPfront2 = (YFlame - (YP - Y_pyrolysis(room, CeilingIgniteStep(room) - 1))) / tig
                        'Exit Function
                    Else
                        RK_YPfront2 = 1000 'very fast
                    End If
                End If
            End If
        End If

        'RK_YPfront2 = 0
        If RK_YPfront2 > 0.1 Then RK_YPfront2 = 0.1

    End Function
	
    Public Function Get_BurnerFlameLength(ByRef id As Short, ByRef Qburner As Single) As Single
        '************************************************************************************
        '* calculate the burner flame height
        '*  created 26/8/2001  C A Wade
        '************************************************************************************
        On Error GoTo errorhandler

        Dim qstar As Double

        'non dimensional HRR
        qstar = Qburner / (1110 * BurnerWidth ^ (5 / 2))

        If FireLocation(id) = 0 Then
            'assumes burner in centre of room - heskastad
            Get_BurnerFlameLength = -1.02 * BurnerDiameter + 0.235 * (Qburner) ^ (2 / 5)
            If BurnerFlameLength < 0 Then Get_BurnerFlameLength = 0
        ElseIf FireLocation(id) = 1 Then
            'for wall configuration
            'Hasemi and Tokunaga
            'Get_BurnerFlameLength = BurnerWidth * 2.8 * (qstar / BurnerWidth) ^ (2 / 3)

            'use the open air correlation
            Get_BurnerFlameLength = -1.02 * BurnerDiameter + 0.235 * (Qburner) ^ (2 / 5)
            If BurnerFlameLength < 0 Then Get_BurnerFlameLength = 0
        Else
            'for corner configuration
            'kokkala/heskestad use - when flame height is below ceiling for a corner configuration
            '     Get_BurnerFlameLength = BurnerWidth * (-2.04 + 6.62 * qstar ^ (2 / 5))

            ' If Get_BurnerFlameLength < 0 Then
            '     Get_BurnerFlameLength = 0.01
            ' ElseIf Get_BurnerFlameLength + FireHeight(1) > RoomHeight(fireroom) Then
            '     'flame extension beneath ceiling, use thomas and karlsson
            '     qstar = Qburner / (ReferenceDensity * InteriorTemp * SpecificHeat_air * (RoomHeight(fireroom) - FireHeight(1) + 3 * BurnerWidth) ^ 2 * Sqr(G * (RoomHeight(fireroom) - FireHeight(1) + 3 * BurnerWidth)))
            '     Get_BurnerFlameLength = (25 * qstar - 0.15) * (RoomHeight(fireroom) - FireHeight(1) + 3 * BurnerWidth) + RoomHeight(fireroom) - FireHeight(1)
            ' Else
            '
            ' End If

            'sfpe h/b 3rd ed 2-272
            qstar = Qburner / (1110 * BurnerWidth ^ (5 / 2))
            Get_BurnerFlameLength = BurnerWidth * 5.9 * Sqrt(qstar)
        End If
        Exit Function

errorhandler:
        MsgBox(ErrorToString(Err.Number) & " Function Get_BurnerFlameLength in Module QUINT")
        ERRNO = Err.Number
    End Function
	
	Public Function hrr_linings2(ByVal room As Integer, ByVal current_time As Single, ByVal floorflux_f As Single, ByVal floorflux_b As Single, ByVal NetFlux_F As Single, ByVal NetFLux_B As Single, ByRef QFloor As Single, ByRef QWall As Single, ByRef QCeiling As Single, ByRef QDotW As Double, ByRef QDotC As Double, ByRef QDotF As Double) As Object
		'====================================================================
		' Determines the heat release rate from the compartment lining materials
		' by storing the change in pyrolysis area for each time step and
		' using the elapsed time to determine the heat release rate for each
		' burning segment from the cone calorimeter input
		' 19 April 2002 C Wade
		'====================================================================
		
		Dim i, j As Integer
		Dim hrr As Double
		Dim s() As Double
		Dim index() As Integer
		Dim x4, x1, x2, x5 As Single
		Dim burnarea As Double
		Dim curve As Integer
        Dim YW(1) As Double
        Dim YC(1) As Double
        Dim YF(1) As Double
		Dim T, x3, elapsed As Single
		'UPGRADE_NOTE: step was upgraded to step_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim QDotWB, QDotFB As Double
		Dim step_Renamed, lastpoint As Integer
		Dim ExtFlux() As Double
		
		On Error GoTo errorhandler
		
		If CurveNumber_W(room) > 0 Then ReDim YW(CurveNumber_W(room))
		If CurveNumber_C(room) > 0 Then ReDim YC(CurveNumber_C(room))
		If CurveNumber_F(room) > 0 Then ReDim YF(CurveNumber_F(room))
		
		hrr = 0
		QWall = 0
		QCeiling = 0
		QFloor = 0
		
		'the external heat flux has to be within the range of the cone data available
		If CurveNumber_W(room) > 0 Then
			x1 = NetFLux_B
			If x1 < ExtFlux_W(1, room) Then x1 = ExtFlux_W(1, room)
			If x1 > ExtFlux_W(CurveNumber_W(room), room) Then x1 = ExtFlux_W(CurveNumber_W(room), room)
			x2 = NetFlux_F
			If x2 < ExtFlux_W(1, room) Then x2 = ExtFlux_W(1, room)
			If x2 > ExtFlux_W(CurveNumber_W(room), room) Then x2 = ExtFlux_W(CurveNumber_W(room), room)
		End If
		If CurveNumber_C(room) > 0 Then
			x3 = NetFlux_F
			If x3 < ExtFlux_C(1, room) Then x3 = ExtFlux_C(1, room)
			If x3 > ExtFlux_C(CurveNumber_C(room), room) Then x3 = ExtFlux_C(CurveNumber_C(room), room)
		End If
		If CurveNumber_F(room) > 0 Then
			x4 = floorflux_f 'ahead of the burner flame
			If x4 < ExtFlux_F(1, room) Then x4 = ExtFlux_F(1, room)
			If x4 > ExtFlux_F(CurveNumber_F(room), room) Then x4 = ExtFlux_F(CurveNumber_F(room), room)
			x5 = floorflux_b 'in the region of the burner flame
			If x5 < ExtFlux_F(1, room) Then x5 = ExtFlux_F(1, room)
			If x5 > ExtFlux_F(CurveNumber_F(room), room) Then x5 = ExtFlux_F(CurveNumber_F(room), room)
		End If
		
		'======================================================
		'the wall next to the burner flamer, area first ignited
		'======================================================
		If WallIgniteStep(room) > 0 And (WallConeDataFile(room) <> "" And WallConeDataFile(room) <> "null.txt") Then
			'call routine to get the qdot from the wall in the burner region = QDotWB
			If current_time > WallIgniteTime(room) Then
				elapsed = current_time - WallIgniteTime(room)
				For curve = 1 To CurveNumber_W(room)
					If elapsed > ConeYW(room, curve, ConeNumber_W(curve, room)) Then
						YW(curve) = ConeXW(room, curve, ConeNumber_W(curve, room)) 'hrr
					Else
						Call Cone_Data(room, 1, curve, elapsed, YW(curve))
					End If
				Next curve
				
				ReDim ExtFlux(CurveNumber_W(room))
				For j = 1 To CurveNumber_W(room)
					ExtFlux(j) = ExtFlux_W(j, room)
				Next j
				
				ReDim s(CurveNumber_W(room))
				ReDim index(CurveNumber_W(room))
				If CurveNumber_W(room) > 1 Then
					Call CSCOEF(CurveNumber_W(room), ExtFlux, YW, s, index)
					Call CSFIT1(CurveNumber_W(room), ExtFlux, YW, s, index, x1, QDotWB)
				Else
					QDotWB = YW(1)
				End If
			Else
				QDotWB = 0
			End If
			
			If x1 = ExtFlux_W(CurveNumber_W(room), room) Then
				'extrapolate
				QDotWB = QDotWB * NetFLux_B / x1
			End If
			
			QWall = QWall + QDotWB * delta_area(room, 1, WallIgniteStep(room))
		End If
		
		'======================================================
		'the rest of the wall linings
		'======================================================
		If WallIgniteFlag(room) > 0 Then 'wall has ignited
			i = WallIgniteStep(room) + 1 'exclude the area first ignited
			Do While i < stepcount
				elapsed = current_time - tim(i, 1) 'section i started burning at tim(i,1) and tim(stepcount,1) is the curent time
				For curve = 1 To CurveNumber_W(room)
					
					lastpoint = ConeNumber_W(curve, room)
					Do While ConeYW(room, curve, lastpoint) = 0
						lastpoint = lastpoint - 1
					Loop 
					
					If elapsed > ConeYW(room, curve, lastpoint) Then
						YW(curve) = ConeXW(room, curve, lastpoint) 'hrr
					Else
						Call Cone_Data(room, 1, curve, elapsed, YW(curve))
					End If
				Next curve
				
				ReDim ExtFlux(CurveNumber_W(room))
				For j = 1 To CurveNumber_W(room)
					ExtFlux(j) = ExtFlux_W(j, room)
				Next j
				
				ReDim s(CurveNumber_W(room))
				ReDim index(CurveNumber_W(room))
				If CurveNumber_W(room) > 1 Then
					Call CSCOEF(CurveNumber_W(room), ExtFlux, YW, s, index)
					Call CSFIT1(CurveNumber_W(room), ExtFlux, YW, s, index, x2, QDotW)
				Else
					QDotW = YW(1)
				End If
				
				If x2 = ExtFlux_W(CurveNumber_W(room), room) Then
					'extrapolate
					QDotW = QDotW * NetFlux_F / x2
				End If
				
				'Debug.Print current_time, elapsed, QDotW, delta_area(room, 1, i)
				QWall = QWall + QDotW * delta_area(room, 1, i) 'excludes the area first ignited by the burner
				i = i + 1
			Loop 
		End If
		
		'======================================================
		'ceiling linings
		'======================================================
		If CeilingIgniteFlag(room) > 0 Then 'ceiling has ignited
			i = CeilingIgniteStep(room)
			Do While i < stepcount
				elapsed = current_time - tim(i, 1) 'section i started burning at tim(i,1) and tim(stepcount,1) is the curent time
				For curve = 1 To CurveNumber_C(room)
					
					lastpoint = ConeNumber_C(curve, room)
					Do While ConeYC(room, curve, lastpoint) = 0
						lastpoint = lastpoint - 1
					Loop 
					
					If elapsed > ConeYC(room, curve, lastpoint) Then
						YC(curve) = ConeXC(room, curve, lastpoint) 'hrr
					Else
						Call Cone_Data(room, 2, curve, elapsed, YC(curve))
					End If
				Next curve
				
				ReDim ExtFlux(CurveNumber_C(room))
				For j = 1 To CurveNumber_C(room)
					ExtFlux(j) = ExtFlux_C(j, room)
				Next j
				
				ReDim s(CurveNumber_C(room))
				ReDim index(CurveNumber_C(room))
				If CurveNumber_C(room) > 1 Then
					Call CSCOEF(CurveNumber_C(room), ExtFlux, YC, s, index)
					Call CSFIT1(CurveNumber_C(room), ExtFlux, YC, s, index, x3, QDotC)
				Else
					QDotC = YC(1)
				End If
				'Debug.Print current_time, elapsed, QDotC, delta_area(room, 2, i)
				If x3 = ExtFlux_C(CurveNumber_C(room), room) Then
					'extrapolate
					QDotC = QDotC * NetFlux_F / x3
				End If
				QCeiling = QCeiling + QDotC * delta_area(room, 2, i)
				i = i + 1
			Loop 
		End If
		
		'======================================================
		'the floor next to the burner flamer, area first ignited
		'======================================================
		If FloorIgniteStep(room) > 0 And (FloorConeDataFile(room) <> "" And FloorConeDataFile(room) <> "null.txt") Then
			'call routine to get the qdot from the floor in the burner region = QDotFB
			If current_time > FloorIgniteTime(room) Then
				elapsed = current_time - FloorIgniteTime(room)
				For curve = 1 To CurveNumber_F(room)
					If elapsed > ConeYF(room, curve, ConeNumber_F(curve, room)) Then
						YF(curve) = ConeXF(room, curve, ConeNumber_F(curve, room)) 'hrr
					Else
						Call Cone_Data(room, 3, curve, elapsed, YF(curve))
					End If
				Next curve
				
				ReDim ExtFlux(CurveNumber_F(room))
				For j = 1 To CurveNumber_F(room)
					ExtFlux(j) = ExtFlux_F(j, room)
				Next j
				
				ReDim s(CurveNumber_F(room))
				ReDim index(CurveNumber_F(room))
				If CurveNumber_F(room) > 1 Then
					Call CSCOEF(CurveNumber_F(room), ExtFlux, YF, s, index)
					Call CSFIT1(CurveNumber_F(room), ExtFlux, YF, s, index, x5, QDotFB)
				Else
					QDotFB = YF(1)
				End If
			Else
				QDotFB = 0
			End If
			QFloor = QFloor + QDotFB * delta_area(room, 3, FloorIgniteStep(room))
		End If
		
		'======================================================
		'rest of the floor linings
		'======================================================
		If FloorIgniteFlag(room) > 0 Then 'floor has ignited
			i = FloorIgniteStep(room) + 1
			Do While i < stepcount
				elapsed = current_time - tim(i, 1) 'section i started burning at tim(i,1) and tim(stepcount,1) is the curent time
				For curve = 1 To CurveNumber_F(room)
					
					lastpoint = ConeNumber_F(curve, room)
					Do While ConeYF(room, curve, lastpoint) = 0
						lastpoint = lastpoint - 1
					Loop 
					
					If elapsed > ConeYF(room, curve, lastpoint) Then
						YF(curve) = ConeXF(room, curve, lastpoint) 'hrr
					Else
						Call Cone_Data(room, 3, curve, elapsed, YF(curve))
					End If
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
				'Debug.Print current_time, elapsed, QDotF, delta_area(room, 3, i)
				QFloor = QFloor + QDotF * delta_area(room, 3, i)
				'If delta_area(room, 3, i) < 0 Then Stop
				i = i + 1
			Loop 
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object hrr_linings2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hrr_linings2 = QWall + QCeiling + QFloor 'kw
		Erase YW
		Erase YC
		Erase YF
		Erase ExtFlux
		
		Exit Function
		
errorhandler: 
		MsgBox(ErrorToString(Err.Number) & " hrr_linings2")
		ERRNO = Err.Number
    End Function
	
    Private Sub Pyrol_Area_NB2(ByRef id As Short, ByRef T As Double, ByRef WallArea As Double, ByRef CeilingArea As Double, ByRef FloorArea As Double)
        '=================================================================
        '=  This function calculates the wall pyrolysis area not including
        '=  burnout following Quintiere's description of the areas
        '=
        '=  Function called by HRR_Quintiere
        '=
        '=  Revised 14 August 1998 Colleen Wade
        '=================================================================

        Dim z_pyrol, jetdepth As Double
        Dim APJ1, AP1, APC1 As Double

        On Error GoTo AREAerrorhandler

        jetdepth = 0.12 'depth of ceiling jet 12% of fire-ceiling height

        If T >= WallIgniteTime(fireroom) Then

            'the wall has ignited
            If Y_pyrolysis(fireroom, stepcount) < RoomHeight(fireroom) Then
                'Case 1a
                'The ignitor burner has caused ignition of wall
                'The pyrolysis height is less than the ceiling height
                If FireLocation(id) = 2 Then
                    'corner
                    WallArea = 2 * ((Y_pyrolysis(fireroom, stepcount) - FireHeight(id)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (Y_pyrolysis(fireroom, WallIgniteStep(fireroom)) - FireHeight(id)) + 0.5 * (Y_pyrolysis(fireroom, stepcount) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))))
                ElseIf FireLocation(id) = 1 Then
                    'wall
                    'WallArea = ((Y_pyrolysis(fireroom, stepcount) - FireHeight(1)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (Y_pyrolysis(fireroom, WallIgniteStep(fireroom)) - FireHeight(1)) + 0.5 * (Y_pyrolysis(fireroom, stepcount) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))))
                    'ver2002.2
                    WallArea = (Y_pyrolysis(fireroom, stepcount) - FireHeight(id)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + 2 * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (Y_pyrolysis(fireroom, WallIgniteStep(fireroom)) - FireHeight(id)) + (Y_pyrolysis(fireroom, stepcount) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom)))
                Else
                    'centre of room
                    WallArea = 0

                    'as for wall location
                    WallArea = (Y_pyrolysis(fireroom, stepcount) - FireHeight(id)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + 2 * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (Y_pyrolysis(fireroom, WallIgniteStep(fireroom)) - FireHeight(id)) + (Y_pyrolysis(fireroom, stepcount) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom)))

                End If
                CeilingArea = 0

                'Exit Sub
            Else
                'Case 1b
                'The ignitor burner has caused ignition of wall
                'The pyrolysis height is greater than the ceiling height
                If CeilingIgniteFlag(fireroom) = 0 And CeilingEffectiveHeatofCombustion(fireroom) > 0 Then '17/11/2011
                    CeilingIgniteFlag(fireroom) = 1
                    'MDIFrmMain.ToolStripStatusLabel3.Text = "Ceiling in Room " & CStr(fireroom) & " has ignited at " & CStr(T) & " seconds."

                    Dim Message As String = "Ceiling in Room " & CStr(fireroom) & " has ignited at " & CStr(T) & " seconds (due to wall burning)."
                    MDIFrmMain.ToolStripStatusLabel2.Text = Message.ToString
                    If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                    CeilingIgniteTime(fireroom) = T
                    CeilingIgniteStep(fireroom) = stepcount
                End If

                If FireLocation(id) = 2 Then
                    AP1 = 2 * ((RoomHeight(fireroom) - FireHeight(id)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (Y_pyrolysis(fireroom, WallIgniteStep(fireroom)) - FireHeight(id)) + 0.5 * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (RoomHeight(fireroom) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom))))
                ElseIf FireLocation(id) = 1 Then
                    'AP1 = ((RoomHeight(fireroom) - FireHeight(1)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (Y_pyrolysis(fireroom, WallIgniteStep(fireroom)) - FireHeight(1)) + 0.5 * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (RoomHeight(fireroom) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom))))
                    'v2002.2
                    AP1 = 2 * ((RoomHeight(fireroom) - FireHeight(id)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (Y_pyrolysis(fireroom, WallIgniteStep(fireroom)) - FireHeight(id)) + 0.5 * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (RoomHeight(fireroom) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom)))) - (RoomHeight(fireroom) - FireHeight(id)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom))
                Else
                    'centre of room
                    AP1 = 0

                    'as for wall, 17/11/2011
                    AP1 = 2 * ((RoomHeight(fireroom) - FireHeight(id)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom)) + (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (Y_pyrolysis(fireroom, WallIgniteStep(fireroom)) - FireHeight(id)) + 0.5 * (X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) * (RoomHeight(fireroom) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom)))) - (RoomHeight(fireroom) - FireHeight(id)) * X_pyrolysis(fireroom, WallIgniteStep(fireroom))

                End If

                'The z_pyrolysis length cannot be greater than (1-0.12) * room height
                'z_pyrol = (1 - jetdepth) * (RoomHeight(fireroom) - FireHeight(1))
                z_pyrol = (1 - jetdepth) * (RoomHeight(fireroom))
                If z_pyrol > X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom)) Then
                    z_pyrol = X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))
                End If
                If T = tim(stepcount, 1) Then
                    Z_pyrolysis(fireroom, stepcount) = z_pyrol
                End If

                If z_pyrol > 0 Then
                    'corner
                    APJ1 = 2 * ((Y_pyrolysis(fireroom, stepcount) - RoomHeight(fireroom)) * (jetdepth * (RoomHeight(fireroom) - FireHeight(id))) + 0.5 * z_pyrol * (Y_pyrolysis(fireroom, stepcount) - RoomHeight(fireroom)) - 0.5 * (jetdepth * (RoomHeight(fireroom) - FireHeight(id)) + z_pyrol) ^ 2 * ((X_pyrolysis(fireroom, stepcount) - X_pyrolysis(fireroom, WallIgniteStep(fireroom))) / (RoomHeight(fireroom) - Y_pyrolysis(fireroom, WallIgniteStep(fireroom)))))
                Else
                    'corner
                    APJ1 = 2 * (Y_pyrolysis(fireroom, stepcount) - RoomHeight(fireroom)) * (jetdepth * (RoomHeight(fireroom) - FireHeight(id)))
                End If
                If APJ1 < 0 Then APJ1 = 0

                WallArea = AP1 + APJ1
                If WallArea > 2 * (RoomHeight(fireroom) * RoomLength(fireroom) + RoomHeight(fireroom) * RoomWidth(fireroom)) Then
                    WallArea = 2 * (RoomHeight(fireroom) * RoomLength(fireroom) + RoomHeight(fireroom) * RoomWidth(fireroom))
                End If
                CeilingArea = PI / 4 * (Y_pyrolysis(fireroom, stepcount) - RoomHeight(fireroom)) ^ 2
                If CeilingArea > RoomFloorArea(fireroom) Then CeilingArea = RoomFloorArea(fireroom)
            End If
        Else
            WallArea = 0
            CeilingArea = 0
        End If

        If T >= CeilingIgniteTime(fireroom) Then
            'the ceiling has ignited
            If FireLocation(id) = 0 Then
                'centre of room
                If CeilingArea < PI * (Y_pyrolysis(fireroom, stepcount) - RoomHeight(fireroom)) ^ 2 Then
                    CeilingArea = PI * (Y_pyrolysis(fireroom, stepcount) - RoomHeight(fireroom)) ^ 2
                End If
            ElseIf FireLocation(id) = 2 Then
                'corner location
                If CeilingArea < PI / 4 * (Y_pyrolysis(fireroom, stepcount) - RoomHeight(fireroom)) ^ 2 Then
                    CeilingArea = PI / 4 * (Y_pyrolysis(fireroom, stepcount) - RoomHeight(fireroom)) ^ 2
                End If
            Else
                'wall location
                If CeilingArea < PI / 2 * (Y_pyrolysis(fireroom, stepcount) - RoomHeight(fireroom)) ^ 2 Then
                    CeilingArea = PI / 2 * (Y_pyrolysis(fireroom, stepcount) - RoomHeight(fireroom)) ^ 2
                End If
            End If
            If CeilingArea > RoomFloorArea(fireroom) Then CeilingArea = RoomFloorArea(fireroom)
        End If

        '=======================
        ' floor
        '=======================
        Dim theta As Single
        theta = 60

        If T >= FloorIgniteTime(fireroom) Then
            'the floor has ignited
            'If frmQuintiere.chkFloorFlameSpread.Value = vbChecked Then
            If frmQuintiere.optWindAided.Checked = True Then
                'this is when we have wind-aided spread
                'the flame spread is assumed to be radially outside as a cone with angle of 60 deg
                'until wall is reached
                If FireLocation(id) = 0 Then
                    'centre of room
                    'wind-aided spread
                    If Yf_pyrolysis(fireroom, stepcount) < RoomWidth(fireroom) / (2 * Cos(PI / 180 * (90 - theta / 2))) Then
                        'pyrolysis front not reached side wall
                        FloorArea = theta / 360 * PI * Yf_pyrolysis(fireroom, stepcount) ^ 2
                    Else
                        FloorArea = PI * Yf_pyrolysis(fireroom, stepcount) ^ 2 / 2 - 2 * ((90 - theta / 2) * PI * Yf_pyrolysis(fireroom, stepcount) ^ 2 / 360 - RoomWidth(fireroom) / 4 * RoomWidth(fireroom) / (2 * Cos(PI / 180 * (90 - theta / 2))))
                        FloorArea = FloorArea - (RoomWidth(fireroom) / (2 * Cos(90 - theta / 2)) * RoomWidth(fireroom) / 2)
                    End If
                    'limit max area to 1/2 the room floor area
                    'If FloorArea > 0.5 * RoomFloorArea(fireroom) - (RoomWidth(fireroom) / (2 * Cos(PI / 180 * (90 - theta / 2))) * RoomWidth(fireroom) / 2) Then
                    '    FloorArea = 0.5 * RoomFloorArea(fireroom) - (RoomWidth(fireroom) / (2 * Cos(PI / 180 * (90 - theta / 2))) * RoomWidth(fireroom) / 2)
                    'End If
                ElseIf FireLocation(id) = 2 Then
                    'corner location
                    FloorArea = 0 'no wind-aided spread from a corner location
                Else
                    'wall location
                    'wind-aided spread + opposed flow
                    FloorArea = Yf_pyrolysis(fireroom, stepcount) * BurnerWidth + 2 * (Yf_pyrolysis(fireroom, FloorIgniteStep(fireroom)) * (Xf_pyrolysis(fireroom, stepcount) - Xf_pyrolysis(fireroom, FloorIgniteTime(fireroom))) + 0.5 * (Xf_pyrolysis(fireroom, stepcount) - Xf_pyrolysis(fireroom, FloorIgniteStep(fireroom))) * (Yf_pyrolysis(fireroom, stepcount) - Yf_pyrolysis(fireroom, FloorIgniteStep(fireroom))))
                    'limit max area to the room floor area
                    If FloorArea > RoomFloorArea(fireroom) Then FloorArea = RoomFloorArea(fireroom)
                End If
            ElseIf frmQuintiere.optOpposedFlow.Checked = True Then
                'opposed flow spread
                If FireLocation(id) = 0 Then
                    'centre of room
                    FloorArea = PI * Xf_pyrolysis(fireroom, stepcount) ^ 2 'area of a circle
                ElseIf FireLocation(id) = 2 Then
                    'corner location
                    FloorArea = PI / 4 * Xf_pyrolysis(fireroom, stepcount) ^ 2 'area of a quarter-circle
                Else
                    'against a wall
                    FloorArea = PI / 2 * Xf_pyrolysis(fireroom, stepcount) ^ 2 'area of a semi-circle
                End If

                'limit max area to the room floor area
                If FloorArea > RoomFloorArea(fireroom) Then FloorArea = RoomFloorArea(fireroom)

            Else
                'the whole room floor area has ignited, assuming uniform radiation from a hot layer
                FloorArea = RoomFloorArea(fireroom)
            End If
        End If

       

        Exit Sub

AREAerrorhandler:
        MsgBox(ErrorToString(Err.Number) & "in Pyrol_Area_NB2")
        ERRNO = Err.Number

    End Sub

    Private Sub Pyrol_Area_NB3(ByRef id As Short, ByRef T As Double, ByRef WallArea As Double, ByRef CeilingArea As Double, ByRef FloorArea As Double)
        '=================================================================
        '=  This function calculates the wall and ceiling pyrolysis area not including
        '=  burnout following Quintiere's description of the areas
        '=
        '=  Function called by HRR_Quintiere
        '=================================================================

        ''hardcode limit values for testing 
        'Dim x_limit As Double = 30 'm horizontal wall limit measured from burner
        'Dim y_limit As Double = 2.4 'm vertical wall limit measured from floor
        Dim y_pyrol_wall As Double = Y_pyrolysis(fireroom, stepcount)
        Dim x_pyrol As Double = X_pyrolysis(fireroom, stepcount)
        Dim y_pyrol As Double = Y_pyrolysis(fireroom, stepcount)
        Dim RH As Double = RoomHeight(fireroom)
        Dim y_pyrol_ign As Double = Y_pyrolysis(fireroom, WallIgniteStep(fireroom))
        Dim x_pyrol_ign As Double = X_pyrolysis(fireroom, WallIgniteStep(fireroom))

        If PessimiseCombWall = False Then
            If wallVFSlimit < RH Then
                If y_pyrol_wall > wallVFSlimit Then
                    y_pyrol_wall = wallVFSlimit 'constrain
                End If
            End If
            If x_pyrol > wallHFSlimit Then
                x_pyrol = wallHFSlimit 'constrain
            End If
        End If

        Dim z_pyrol, jetdepth As Double
        Dim APJ1, AP1, APC1 As Double

        On Error GoTo AREAerrorhandler

        jetdepth = 0.12 'depth of ceiling jet 12% of fire-ceiling height

        If T >= WallIgniteTime(fireroom) Then

            'the wall has ignited
            If y_pyrol_wall <= RH Then
                'Case 1a
                'The ignitor burner has caused ignition of wall
                'The pyrolysis height is less than the ceiling height
                If FireLocation(id) = 2 Then
                    'corner
                    If x_pyrol > x_pyrol_ign And y_pyrol_wall > y_pyrol_ign Then
                        WallArea = 2 * ((y_pyrol_wall - FireHeight(id)) * x_pyrol_ign + (x_pyrol - x_pyrol_ign) * (y_pyrol_ign - FireHeight(id)) + 0.5 * (y_pyrol_wall - y_pyrol_ign) * (x_pyrol - x_pyrol_ign))
                    Else
                        WallArea = 2 * ((y_pyrol_wall - FireHeight(id)) * x_pyrol_ign + (x_pyrol - x_pyrol_ign) * (y_pyrol_ign - FireHeight(id)))
                    End If
                ElseIf FireLocation(id) = 1 Then
                    If x_pyrol > x_pyrol_ign And y_pyrol_wall > y_pyrol_ign Then
                        'wall
                        WallArea = (y_pyrol_wall - FireHeight(id)) * x_pyrol_ign + 2 * (x_pyrol - x_pyrol_ign) * (y_pyrol_ign - FireHeight(id)) + (y_pyrol_wall - y_pyrol_ign) * (x_pyrol - x_pyrol_ign)
                    Else
                        WallArea = (y_pyrol_wall - FireHeight(id)) * x_pyrol_ign + 2 * (x_pyrol - x_pyrol_ign) * (y_pyrol_ign - FireHeight(id))
                    End If
                Else
                    'centre of room
                    WallArea = 0

                    'as for wall location (why?)
                    If x_pyrol > x_pyrol_ign And y_pyrol_wall > y_pyrol_ign Then
                        'wall
                        WallArea = (y_pyrol_wall - FireHeight(id)) * x_pyrol_ign + 2 * (x_pyrol - x_pyrol_ign) * (y_pyrol_ign - FireHeight(id)) + (y_pyrol_wall - y_pyrol_ign) * (x_pyrol - x_pyrol_ign)
                    Else
                        WallArea = (y_pyrol_wall - FireHeight(id)) * x_pyrol_ign + 2 * (x_pyrol - x_pyrol_ign) * (y_pyrol_ign - FireHeight(id))
                    End If

                End If
                CeilingArea = 0

            Else
                'Case 1b
                'The ignitor burner has caused ignition of wall
                'The pyrolysis height is greater than the ceiling height
                'wall burning continues and is assumed to ignite the ceiling
                If CeilingIgniteFlag(fireroom) = 0 And CeilingEffectiveHeatofCombustion(fireroom) > 0 Then
                    'ceiling is combustible and has not previously ignited 
                    CeilingIgniteFlag(fireroom) = 1 'true

                    Dim Message As String = "Ceiling in Room " & CStr(fireroom) & " has ignited at " & CStr(T) & " seconds (due to wall burning)."
                    MDIFrmMain.ToolStripStatusLabel2.Text = Message.ToString
                    If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                    CeilingIgniteTime(fireroom) = T
                    CeilingIgniteStep(fireroom) = stepcount
                End If

                If FireLocation(id) = 2 Then 'corner
                    AP1 = 2 * ((RH - FireHeight(id)) * x_pyrol_ign + (x_pyrol - x_pyrol_ign) * (y_pyrol_ign - FireHeight(id)) + 0.5 * (x_pyrol - x_pyrol_ign) * (RH - y_pyrol_ign))
                ElseIf FireLocation(id) = 1 Then 'wall
                    AP1 = 2 * ((RH - FireHeight(id)) * x_pyrol_ign + (x_pyrol - x_pyrol_ign) * (y_pyrol_ign - FireHeight(id)) + 0.5 * (x_pyrol - x_pyrol_ign) * (RH - y_pyrol_ign)) - (RH - FireHeight(id)) * x_pyrol_ign
                Else
                    'centre of room
                    AP1 = 0

                    'as for wall, 17/11/2011
                    AP1 = 2 * ((RH - FireHeight(id)) * x_pyrol_ign + (x_pyrol - x_pyrol_ign) * (y_pyrol_ign - FireHeight(id)) + 0.5 * (x_pyrol - x_pyrol_ign) * (RH - y_pyrol_ign)) - (RH - FireHeight(id)) * x_pyrol_ign

                End If

                'The z_pyrolysis length cannot be greater than (1-0.12) * room height
                z_pyrol = (1 - jetdepth) * RH
                If z_pyrol > x_pyrol - x_pyrol_ign Then
                    z_pyrol = x_pyrol - x_pyrol_ign
                End If
                If T = tim(stepcount, 1) Then
                    Z_pyrolysis(fireroom, stepcount) = z_pyrol
                End If

                If z_pyrol > 0 Then 'there is some lateral spread
                    'corner
                    APJ1 = 2 * ((y_pyrol - RH) * (jetdepth * (RH - FireHeight(id))) + 0.5 * z_pyrol * (y_pyrol - RH) - 0.5 * (jetdepth * (RH - FireHeight(id)) + z_pyrol) ^ 2 * ((x_pyrol - x_pyrol_ign) / (RH - y_pyrol_ign)))
                Else
                    'corner
                    APJ1 = 2 * (y_pyrol - RH) * (jetdepth * (RH - FireHeight(id)))
                End If
                If APJ1 < 0 Then APJ1 = 0

                WallArea = AP1 + APJ1
                If WallArea > 2 * RH * (RoomLength(fireroom) + RoomWidth(fireroom)) Then
                    WallArea = 2 * RH * (RoomLength(fireroom) + RoomWidth(fireroom))
                End If
                'CeilingArea = PI / 4 * (Y_pyrolysis(fireroom, stepcount) - RoomHeight(fireroom)) ^ 2
                'If CeilingArea > RoomFloorArea(fireroom) Then CeilingArea = RoomFloorArea(fireroom)
            End If
        Else
            WallArea = 0
            CeilingArea = 0
        End If

        If T >= CeilingIgniteTime(fireroom) Then
            'the ceiling has ignited
            If FireLocation(id) = 0 Then
                'centre of room
                If CeilingArea < PI * (y_pyrol - RH) ^ 2 Then
                    CeilingArea = PI * (y_pyrol - RH) ^ 2
                End If
            ElseIf FireLocation(id) = 2 Then
                'corner location
                If CeilingArea < PI / 4 * (y_pyrol - RH) ^ 2 Then
                    CeilingArea = PI / 4 * (y_pyrol - RH) ^ 2
                End If
            Else
                'wall location
                If CeilingArea < PI / 2 * (y_pyrol - RH) ^ 2 Then
                    CeilingArea = PI / 2 * (y_pyrol - RH) ^ 2
                End If
            End If
            If CeilingArea > RoomFloorArea(fireroom) Then CeilingArea = RoomFloorArea(fireroom)
        End If

       
        'surface area of wall including openings
        Dim maxwallarea As Single = 2 * RoomHeight(fireroom) * (RoomLength(fireroom) + RoomWidth(fireroom))
        Dim maxceilingarea As Single = RoomLength(fireroom) * RoomWidth(fireroom)

        If PessimiseCombWall = True Then
            If WallArea > maxwallarea * (wallpercent) / 100 Then
                'compare burning area to maximum combustible area available
                WallArea = maxwallarea * (wallpercent) / 100
            End If
        End If

        If CeilingArea > maxceilingarea * (ceilingpercent) / 100 Then
            'compare burning area to maximum combustible area available
            CeilingArea = maxceilingarea * (ceilingpercent) / 100
        End If

        Exit Sub

AREAerrorhandler:
        MsgBox(ErrorToString(Err.Number) & "in Pyrol_Area_NB3")
        ERRNO = Err.Number

    End Sub
	
    Public Function hrr_linings3(ByVal id As Integer, ByVal room As Integer, ByVal current_time As Single, ByVal floorflux_f As Single, ByVal floorflux_b As Single, ByVal NetFlux_F As Single, ByVal NetFLux_B As Single, ByRef QFloor As Single, ByRef QWall As Single, ByRef QCeiling As Single, ByRef QDotW As Double, ByRef QDotC As Double, ByRef QDotF As Double) As Object
        '====================================================================
        ' Determines the heat release rate from the compartment lining materials
        ' by storing the change in pyrolysis area for each time step and
        ' using the elapsed time to determine the heat release rate for each
        ' burning segment from the cone calorimeter input
        ' 11 June 2003 C Wade
        '  version 2003.1
        '====================================================================

        Dim i, j As Integer
        Dim hrr As Double
        Dim s() As Double
        Dim index() As Integer
        Dim x4, x1, x2, x5 As Single
        Dim burnarea As Double
        Dim curve As Integer
        Dim YW(1) As Double
        Dim YC(1) As Double
        Dim YF(1) As Double
        Dim T, x3, elapsed As Single
        'UPGRADE_NOTE: step was upgraded to step_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim QDotWB, QDotFB As Double
        Dim step_Renamed, lastpoint As Integer
        Dim ExtFlux() As Double
        Dim x21, x11, x31 As Single
        Dim Lf As Double

        On Error GoTo errorhandler

        If CurveNumber_W(room) > 0 Then ReDim YW(CurveNumber_W(room))
        If CurveNumber_C(room) > 0 Then ReDim YC(CurveNumber_C(room))
        If CurveNumber_F(room) > 0 Then ReDim YF(CurveNumber_F(room))

        hrr = 0
        QWall = 0
        QCeiling = 0
        QFloor = 0

        'the external heat flux has to be within the range of the cone data available
        If CurveNumber_W(room) > 0 Then
            x1 = NetFLux_B
            If x1 < ExtFlux_W(1, room) Then x1 = ExtFlux_W(1, room)
            If x1 > ExtFlux_W(CurveNumber_W(room), room) Then x11 = ExtFlux_W(CurveNumber_W(room), room) Else x11 = x1

            If BurnerFlameLength >= RoomHeight(room) And Y_pyrolysis(room, stepcount) >= RoomHeight(room) Then
                x2 = NetFLux_B 'the burner flame is extending to the top of the wall
            Else
                x2 = NetFlux_F
            End If

            If x2 < ExtFlux_W(1, room) Then x2 = ExtFlux_W(1, room)
            If x2 > ExtFlux_W(CurveNumber_W(room), room) Then x21 = ExtFlux_W(CurveNumber_W(room), room) Else x21 = x2
        End If

        If CurveNumber_C(room) > 0 Then
            If BurnerFlameLength >= RoomHeight(room) And Y_pyrolysis(room, stepcount) >= RoomHeight(room) Then
                x3 = NetFLux_B 'the burner flame is extending to the top of the wall
            Else
                x3 = NetFlux_F
            End If

            'incident flux on the ceiling, SFPE h/b 3rd ed. page 2-272
            'also used in the ignition calculation
            '9June2003
            Lf = BurnerWidth * 5.9 * Sqrt(Qburner / 1110 / BurnerWidth ^ (5 / 2))
            If (RoomHeight(fireroom) - FireHeight(id)) / Lf <= 0.52 Then
                x3 = 120
            Else
                x3 = 13 * ((RoomHeight(fireroom) - FireHeight(id)) / Lf) ^ (-3.5)
            End If
            'If x3 < NetFlux_F Then x3 = NetFlux_F

            If x3 < ExtFlux_C(1, room) Then x3 = ExtFlux_C(1, room)
            If x3 > ExtFlux_C(CurveNumber_C(room), room) Then x31 = ExtFlux_C(CurveNumber_C(room), room) Else x31 = x3
        End If

        If CurveNumber_F(room) > 0 Then
            x4 = floorflux_f 'ahead of the burner flame
            If x4 < ExtFlux_F(1, room) Then x4 = ExtFlux_F(1, room)
            If x4 > ExtFlux_F(CurveNumber_F(room), room) Then x4 = ExtFlux_F(CurveNumber_F(room), room)
            x5 = floorflux_b 'in the region of the burner flame
            If x5 < ExtFlux_F(1, room) Then x5 = ExtFlux_F(1, room)
            If x5 > ExtFlux_F(CurveNumber_F(room), room) Then x5 = ExtFlux_F(CurveNumber_F(room), room)
        End If

        '======================================================
        'the wall next to the burner flamer, area first ignited
        '======================================================
        If WallIgniteStep(room) > 0 And (WallConeDataFile(room) <> "" And WallConeDataFile(room) <> "null.txt") Then
            'call routine to get the qdot from the wall in the burner region = QDotWB
            If current_time > WallIgniteTime(room) Then
                elapsed = current_time - WallIgniteTime(room)
                For curve = 1 To CurveNumber_W(room)
                    If elapsed > ConeYW(room, curve, ConeNumber_W(curve, room)) Then
                        YW(curve) = ConeXW(room, curve, ConeNumber_W(curve, room)) 'hrr
                    Else
                        Call Cone_Data(room, 1, curve, elapsed, YW(curve))
                    End If
                Next curve

                ReDim ExtFlux(CurveNumber_W(room))
                For j = 1 To CurveNumber_W(room)
                    ExtFlux(j) = ExtFlux_W(j, room)
                Next j

                ReDim s(CurveNumber_W(room))
                ReDim index(CurveNumber_W(room))
                If CurveNumber_W(room) > 1 Then
                    Call CSCOEF(CurveNumber_W(room), ExtFlux, YW, s, index)
                    Call CSFIT1(CurveNumber_W(room), ExtFlux, YW, s, index, x11, QDotWB)
                Else
                    QDotWB = YW(1)
                End If
            Else
                QDotWB = 0
            End If

            If x1 = ExtFlux_W(CurveNumber_W(room), room) Then
                'extrapolate
                QDotWB = QDotWB * x1 / x11
            End If

            QWall = QWall + QDotWB * delta_area(room, 1, WallIgniteStep(room))
        End If

        '======================================================
        'the rest of the wall linings
        '======================================================
        If WallIgniteFlag(room) > 0 Then 'wall has ignited
            i = WallIgniteStep(room) + 1 'exclude the area first ignited
            Do While i < stepcount
                elapsed = current_time - tim(i, 1) 'section i started burning at tim(i,1) and tim(stepcount,1) is the curent time
                For curve = 1 To CurveNumber_W(room)

                    lastpoint = ConeNumber_W(curve, room)
                    Do While ConeYW(room, curve, lastpoint) = 0
                        lastpoint = lastpoint - 1
                    Loop

                    If elapsed > ConeYW(room, curve, lastpoint) Then
                        YW(curve) = ConeXW(room, curve, lastpoint) 'hrr
                    Else
                        Call Cone_Data(room, 1, curve, elapsed, YW(curve))
                    End If
                Next curve

                ReDim ExtFlux(CurveNumber_W(room))
                For j = 1 To CurveNumber_W(room)
                    ExtFlux(j) = ExtFlux_W(j, room)
                Next j

                ReDim s(CurveNumber_W(room))
                ReDim index(CurveNumber_W(room))
                If CurveNumber_W(room) > 1 Then
                    Call CSCOEF(CurveNumber_W(room), ExtFlux, YW, s, index)
                    Call CSFIT1(CurveNumber_W(room), ExtFlux, YW, s, index, x21, QDotW)
                Else
                    QDotW = YW(1)
                End If

                If x2 = ExtFlux_W(CurveNumber_W(room), room) Then
                    'extrapolate if the heat flux is greater than that for which we have data
                    QDotW = QDotW * x2 / x21
                End If

                'Debug.Print current_time, elapsed, QDotW, delta_area(room, 1, i)
                QWall = QWall + QDotW * delta_area(room, 1, i) 'excludes the area first ignited by the burner
                i = i + 1
            Loop
        End If

        '======================================================
        'ceiling linings
        '======================================================
        If CeilingIgniteFlag(room) > 0 Then 'ceiling has ignited
            i = CeilingIgniteStep(room)
            Do While i < stepcount
                elapsed = current_time - tim(i, 1) 'section i started burning at tim(i,1) and tim(stepcount,1) is the curent time
                For curve = 1 To CurveNumber_C(room)

                    lastpoint = ConeNumber_C(curve, room)

                    Do While ConeYC(room, curve, lastpoint) = 0
                        lastpoint = lastpoint - 1
                    Loop
                    If elapsed > ConeYC(room, curve, lastpoint) Then
                        YC(curve) = ConeXC(room, curve, lastpoint) 'hrr
                    Else
                        Call Cone_Data(room, 2, curve, elapsed, YC(curve))
                    End If
                Next curve

                ReDim ExtFlux(CurveNumber_C(room))
                For j = 1 To CurveNumber_C(room)
                    ExtFlux(j) = ExtFlux_C(j, room)
                Next j

                ReDim s(CurveNumber_C(room))
                ReDim index(CurveNumber_C(room))
                If CurveNumber_C(room) > 1 Then
                    Call CSCOEF(CurveNumber_C(room), ExtFlux, YC, s, index)
                    Call CSFIT1(CurveNumber_C(room), ExtFlux, YC, s, index, x31, QDotC)
                Else
                    QDotC = YC(1)
                End If
                'Debug.Print current_time, elapsed, QDotC, delta_area(room, 2, i)
                If x3 >= ExtFlux_C(CurveNumber_C(room), room) Then
                    'extrapolate
                    QDotC = QDotC * x3 / x31
                End If
                QCeiling = QCeiling + QDotC * delta_area(room, 2, i)
                i = i + 1
            Loop
        End If

        '======================================================
        'the floor next to the burner flamer, area first ignited
        '======================================================
        If FloorIgniteStep(room) > 0 And (FloorConeDataFile(room) <> "" And FloorConeDataFile(room) <> "null.txt") Then
            'call routine to get the qdot from the floor in the burner region = QDotFB
            If current_time > FloorIgniteTime(room) Then
                elapsed = current_time - FloorIgniteTime(room)
                For curve = 1 To CurveNumber_F(room)
                    If elapsed > ConeYF(room, curve, ConeNumber_F(curve, room)) Then
                        YF(curve) = ConeXF(room, curve, ConeNumber_F(curve, room)) 'hrr
                    Else
                        Call Cone_Data(room, 3, curve, elapsed, YF(curve))
                    End If
                Next curve

                ReDim ExtFlux(CurveNumber_F(room))
                For j = 1 To CurveNumber_F(room)
                    ExtFlux(j) = ExtFlux_F(j, room)
                Next j

                ReDim s(CurveNumber_F(room))
                ReDim index(CurveNumber_F(room))
                If CurveNumber_F(room) > 1 Then
                    Call CSCOEF(CurveNumber_F(room), ExtFlux, YF, s, index)
                    Call CSFIT1(CurveNumber_F(room), ExtFlux, YF, s, index, x5, QDotFB)
                Else
                    QDotFB = YF(1)
                End If
            Else
                QDotFB = 0
            End If
            QFloor = QFloor + QDotFB * delta_area(room, 3, FloorIgniteStep(room))
        End If

        '======================================================
        'rest of the floor linings
        '======================================================
        If FloorIgniteFlag(room) > 0 Then 'floor has ignited
            i = FloorIgniteStep(room) + 1
            Do While i < stepcount
                elapsed = current_time - tim(i, 1) 'section i started burning at tim(i,1) and tim(stepcount,1) is the curent time
                For curve = 1 To CurveNumber_F(room)

                    lastpoint = ConeNumber_F(curve, room)
                    Do While ConeYF(room, curve, lastpoint) = 0
                        lastpoint = lastpoint - 1
                    Loop

                    If elapsed > ConeYF(room, curve, lastpoint) Then
                        YF(curve) = ConeXF(room, curve, lastpoint) 'hrr
                    Else
                        Call Cone_Data(room, 3, curve, elapsed, YF(curve))
                    End If
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
                'Debug.Print current_time, elapsed, QDotF, delta_area(room, 3, i)
                QFloor = QFloor + QDotF * delta_area(room, 3, i)
                'If delta_area(room, 3, i) < 0 Then Stop
                i = i + 1
            Loop
        End If

        hrr_linings3 = QWall + QCeiling + QFloor 'kw
        Erase YW
        Erase YC
        Erase YF
        Erase ExtFlux

        Exit Function

errorhandler:
        MsgBox(ErrorToString(Err.Number) & " hrr_linings3")
        ERRNO = Err.Number
        'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'System.Windows.Forms.Cursor.Current = Default_Renamed



    End Function
End Module