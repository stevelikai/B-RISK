Option Strict Off
Option Explicit On
Module SmokeDetection
	
	Public Sub RK5_smoke(ByVal Tau As Double, ByVal OD_out As Double, ByRef Y() As Double, ByRef DYDX() As Double, ByVal n As Integer, ByRef X As Double, ByRef Htry As Double, ByVal EPS As Double, ByRef Yscal() As Double, ByRef Hdid As Double, ByRef hnext As Double)
		'////////////////////////////////////////////////////////////////////////
		'   Rk5 Performs A Single Step Of Fifth-order Runge-kutta Integration
		'   With Local Truncation Error Estimate And Corresponding Step
		'   Size Adjustment. Routine Adapted From Numerical Recipes By
		'   William H. Price Et Al.
		'   From Promath 2.0.
		'
		'*  Calls RK4_Smoke
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
		
		Call RK4_Smoke(Tau, OD_out, YSAV, DYSAV, n, Xsaved, Hh, Ytemp)
		
		X = Xsaved + Hh
		
		Call Derivs_smoke(Tau, OD_out, Ytemp, DYDX)
		
		Call RK4_Smoke(Tau, OD_out, Ytemp, DYDX, n, X, Hh, Y)
		
		X = Xsaved + h
		
		If X = Xsaved Then
			MsgBox("Stepsize Not Significant In Rkqc.")
			flagstop = 1
			Exit Sub
		End If
		
		Call RK4_Smoke(Tau, OD_out, YSAV, DYSAV, n, Xsaved, h, Ytemp)
		
		errmax = 0#
		
		For i = 1 To n
			Ytemp(i) = Y(i) - Ytemp(i)
			If System.Math.Abs(Ytemp(i) / Yscal(i)) > errmax Then
				errmax = System.Math.Abs(Ytemp(i) / Yscal(i))
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
	
    Public Sub RK4_Smoke(ByVal Tau As Double, ByVal OD_out As Double, ByRef Y() As Double, ByRef DYDX() As Double, ByVal n As Integer, ByRef X As Double, ByVal h As Double, ByRef Yout() As Double)
        '///////////////////////////////////////////////////////////////////////
        '   Rk4 Advances A Solution Vector Y(N) Of A Set Of Ordinary Diff.
        '   Eqs. Over A Single Small Interval H Using Fourth-order Runge-kutta
        '   Method.
        '//////////////////////////////////////////////////////////////////////

        Dim H6, Hh, XH As Double
        Dim i As Integer
        Dim Yt() As Double
        Dim Dyt() As Double
        Dim Dym() As Double
        ReDim Yt(n)
        ReDim Dyt(n)
        ReDim Dym(n)

        'UPGRADE_WARNING: Couldn't resolve default property of object h. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Hh = h * 0.5
        'UPGRADE_WARNING: Couldn't resolve default property of object h. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        H6 = h / 6.0#
        XH = X + Hh

        For i = 1 To n
            Yt(i) = Y(i) + Hh * DYDX(i)
        Next i

        Call Derivs_smoke(Tau, OD_out, Yt, Dyt)

        For i = 1 To n
            Yt(i) = Y(i) + Hh * Dyt(i)
        Next i

        Call Derivs_smoke(Tau, OD_out, Yt, Dym)

        For i = 1 To n
            'UPGRADE_WARNING: Couldn't resolve default property of object h. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Yt(i) = Y(i) + h * Dym(i)
            Dym(i) = Dyt(i) + Dym(i)
        Next i

        Call Derivs_smoke(Tau, OD_out, Yt, Dyt)

        For i = 1 To n
            Yout(i) = Y(i) + H6 * (DYDX(i) + Dyt(i) + 2.0# * Dym(i))
        Next i

    End Sub
    Public Sub Derk_Smoke(ByVal Tau As Double, ByVal OD_out As Double, ByRef Xstart() As Double, ByVal Nvariables As Integer, ByVal x1 As Double, ByVal x2 As Double, ByVal EPS As Double, ByVal H1 As Double, ByVal HMIN As Double, ByVal hnext As Double, ByVal kmax As Integer)
        '////////////////////////////////////////////////////////////////////////
        '   Derk2 Integrates A Set Of Diff. Eqs. Over An Interval Using
        '   Fifth-order Runge-kutta Method With Adaptive Step Size.
        '   This Routine Calls Rk5 To Do The Integration
        '***************************************************************************
        Dim i, maxsteps, Kount, nstp As Integer
        Dim tiny As Double
        Dim flagcount As Integer
        Dim Y() As Double
        Dim Yscal() As Double
        Dim DYDX() As Double

        flagcount = 0

        maxsteps = 10000
        tiny = 1.0E-30
        Dim Hdid, X, h As Double
        ReDim Yscal(Nvariables)
        ReDim Y(Nvariables)
        ReDim DYDX(Nvariables)

        X = x1
        h = Sign(H1, x2 - x1)
        Kount = 0
        For i = 1 To Nvariables : Y(i) = Xstart(i) : Next i
        For nstp = 1 To maxsteps
            Call Derivs_smoke(Tau, OD_out, Y, DYDX)
            For i = 1 To Nvariables
                Yscal(i) = System.Math.Abs(Y(i)) + System.Math.Abs(h * DYDX(i)) + tiny
            Next i
            If (X + h - x2) * (X + h - x1) > 0.0# Then h = x2 - X
            Call RK5_smoke(Tau, OD_out, Y, DYDX, Nvariables, X, h, EPS, Yscal, Hdid, hnext)
            If (X - x2) * (x2 - x1) >= 0.0# Then
                For i = 1 To Nvariables
                    Xstart(i) = Y(i)
                Next i
                If kmax Then Kount = Kount + 1
                Exit Sub
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object HMIN. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If System.Math.Abs(hnext) < HMIN Then
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

        If flagstop <> 1 Then MsgBox("Too Many Steps.")


    End Sub
	Public Sub Derivs_smoke(ByVal Tau As Double, ByVal OD_out As Double, ByRef YZ() As Double, ByRef Dyz() As Double)
		
		'rate of change of the optical density inside the detector chamber
		Dyz(1) = (OD_out - YZ(1)) / Tau
		
	End Sub
    Public Function SmokeJET(ByVal SDRadialDist As Single, ByVal mw_upper As Double, ByVal mw_lower As Double, ByVal pressure As Double, ByVal Yield As Single, ByVal room As Integer, ByVal layerz As Double, ByVal Q As Double, ByVal SMF As Double, ByVal diameter As Double, ByVal UT As Double, ByVal LT As Double) As Double
        '*  =====================================================================
        '*
        '*       Based on NIST Model by Davis, Cleary, Donnelly and Hellerman
        '*       15/10/2003
        '*       calculation applies to burning object number 1 only
        '*  =====================================================================
        Dim H1, Z1, H2, Z2, C1 As Double
        Dim k, zo, D, qstar, Q2star As Double
        Dim rhoL, Cso, Cspo, CL, rhoU As Double

        On Error GoTo errorhandler

        If EnergyYield(1) = 0 Then Exit Function
        C1 = 0.12
        Z1 = layerz - FireHeight(1)
        If Z1 < 0 Then Z1 = 0

        H1 = RoomHeight(room) - FireHeight(1) 'height above fire

        'location of virtual point source
        zo = Virtual_Source(Q, diameter)

        'density of lower layer
        rhoL = (Atm_Pressure + pressure / 1000) / (Gas_Constant / mw_lower * LT)

        'density of upper layer
        rhoU = (Atm_Pressure + pressure / 1000) / (Gas_Constant / mw_upper * UT)

        'calculate Q star
        If Z1 >= zo Then
            qstar = Q / (rhoL * SpecificHeat_air * LT * G ^ 0.5 * (Z1 - zo) ^ (5 / 2))
        Else
            qstar = Q / (rhoL * SpecificHeat_air * LT * G ^ 0.5 * (Z1) ^ (5 / 2))
        End If

        D = Yield * (1 + 1 / 1.157) * rhoL * SpecificHeat_air * LT / (3.4 * EnergyYield(1) * 1000 * PI * (1 - NewRadiantLossFraction(1)) ^ (1 / 3) * 1.201 ^ 2 * C1 ^ 2)

        k = 9.1 * (1 - NewRadiantLossFraction(1)) ^ (2 / 3)

        'species concentration in the upper layer
        CL = SMF * rhoU 'kg/m3

        'substitute source strength
        If (D * qstar ^ (2 / 3) - CL * (1 + k * qstar ^ (2 / 3))) >= 0 Then
            If ((D * qstar ^ (2 / 3) - CL * (1 + k * qstar ^ (2 / 3))) / (D + CL * k * (1 + k * qstar) ^ (2 / 3))) > 0 Then
                Q2star = ((D * qstar ^ (2 / 3) - CL * (1 + k * qstar ^ (2 / 3))) / (D + CL * k * (1 + k * qstar) ^ (2 / 3))) ^ (3 / 2)
            Else
                Q2star = 0
            End If
        Else
            'MsgBox "Error in function SmokeJet."
            'Exit Function
            Q2star = 0
        End If

        'location of substitute source
        If qstar > 0 And Q2star > 0 Then
            Z2 = qstar / (Q2star + (1 + 1 / 1.157) * CL / D * (1 + k * Q2star ^ (2 / 3)) * Q2star ^ (1 / 3))
        Else
            Z2 = 0
        End If
        Z2 = Z2 ^ (2 / 5)
        Z2 = Z2 * Z1

        H2 = H1 - Z1 + Z2 'H-Z is the layer depth to be the same in both cases

        'plume centerline scalar concentration at ceiling
        If H2 > 0 Then Cspo = D * Q2star ^ (2 / 3) * (Z2 / H2) ^ (5 / 3) / (1 + k * Q2star ^ (2 / 3) * (Z2 / H2) ^ (5 / 3)) + CL

        'max scalar concentration in ceiling jet at r=0.018H
        Cso = System.Math.Sqrt(2) * (1.157 / 2.157) * (Cspo - CL) + CL

        If SDRadialDist > 0.18 * H1 Then
            'max scalar concentration in ceiling jet at r
            SmokeJET = (Cso - CL) * (0.18 * H1 / SDRadialDist) ^ 0.57 + CL
        Else
            SmokeJET = Cso
        End If
        Exit Function

errorhandler:
        Resume Next

    End Function

    Function CeilingJet_Velocity3(ByVal Q As Single, ByVal room As Single, ByVal layerz As Double, ByVal detector_radial As Single, ByVal detector_depth As Single) As Double
        '*  =====================================================================
        '*
        '*      Calculate the ceiling jet velocity as used in the JET4 routine here.
        '*      Account for both radial distance from plume and distance below the ceiling
        '*      20/10/2001
        '*      Colleen Wade
        '*  =====================================================================

        Dim Z1, H2, H1 As Single
        Dim yj, yl As Single
        Dim ro, delta As Double
        Dim velocity, ratio_v As Double

        On Error GoTo errhandler

        If room = fireroom Then

            Z1 = layerz - FireHeight(1) 'Fireheight(1) = the height of the base of the fire above the floor
            H1 = RoomHeight(room) - FireHeight(1) 'uses height of object 1

            yj = 0.1 * H1 'ceiling jet thickness
            yl = RoomHeight(room) - layerz 'layer thickness
            ro = 0.15 * H1

            If detector_radial > ro Then
                velocity = 0.195 * ((1 - NewRadiantLossFraction(1)) * Q) ^ (1 / 3) * System.Math.Sqrt(H1) / (detector_radial ^ (5 / 6))
            Else
                velocity = 0.96 * ((1 - NewRadiantLossFraction(1)) * Q / H1) ^ (1 / 3)
            End If
            velocity = velocity * (1 - 0.25 * (1 - System.Math.Exp(-yl / yj)))


            'variation with depth of link beneath ceiling - only the region from the ceiling surface to the
            'depth at which the maximum temperature occurs is considered, ceiling jet at greater depths are
            'assumed to be at the maximum temperature
            'LAVENT NPFA 204 Appendix B
            If velocity > 0 Then
                delta = 0.1 * H1 * (detector_radial / H1) ^ 0.9

                'adjustment for location of link below ceiling
                If detector_radial / H1 >= 0.2 Then
                    If detector_depth >= 0.23 * delta Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object ACOSH(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object COSH(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        ratio_v = COSH((0.23 / 0.77) * ACOSH(System.Math.Sqrt(2)) * (detector_depth / 0.23 / delta - 1)) ^ (-2)
                    Else
                        ratio_v = 8 / 7 * (detector_depth / 0.23 / delta) ^ (1 / 7) * (1 - detector_depth / 0.23 / delta / 8)
                    End If
                Else
                    ratio_v = 1
                End If
            Else
                ratio_v = 0
            End If

            CeilingJet_Velocity3 = velocity * ratio_v 'velocity in the ceiling jet at the position of the sprinkler

            Exit Function
        Else
            'this is not the fireroom, there is no ceiling jet
            CeilingJet_Velocity3 = 0
        End If
        Exit Function
errhandler:
        If Err.Number = 6 Then
            Err.Clear()
            Resume Next
        Else
            MsgBox(ErrorToString(Err.Number) & " in subroutine CeilingJet_Velocity3")
        End If

    End Function
	
	Public Sub Smoke_Solver(ByVal Tau As Double, ByRef Xstart() As Double, ByRef OD_out As Double)
		'*  ====================================================================
		'*  Solve the ODE's for the flame spread using a 5th order Runge-Kutta numerical solution
		'*
		'*  ====================================================================
		
		Dim Nvariables As Integer
		Dim x2, x1, EPS As Double
		Dim HMIN, H1, hnext As Double
		Dim kmax As Integer
		Dim NStep As Integer
		
		On Error GoTo ODEerrorhandler
		
		Nvariables = 1 'number of equations
		x1 = tim(stepcount, 1) 'initial time
		x2 = x1 + Timestep 'final time
		EPS = 0.000001 'decimal place accuracy
		H1 = Timestep 'suggested time step
		hnext = Timestep 'suggested time step
		HMIN = 0.00000000001 'minimum time step
		kmax = 5000 'max number of steps to be attempted in adapting step size
		
		Call Derk_Smoke(Tau, OD_out, Xstart, Nvariables, x1, x2, EPS, H1, HMIN, hnext, kmax)
		
		Exit Sub
		
ODEerrorhandler: 
		MsgBox(ErrorToString(Err.Number) & " Smoke_Solver")
		ERRNO = Err.Number
			
	End Sub
End Module