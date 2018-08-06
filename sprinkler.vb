Imports System.Collections.Generic
Imports System.Math

Public Class oSprinkler
    Private m_rti As Single
    Private m_cfactor As Single
    Private m_acttemp As Single
    Private m_sprdensity As Single
    Private m_sprr As Double
    Private m_sprz As Single
    Private m_sprx As Single
    Private m_spry As Single
    Private m_sprid As Integer
    Private m_room As Integer
    Private m_responsetime As Single

    Public Sub New()
    End Sub

    Public Sub New(ByVal actuationtemp As Single, ByVal room As Integer, ByVal RTI As Single, ByVal cfactor As Single, ByVal sprx As Single, ByVal spry As Single, ByVal sprz As Single, ByVal sprr As Double, ByVal sprdensity As Single, ByVal rtidistribution As String, ByVal rtimean As Single, ByVal rtivariance As Single, ByVal rtilbound As Single, ByVal rtiubound As Single, ByVal cfdistribution As String, ByVal cfmean As Single, ByVal cfvariance As Single, ByVal cflbound As Single, ByVal cfubound As Single, ByVal actdistribution As String, ByVal actmean As Single, ByVal actvariance As Single, ByVal actlbound As Single, ByVal actubound As Single, ByVal dendistribution As String, ByVal denmean As Single, ByVal denvariance As Single, ByVal denlbound As Single, ByVal denubound As Single, ByVal rdistribution As String, ByVal rmean As Single, ByVal rvariance As Single, ByVal rlbound As Single, ByVal rubound As Single, ByVal zdistribution As String, ByVal zmean As Single, ByVal zvariance As Single, ByVal zlbound As Single, ByVal zubound As Single)

        Me.sprx = sprx
        Me.spry = spry
        Me.room = room
        Me.responsetime = Nothing
        Me.rti = RTI
        Me.cfactor = cfactor
        Me.acttemp = actuationtemp
        Me.sprdensity = sprdensity
        Me.sprr = sprr
        Me.sprz = sprz

    End Sub
    
  

    Public Property responsetime() As Single
        Get
            Return m_responsetime
        End Get
        Set(ByVal value As Single)
            m_responsetime = value
        End Set
    End Property
    Public Property acttemp() As Single
        Get
            Return m_acttemp
        End Get
        Set(ByVal value As Single)
            m_acttemp = value
        End Set
    End Property
    Public Property room() As Single
        Get
            Return m_room
        End Get
        Set(ByVal value As Single)
            m_room = value
        End Set
    End Property
    Public Property rti() As Single
        Get
            Return m_rti
        End Get
        Set(ByVal value As Single)
            m_rti = value
        End Set
    End Property
    Public Property cfactor() As Single
        Get
            Return m_cfactor
        End Get
        Set(ByVal value As Single)
            m_cfactor = value
        End Set
    End Property
    Public Property sprx() As Single
        Get
            Return m_sprx
        End Get
        Set(ByVal value As Single)
            m_sprx = value
        End Set
    End Property
    Public Property spry() As Single
        Get
            Return m_spry
        End Get
        Set(ByVal value As Single)
            m_spry = value
        End Set
    End Property
    Public Property sprz() As Single
        Get
            Return m_sprz
        End Get
        Set(ByVal value As Single)
            m_sprz = value
        End Set
    End Property
    Public Property sprr() As Double
        Get
            Return m_sprr
        End Get
        Set(ByVal value As Double)
            m_sprr = value
        End Set
    End Property
    Public Property sprid() As Integer
        Get
            Return m_sprid
        End Get
        Set(ByVal value As Integer)
            m_sprid = value
        End Set
    End Property
    Public Property sprdensity() As Single
        Get
            Return m_sprdensity
        End Get
        Set(ByVal value As Single)
            m_sprdensity = value
        End Set
    End Property
    
    Public Function GetDisplayText(ByVal sep As String) As String
        Dim text As String
        text = room & sep & VB6.Format(sprx, "0.000") & sep & VB6.Format(spry, "0.000")
        Return text
    End Function

End Class

Module Sprinklers
    Dim oSprinklers As New List(Of oSprinkler)

    Sub Derk_Sprinklers(ByVal oSprinklers As List(Of oSprinkler), ByRef SPstart() As Double, ByVal Nvariables As Integer, ByVal x1 As Double, ByVal x2 As Double, ByVal EPS As Double, ByVal H1 As Double, ByVal HMIN As Double, ByVal hnext As Double, ByVal kmax As Integer)
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

        maxsteps = 10000 'good
        'maxsteps = 10
        tiny = 1.0E-30
        Dim Hdid, X, h As Double
        ReDim Yscal(Nvariables)
        ReDim DYDX(Nvariables)
        ReDim Y(Nvariables)

        X = x1
        h = Sign(H1, x2 - x1)
        Kount = 0
        For i = 1 To Nvariables : Y(i) = SPstart(i) : Next i
        For nstp = 1 To maxsteps
            Call Derivs_sprinklers(oSprinklers, X, Y, DYDX, Nvariables)
            For i = 1 To Nvariables
                Yscal(i) = Abs(Y(i)) + Abs(h * DYDX(i)) + tiny
            Next i
            If (X + h - x2) * (X + h - x1) > 0.0# Then h = x2 - X
            Call Rk5_Sprinklers(oSprinklers, Y, DYDX, Nvariables, X, h, EPS, Yscal, Hdid, hnext)
            If (X - x2) * (x2 - x1) >= 0.0# Then
                For i = 1 To Nvariables
                    SPstart(i) = Y(i)
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

        If flagstop <> 1 Then MsgBox("Too Many Steps.")
    End Sub

    Sub ODE_Sprinkler_Solver(ByVal oSprinklers As List(Of oSprinkler), ByRef SPstart() As Double)
        '*  ====================================================================
        '*  Solve the ODE's using a 5th order Runge-Kutta numerical solution
        '*  ====================================================================

        Dim Nvariables As Integer
        Dim x2, x1, EPS As Double
        Dim HMIN, H1, hnext As Double
        Dim kmax As Integer

        Nvariables = NumSprinklers 'number of equations
        x1 = tim(stepcount, 1) 'initial time
        x2 = x1 + Timestep 'final time
        EPS = 0.000001 'decimal place accuracy
        H1 = Timestep 'suggested time step
        hnext = Timestep 'suggested time step
        HMIN = 0.00000000001 'minimum time step
        kmax = 5000 'max number of steps to be attempted in adapting step size

        Call Derk_Sprinklers(oSprinklers, SPstart, Nvariables, x1, x2, EPS, H1, HMIN, hnext, kmax)

        Exit Sub

    End Sub
    Sub Derivs_sprinklers(ByVal oSprinklers As List(Of oSprinkler), ByVal X As Double, ByVal Y() As Double, ByVal DYDX() As Double, ByVal Nvariables As Integer)

        Dim j As Integer = 1
        Dim i As Integer = stepcount
        Dim CJVel, CJTemp, CJMax As Double

        'detector link temperature
        For Each oSprinkler In oSprinklers
            If oSprinkler.room = fireroom And Flashover = False Then
                If cjModel = cjAlpert Then
                    DYDX(j) = Unconfined_Jet2(oSprinkler.rti, oSprinkler.cfactor, oSprinkler.sprr, oSprinkler.room, X, uppertemp(fireroom, i), Y(j), lowertemp(fireroom, i), layerheight(fireroom, i), CJTemp, CJMax, CJVel) 'alperts unconfined ceiling, link at max position
                ElseIf cjModel = cjJET Then
                    DYDX(j) = JET6(oSprinkler.rti, oSprinkler.cfactor, oSprinkler.sprz, oSprinkler.sprr, oSprinkler.room, X, layerheight(oSprinkler.room, i), uppertemp(oSprinkler.room, i), lowertemp(oSprinkler.room, i), Y(j), CJTemp, CJMax, CJVel) 'Davis' JET correlation, hot layer effects, link at variable position
                End If
                If X = tim(i, 1) Then
                    CJetTemp(i, 1, oSprinkler.sprid - 1) = CJTemp
                    CJetTemp(i, 2, oSprinkler.sprid - 1) = CJMax
                    CJetTemp(i, 0, oSprinkler.sprid - 1) = CJVel
                End If
            Else
                'not in room of origin
                DYDX(j) = 0
            End If
            j = j + 1
        Next


    End Sub

    Sub Rk5_Sprinklers(ByVal oSprinklers As List(Of oSprinkler), ByRef Y() As Double, ByRef DYDX() As Double, ByVal n As Integer, ByRef X As Double, ByRef Htry As Double, ByVal EPS As Double, ByRef Yscal() As Double, ByRef Hdid As Double, ByRef hnext As Double)
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
1:      Hh = 0.5 * h ' Take Two Half Step Sizes

        Call RK4_Sprinklers(oSprinklers, YSAV, DYSAV, n, Xsaved, Hh, Ytemp)

        X = Xsaved + Hh

        Call Derivs_sprinklers(oSprinklers, X, Ytemp, DYDX, n)

        Call RK4_Sprinklers(oSprinklers, Ytemp, DYDX, n, X, Hh, Y)

        X = Xsaved + h

        If X = Xsaved Then
            MsgBox("Stepsize Not Significant In Rkqc.")
            flagstop = 1
            Exit Sub
        End If

        Call RK4_Sprinklers(oSprinklers, YSAV, DYSAV, n, Xsaved, h, Ytemp)

        errmax = 0.0#

        For i = 1 To n
            Ytemp(i) = Y(i) - Ytemp(i)
            If Abs(Ytemp(i) / Yscal(i)) > errmax Then
                errmax = Abs(Ytemp(i) / Yscal(i))
            End If
        Next i

        errmax = errmax / EPS

        If errmax > 1.0# Then
            h = Safety * h * errmax ^ Pshrnk
            GoTo 1
        Else
            Hdid = h
            If errmax > Errcon Then
                hnext = Safety * h * errmax ^ PGrow
            Else
                hnext = 4.0# * h
            End If
        End If

        For i = 1 To n
            Y(i) = Y(i) + Ytemp(i) * Fcor
        Next i
    End Sub
    Sub RK4_Sprinklers(ByVal oSprinklers As List(Of oSprinkler), ByRef Y() As Double, ByRef DYDX() As Double, ByVal n As Integer, ByRef X As Double, ByVal h As Double, ByRef Yout() As Double)
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

        Call Derivs_sprinklers(oSprinklers, XH, Yt, Dyt, n)

        For i = 1 To n
            Yt(i) = Y(i) + Hh * Dyt(i)
        Next i

        Call Derivs_sprinklers(oSprinklers, XH, Yt, Dym, n)

        For i = 1 To n
            Yt(i) = Y(i) + h * Dym(i)
            Dym(i) = Dyt(i) + Dym(i)
        Next i

        Call Derivs_sprinklers(oSprinklers, X + h, Yt, Dyt, n)

        For i = 1 To n
            Yout(i) = Y(i) + H6 * (DYDX(i) + Dyt(i) + 2.0# * Dym(i))
        Next i

    End Sub

    Public Function JET6(ByVal rti As Single, ByVal cfactor As Single, ByVal SprinkDistance As Single, ByVal RadialDistance As Double, ByVal room As Integer, ByVal tim As Double, ByVal layerz As Double, ByVal UT As Double, ByVal LT As Double, ByVal LinkTemp As Double, ByRef CJTemp As Double, ByRef CJMax As Double, ByRef CJVel As Double) As Double
        '*  =====================================================================
        '*
        '*      Variables returned by the function
        '*      CJTemp = the ceiling jet temperature at a specified radial distance and distance below the ceiling in K
        '*      CJMax = the maximum ceiling jet temperature at a specified radial distance in m
        '*
        '*       Based on NIST Model JET by Davis, NISTIR 6324
        '*       21/5/2003
        '*  =====================================================================

        Dim Qi1, beta, CT, Qi2 As Double
        Dim Zi2, H2, Zi1, H1 As Single
        Dim delPlumeTemp As Double
        Dim delCeilingJetTemp As Double
        Dim yj, alpha, yl As Single
        Dim ro As Double
        Dim gamma_2, epsilon, Q As Single
        Dim ratio_v, velocity, C, delta, ratio_t As Double
        Dim psi As Double

        Dim itemid As Integer = 1

        beta = 0.9555 'velocity to temperature ratio of gaussian profile half widths
        CT = 9.115
        alpha = 0.44

        If room = fireroom Then
            'get total heat release rate for object 1 at this time (q in kW)
            Q = Get_HRR(itemid, tim, 0, 0, 0, 0, 0)

            'use method of reflection to account for fires located in corner or against wall
            If FireLocation(itemid) = 2 Then Q = 4 * Q 'corner fire
            If FireLocation(itemid) = 1 Then Q = 2 * Q 'wall fire

            Zi1 = layerz - FireHeight(itemid) 'Fireheight(1) = the height of the base of the fire above the floor
            H1 = RoomHeight(room) - FireHeight(itemid) 'uses height of object 1
            epsilon = UT / LT
            If epsilon < 1 Then Exit Function ' shouldn't happen!

            If Zi1 <= 0.01 Then Zi1 = 0.001 'in the limit - calculate based on the layer being just above the fire height

            'calculate non-dimensional Qi1
            If Zi1 > 0 Then
                'Qi1 = (1 - NewRadiantLossFraction) * Q / ((Atm_Pressure) / (Gas_Constant / MW_air) * SpecificHeat_air * Sqrt(G) * Zi1 ^ (5 / 2))
                Qi1 = (1 - ObjectRLF(itemid)) * Q / ((Atm_Pressure) / (Gas_Constant / MW_air) * SpecificHeat_air * Sqrt(G) * Zi1 ^ (5 / 2))
            Else
                Exit Function
            End If

            If (1 + CT * Qi1 ^ (2 / 3)) / (epsilon * CT) > 1 / CT Then 'error checking, cannot raise a -ve number to power
                'strength of substitute source
                Qi2 = ((1 + CT * Qi1 ^ (2 / 3)) / (epsilon * CT) - 1 / CT) ^ (3 / 2) 'as i read bill davis' paper (version 2002.5 and before)
            Else
                Qi2 = Qi1 'from Davis code
            End If

            'location of substitute source
            Zi2 = ((epsilon * Qi1 * CT) / (Qi2 ^ (1 / 3) * ((epsilon - 1) * (beta ^ 2 + 1) + epsilon * CT * Qi2 ^ (2 / 3)))) ^ (2 / 5) * Zi1
            H2 = H1 - Zi1 + Zi2

            'excess relative to upper layer temp
            'plume centreline temp excess at the ceiling
            delPlumeTemp = 9.28 * UT * (Qi2) ^ (2 / 3) * (Zi2 / H2) ^ (5 / 3) 'uses non-dimensional Qi2

            yj = 0.1 * H1 'ceiling jet thickness
            yl = RoomHeight(room) - layerz 'layer thickness
            ro = 0.18 * H1
            gamma_2 = 2 / 3 - alpha * (1 - Exp(-yl / yj))
            beta = 0.676 + 0.164 * (1 - Exp(-yl / yj))
            C = beta * (ro ^ gamma_2) * (delPlumeTemp + UT - LT)

            If RadialDistance > ro Then
                delCeilingJetTemp = C / (RadialDistance ^ gamma_2)
            Else
                delCeilingJetTemp = C / (ro ^ gamma_2)
            End If

            If RadialDistance > 0.15 * H1 Then
                'velocity = 0.195 * ((1 - NewRadiantLossFraction) * Q) ^ (1 / 3) * Sqrt(H1) / (RadialDistance ^ (5 / 6))
                velocity = 0.195 * ((1 - ObjectRLF(itemid)) * Q) ^ (1 / 3) * Sqrt(H1) / (RadialDistance ^ (5 / 6))
            Else
                'velocity = 0.96 * ((1 - NewRadiantLossFraction) * Q / H1) ^ (1 / 3)
                velocity = 0.96 * ((1 - ObjectRLF(itemid)) * Q / H1) ^ (1 / 3)
            End If
            velocity = velocity * (1 - 0.25 * (1 - Exp(-yl / yj)))

            CJMax = delCeilingJetTemp + LT 'the maximum ceiling jet temperature at specified radial distance in K

            'variation with depth of link beneath ceiling - only the region from the ceiling surface to the
            'depth at which the maximum temperature occurs is considered, ceiling jet at greater depths are
            'assumed to be at the maximum temperature
            'LAVENT NPFA 204 Appendix B
            delta = 0.1 * H1 * (RadialDistance / H1) ^ 0.9

            'adjustment for location of link below ceiling
            If RadialDistance / H1 >= 0.2 Then
                If SprinkDistance >= 0.23 * delta Then
                    ratio_v = COSH((0.23 / 0.77) * ACOSH(Sqrt(2)) * (SprinkDistance / 0.23 / delta - 1)) ^ (-2)
                    ratio_t = ratio_v
                Else
                    ratio_v = 8 / 7 * (SprinkDistance / 0.23 / delta) ^ (1 / 7) * (1 - SprinkDistance / 0.23 / delta / 8)
                    If (delCeilingJetTemp + LT - UT) <> 0 Then
                        psi = (CeilingTemp(room, stepcount) - UT) / (delCeilingJetTemp + LT - UT)
                    Else
                        psi = 0
                    End If
                    ratio_t = psi + 2 * ((1 - psi) * (SprinkDistance / 0.23 / delta)) - ((1 - psi) * (SprinkDistance / 0.23 / delta) ^ 2)
                End If
            Else
                ratio_v = 1
                ratio_t = 1
            End If

            velocity = velocity * ratio_v 'velocity in the ceiling jet at the position of the sprinkler

            delCeilingJetTemp = ratio_t * (delCeilingJetTemp + LT - UT) 'this is the rise (above upper layer) in the ceiling jet temp at the location of sensor link
            delCeilingJetTemp = delCeilingJetTemp + UT - LT 'excess above lower temp

            If (SprinkDistance > yj) And (RadialDistance / H1 >= 0.2) Then 'not in ceiling jet
                If RoomHeight(room) - SprinkDistance < layerz Then 'not in the ceiling jet, but in lower layer
                    delCeilingJetTemp = 0
                End If
            End If

            CJTemp = LT + delCeilingJetTemp 'the maximum ceiling jet temperature at specified radial distance and distance below ceiling in K

            If CJMax < UT And 0.23 * delta < yl Then CJMax = UT 'if max location is in the upper layer
            If CJMax < LT And 0.23 * delta > yl Then CJMax = LT 'if max location is in the lower layer
            If CJTemp < UT And SprinkDistance < yl And SprinkDistance > 0.23 * delta Then CJTemp = UT 'if max location is in the upper layer
            If CJTemp < LT And SprinkDistance > yl And SprinkDistance > 0.23 * delta Then CJTemp = LT 'if max location is in the lower layer

            CJVel = velocity

            If velocity > 0 Then
                'If DetectorType(room) > 1 Then
                JET6 = Sqrt(velocity) / rti * (delCeilingJetTemp + LT - LinkTemp - cfactor / Sqrt(velocity) * (LinkTemp - InteriorTemp))
                'Else
                'JET6 = Sqrt(velocity) / rti * (delCeilingJetTemp + LT - LinkTemp)
                'End If
            Else
                JET6 = 0
            End If
            Exit Function

        Else 'not the fireroom

        End If

        Exit Function
    End Function
    Public Function Unconfined_Jet2(ByVal rti As Single, ByVal cfactor As Single, ByVal radialdistance As Single, ByVal room As Integer, ByVal tim As Double, ByVal UT As Double, ByVal LinkT As Double, ByVal LT As Double, ByVal layerz As Double, ByRef CJTemp As Double, ByRef CJMax As Double, ByRef CJVel As Double) As Double
        '*  =====================================================================
        '*      Find the value of the derivative for the temperature of the detector
        '*      link. Uses Alperts unconfined ceiling jet correlations.
        '*      Ignores presence of a hot layer
        '*
        '*      Argument passed to the function are:
        '*          tim = time (sec)
        '*          UT = upper layer temperature (K)
        '*          LinkT = link temperature (K)
        '*  =====================================================================

        Dim velocity, Q, gastemp As Double

        If room = fireroom Then
            'get heat release rate for object 1 at this time
            Q = Get_HRR(1, tim, 0, 0, 0, 0, 0)

            'use method of reflection to account for fires located in corner or against wall
            If FireLocation(1) = 2 Then Q = 4 * Q 'corner fire
            If FireLocation(1) = 1 Then Q = 2 * Q 'wall fire

            'get maximum velocity in the ceiling jet
            velocity = CeilingJet_Velocity2(Q, radialdistance)

            'get maximum temperature rise in the ceiling jet
            gastemp = CeilingJet_Temp2(UT, Q, radialdistance)

            CJMax = gastemp + InteriorTemp
            CJTemp = gastemp + InteriorTemp
            CJVel = velocity
            ' If stepcount = 150 Then Stop
            If DetectorType(room) > 1 Then
                Unconfined_Jet2 = Sqrt(velocity) / rti * (CJTemp - LinkT) - cfactor / rti * (LinkT - InteriorTemp)
            Else
                Unconfined_Jet2 = Sqrt(velocity) / rti * (CJTemp - LinkT)
            End If
        Else
            Unconfined_Jet2 = 0
        End If
    End Function

End Module
