Option Strict Off
Option Explicit On
Imports System.Math
Imports CenterSpace.NMath.Core
Imports CenterSpace.NMath.Analysis

Module STIFFsolv
    'edits:
    '26/2/07 - corrected HCN error in ode equation, sub: diff_eqns, diff_eqns2
    '
    Dim Y_stiff(,) As Double
    Dim DYDX_stiff(,) As Double


    Function Rigid(ByVal T As Double, ByVal Y As DoubleVector) As DoubleVector

        Dim Nvariables As Integer = 19
        Dim VectorLength As Integer
        Dim dy As New DoubleVector()
        VectorLength = Nvariables * NumberRooms
        dy.Resize(VectorLength)

        Call diff_eqns(T, Y_stiff, DYDX_stiff)

        Dim index As Integer = 0
        For room = 1 To NumberRooms
            For var = 1 To Nvariables
                dy(index) = DYDX_stiff(room, var)
                index = index + 1
            Next
        Next

        Return dy

    End Function





    Sub ODE_Solver_NMath(ByRef Ystart(,) As Double)
        Try

            Dim room As Integer
            Dim index As Integer = 0
            Dim x2, x1 As Double
            Dim Nvariables As Integer = 19
            Dim VLength As Integer

            ReDim Y_stiff(NumberRooms + 1, Nvariables)
            ReDim DYDX_stiff(NumberRooms + 1, Nvariables)

            'NMathConfiguration.LicenseKey = "XTUU-3UM1-YXFF-89WC-DAUS-VUUF-1QS1-WPUF-19SU-WPSS-8EJE-680T-FRPP-X8NS-WQSF-E842-68RT-WPTF-683S-X1D0-YR8V-19P2-9D3Y-LRTU-A10U"

            Y_stiff = Ystart.Clone

            x1 = tim(stepcount, 1) 'initial time
            x2 = x1 + Timestep 'final time

            Dim TimeSpan As New DoubleVector(x1, x2)
            Dim y0 As New DoubleVector()
            VLength = Nvariables * NumberRooms

            y0.Resize(VLength)

            For room = 1 To NumberRooms
                For var = 1 To Nvariables
                    y0(index) = Y_stiff(room, var)
                    index = index + 1
                Next
            Next

            ' Construct the solver.
            Dim Solver As New VariableOrderOdeSolver()
            Dim SolverOptions As New VariableOrderOdeSolver.Options()

            ' Gets and sets the bound on the estimated error at each integration step. 
            SolverOptions.AbsoluteTolerance = New DoubleVector(0.0001, 0.0001, 0.0001, 0.0001, 0.0001, 0.0001, 0.0001, 0.0001, 0.0001, 0.0001, 0.0001, 0.0001, 0.0001, 0.0001, 0.0001, 0.0001, 0.0001, 0.0001, 0.0001)
            'SolverOptions.AbsoluteTolerance = New DoubleVector(0.000001, 0.000001, 0.000001, 0.000001, 0.000001, 0.000001, 0.000001, 0.000001, 0.000001, 0.000001, 0.000001, 0.000001, 0.000001, 0.000001, 0.000001, 0.000001, 0.000001, 0.000001, 0.000001)
            'SolverOptions.AbsoluteTolerance = New DoubleVector() 'this would need to be resized and populated based on the number of rooms in the model.

            ' Bound on the estimated error at each integration step. At the ith integration step the error, e[i] for the estimated solution y[i] satisfies e[i] <= max(RelativeTolerance*Math.Abs(y[i]), AbsoluteTolerance[i]) 
            SolverOptions.RelativeTolerance = 0.001
            'SolverOptions.RelativeTolerance = 0.00001

            'Increases the number of output points by the specified factor producing smoother output. If Refine is n which is greater than 1, the solver subdivides each time step into n smaller intervals and returns solutions at each time point. The extra values produced for Refine are computed by means of continuous extension formulas The default for RungeKutta45OdeSolver solver is 4. 
            SolverOptions.Refine = 1

            SolverOptions.MaxStepSize = 0.0001

            ' Rosenbrock32stiff is a variable-order solver for stiff problems. It is based on the numerical differentiation formulas (NDFs). The NDFs are generally more efficient than the closely related family of backward differentiation formulas (BDFs), also known as Gear's methods. The ode15s properties let you choose among these formulas, as well as specifying the maximum order for the formula used. 
            SolverOptions.UseBackwardDifferention = True

            ' Construct the delegate representing our system of differential equations...
            Dim odeFunction As New Func(Of Double, DoubleVector, DoubleVector)(AddressOf Rigid)

            ' ...and solve. The solution is returned as a key/value pair. The first 'Key' element of the pair is
            ' the time span vector, the second 'Value' element of the pair is the corresponding solution values.
            ' That is, if the computed solution function is y then
            ' y(soln.Key(i)) = soln.Value(i)
            Dim Soln As VariableOrderOdeSolver.Solution(Of DoubleMatrix) = Solver.Solve(odeFunction, TimeSpan, y0, SolverOptions)

            'Console.WriteLine("T = " & Soln.T.ToString("G5"))
            'Console.WriteLine("Y = ")
            'Console.WriteLine(Soln.Y.ToTabDelimited("G5"))

            index = 0
            For room = 1 To NumberRooms
                For var = 1 To Nvariables
                    Y_stiff(room, var) = Soln.Y.Col(var - 1).Last
                    index = index + 1
                Next
            Next

            Ystart = Y_stiff.Clone

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception")
            flagstop = 1
        End Try
    End Sub
    Sub ODE_Solver3(ByRef Ystart(,) As Double)
        '*  ====================================================================
        '*  Solve the ODE's using a 5th order Runge-Kutta numerical solution
        '*
        '*  Called by main_program
        '*
        '*  Revised : 2 May 1996 Colleen Wade
        '*  ====================================================================

        Dim Nvariables As Short
        Dim x2, x1, EPS As Double
        Dim HMIN, H1, hnext As Double
        'Dim kmax As Short

        Nvariables = 19
        x1 = tim(stepcount, 1) 'initial time
        x2 = x1 + Timestep 'final time
        EPS = Error_Control 'decimal place accuracy
        'EPS = 0.01
        H1 = Timestep 'suggested time step
        hnext = Timestep 'suggested time step
        'HMIN = 0.0000001 'minimum time step
        HMIN = 0.000000001 'minimum time step

        'kmax = 5000              'max number of steps to be attempted in adapting step size
        'NStep = 5

        Call STIFF3(Nvariables, Ystart, x1, x2, EPS, H1, HMIN)

    End Sub

    Sub StiffSolve()

        ' Construct the solver.
        Dim Solver As New VariableOrderOdeSolver()

        ' Construct the time span vector. If this vector contains exactly
        ' two points, the solver interprets these to be the initial and
        ' final time values. Step size and function output points are 
        ' provided automatically by the solver. Here the initial time
        ' value t0 is 0.0 and the final time value is 12.0.
        Dim TimeSpan As New DoubleVector(0.0, 12.0)

        ' Initial y vector.
        Dim y0 As New DoubleVector(0.0, 1.0, 1.0)

        ' Construct solver options. Here we set the absolute and relative tolerances to use.
        ' At the ith integration step the error, e(i) for the estimated solution
        ' y(i) satisfies
        ' e(i) <= max(RelativeTolerance * Math.Abs(y(i)), AbsoluteTolerance(i))
        ' The solver can increase the number of output points by a specified factor
        ' called Refine (useful for creating smooth plots). The default value is 
        ' 4. Here we set the Refine value to 1, meaning we do not wish any
        ' additional output points to be added by the solver.
        Dim SolverOptions As New VariableOrderOdeSolver.Options()
        ' Gets and sets the bound on the estimated error at each integration step. 
        SolverOptions.AbsoluteTolerance = New DoubleVector(0.0001, 0.0001, 0.00001)
        ' Bound on the estimated error at each integration step. At the ith integration step the error, e[i] for the estimated solution y[i] satisfies e[i] <= max(RelativeTolerance*Math.Abs(y[i]), AbsoluteTolerance[i]) 
        SolverOptions.RelativeTolerance = 0.0001
        'Increases the number of output points by the specified factor producing smoother output. If Refine is n which is greater than 1, the solver subdivides each time step into n smaller intervals and returns solutions at each time point. The extra values produced for Refine are computed by means of continuous extension formulas The default for RungeKutta45OdeSolver solver is 4. 
        SolverOptions.Refine = 1
        ' Rosenbrock32stiff is a variable-order solver for stiff problems. It is based on the numerical differentiation formulas (NDFs). The NDFs are generally more efficient than the closely related family of backward differentiation formulas (BDFs), also known as Gear's methods. The ode15s properties let you choose among these formulas, as well as specifying the maximum order for the formula used. 
        SolverOptions.UseBackwardDifferention = True

        ' Construct the delegate representing our system of differential equations...
        Dim odeFunction As New Func(Of Double, DoubleVector, DoubleVector)(AddressOf Rigid)

        ' ...and solve. The solution is returned as a key/value pair. The first 'Key' element of the pair is
        ' the time span vector, the second 'Value' element of the pair is the corresponding solution values.
        ' That is, if the computed solution function is y then
        ' y(soln.Key(i)) = soln.Value(i)
        Dim Soln As VariableOrderOdeSolver.Solution(Of DoubleMatrix) = Solver.Solve(odeFunction, TimeSpan, y0, SolverOptions)
        Console.WriteLine("T = " & Soln.T.ToString("G5"))
        Console.WriteLine("Y = ")
        Console.WriteLine(Soln.Y.ToTabDelimited("G5"))

    End Sub

    Public Function soot_mass_rate(ByVal room As Integer, ByVal GER As Double, ByRef Yield() As Single, ByRef WallYield() As Double, ByRef CeilingYield() As Double, ByRef FloorYield() As Double, ByRef T As Double, ByRef mrate() As Double, ByRef mrate_wall As Double, ByRef mrate_ceiling As Double, ByRef mrate_floor As Double) As Double
        '*  ====================================================================
        '*  This function returns the value of the sum of the fuel mass loss rates
        '*  multiplied by the species generation rate for all burning objects.
        '*  edited 8 October 2007 - wade
        '*  ====================================================================
        '* this routine used when we are not using the postflashover crib model

        Dim dummy As Double
        Dim total As Double
        Dim i As Integer
        Dim syield() As Single
        ReDim syield(NumberObjects)

        If frmOptions1.optSootman.Checked = True Then 'manual entry of pre post flashover yields

            For i = 1 To NumberObjects
                If Flashover = True Then

                    syield(i) = postSoot
                Else
                    If VentilationLimitFlag = True And VM2 = True Then
                        syield(i) = postSoot 'VM2 case 5
                    Else
                        syield(i) = preSoot
                    End If

                End If
            Next i
        Else 'auto
            If GER >= 0.1 Then
                'correlation of Tewarson, Jiang and Morikawa - multiplier to the preflashover value for underventilated fires
                SootFactor = 1 + SootAlpha / (Exp(2.5 * GER ^ (-SootEpsilon))) '13/3/2003
            Else
                SootFactor = 1
            End If
        End If

        If room = fireroom Then
            For i = 1 To NumberObjects
                'If frmOptions1.optSootman.Checked = True Then 'manual entry of pre post flashover yields
                If sootmode = True Then 'manual entry of pre post flashover yields

                    dummy = mrate(i) * syield(i)
                    If useCLTmodel = True And KineticModel = False Then dummy = (mrate(1) - mrate_wall - mrate_ceiling - mrate_floor) * syield(i)
                Else 'uses object values of pre post flashover yield
                    'auto 
                    dummy = mrate(i) * Yield(i) * SootFactor 'postflashover
                    If useCLTmodel = True And KineticModel = False Then dummy = (mrate(1) - mrate_wall - mrate_ceiling - mrate_floor) * Yield(i) * SootFactor 'postflashover
                End If

                If i = 1 Then
                    'add in contributions from lining materials
                    dummy = dummy + mrate_wall * WallYield(room) * SootFactor + mrate_ceiling * CeilingYield(room) * SootFactor + mrate_floor * FloorYield(room) * SootFactor
                End If

                total = total + dummy
            Next i
        Else
            'only contribution in an adjacent room is from the ceiling
            total = mrate_ceiling * CeilingYield(room) * SootFactor
        End If

        soot_mass_rate = total 'kg-species per sec

    End Function
    Public Function get_soot_yield(ByVal room As Integer, ByVal GER As Double, ByRef Yield() As Single, ByRef WallYield() As Double, ByRef CeilingYield() As Double, ByRef FloorYield() As Double, ByRef T As Double, ByRef mrate() As Double, ByRef mrate_wall As Double, ByRef mrate_ceiling As Double, ByRef mrate_floor As Double) As Double

        'Dim dummy As Double
        'Dim i As Integer
        'Dim syield() As Single
        'ReDim syield(NumberObjects)



        'If sootmode = True Then 'manual entry of pre post flashover yields
        '    SootFactor = 1
        '    If Flashover = True Then
        '        'If GER > 1 Then
        '        get_soot_yield = postSoot
        '    Else
        '        get_soot_yield = preSoot
        '    End If
        'Else
        '    'sootfactor is a multiplier to determine the postflashover yield from the preflashover yield
        '    If GER >= 0.1 Then
        '        SootFactor = 1 + SootAlpha / (Exp(2.5 * GER ^ (-SootEpsilon))) '13/3/2003
        '    Else
        '        SootFactor = 1
        '    End If

        'End If

        '    If room = fireroom Then
        '        If GER > 1 Then
        '            get_soot_yield = postSoot * SootFactor 'postflashover
        '        Else
        '            dummy = mrate(i) * Yield(i) * SootFactor 'preflashover
        '        End If
        '    Else
        '    get_soot_yield = 0
        '    End If

    End Function
	
    Sub NUMJAC(ByVal n As Integer, ByRef X(,) As Double, ByRef XJAC(,,) As Double, ByVal tim As Double)

        ' Numerical Jacobian subroutine
        ' Forward difference method
        ' Input
        '  N   = number of variables
        '  X() = current evaluation vector

        ' Output
        '  XJAC() = numerical Jacobian at X()

        ' Note: requires SUB USERFUNC

        Static FVEC1(,) As Double
        Static FVEC2(,) As Double
        ReDim FVEC1(NumberRooms + 1, n)
        ReDim FVEC2(NumberRooms + 1, n)
        Static temp, EPSM, EPS, h As Double
        Static j As Integer
        Static i, k As Integer

        'machine epsilon
        EPSM = 0.000000000000000222

        ' working epsilon
        EPS = Sqrt(EPSM)

        ' evaluate function values at X()

        Call diff_eqns(tim, X, FVEC1)

        ' compute components of Jacobian
        For k = 1 To NumberRooms
            For j = 1 To n
                temp = X(k, j)
                h = EPS * Abs(temp)
                If (h = 0) Then h = EPS
                X(k, j) = temp + h
                Call diff_eqns(tim, X, FVEC2)

                If flagstop = 1 Then GoTo 33 'problem in diff_eqns, exit gracefully

                X(k, j) = temp
                For i = 1 To n
                    XJAC(k, i, j) = (FVEC2(k, i) - FVEC1(k, i)) / h
                Next i
            Next j
        Next k

33:
        ' erase working arrays
        Erase FVEC1
        Erase FVEC2

    End Sub
	
    Sub JACOBIAN(ByVal X As Double, ByRef Y(,) As Double, ByRef DFDX(,) As Double, ByRef DFDY(,,) As Double, ByVal n As Integer)

        ' Jacobian matrix subroutine

        ' Input

        '  X   = current X value
        '  Y() = current integration vector
        '        ( 3 rows by 1 column )

        ' Output

        '  DFDX() = vector of partial derivatives of system of
        '           differential equations with respect to X
        '           ( 3 rows by 1 column )
        '  DFDY() = matrix of partial derivatives of system of
        '           differential equations with respect to Y()
        '           ( 3 rows by 3 columns )

        Static i, j As Integer

        For i = 1 To n
            For j = 1 To NumberRooms
                DFDX(j, i) = 0.0#
            Next j
        Next i

        Call NUMJAC(n, Y, DFDY, X) 'added back 5/3/07
        'Call NUMJAC2(ByVal n, Y(), DFDY(), ByVal X)

    End Sub
	
    Sub LUB(ByVal n As Integer, ByRef A(,,) As Double, ByRef index(,) As Double, ByRef X(,) As Double)

        ' LU back substitution subroutine

        ' Input

        '  N       = number of equations
        '  A()     = matrix of coefficients ( N rows by N columns )
        '  INDEX() = permutation vector ( N rows )

        ' Output

        '  X() = solution vector ( N rows by 1 column )

        Static L, i, I1, j, k As Integer
        Static s As Double

        For k = 1 To NumberRooms
            I1 = 0
            For i = 1 To n
                L = index(k, i)
                s = X(k, L)
                X(k, L) = X(k, i)
                If (I1 <> 0) Then
                    For j = I1 To i - 1
                        s = s - A(k, i, j) * X(k, j)
                    Next j
                ElseIf (s <> 0.0#) Then
                    I1 = i
                End If
                X(k, i) = s
            Next i

            For i = n To 1 Step -1
                s = X(k, i)
                If (i < n) Then
                    For j = i + 1 To n
                        s = s - A(k, i, j) * X(k, j)
                    Next j
                End If
                X(k, i) = s / A(k, i, i)
            Next i
        Next k
    End Sub
	
    Sub LUD(ByVal n As Integer, ByRef A(,,) As Double, ByRef index(,) As Double, ByRef ier As Integer)

        ' LU decomposition subroutine

        ' Input

        '  N       = number of equations
        '  A()     = matrix of coefficients ( N rows by N columns )
        '  INDEX() = permutation vector ( N rows )

        ' Output

        '  A() = LU matrix of coefficients ( N rows by N columns )
        '  IER = error flag
        '    0 = no error
        '    1 = singular matrix or factorization not possible

        Static SCALE1() As Double
        ReDim SCALE1(n)
        Static k, i, j, m As Integer
        Static s As Double
        Static pivot, PIVOTMAX, tmp As Double
        Static imax As Integer
        'Dim ROWMAX As Integer
        Static ROWMAX As Double '%%% i changed this to double because of overflow

        For k = 1 To NumberRooms 'cw
            ier = 0
            For i = 1 To n
                ROWMAX = 0
                For j = 1 To n
                    If (Abs(A(k, i, j)) > ROWMAX) Then
                        ROWMAX = Abs(A(k, i, j))
                    End If
                Next j
                ' check for singular matrix
                If (ROWMAX = 0) Then
                    ier = 1
                    GoTo EXITSUB
                Else
                    SCALE1(i) = 1 / ROWMAX
                End If
            Next i

            For j = 1 To n
                If (j > 1) Then
                    For i = 1 To j - 1
                        s = A(k, i, j)
                        If (i > 1) Then
                            For m = 1 To i - 1
                                s = s - A(k, i, m) * A(k, m, j)
                            Next m
                            A(k, i, j) = s
                        End If
                    Next i
                End If

                PIVOTMAX = 0

                For i = j To n
                    s = A(k, i, j)
                    If (j > 1) Then
                        For m = 1 To j - 1
                            s = s - A(k, i, m) * A(k, m, j)
                        Next m
                        A(k, i, j) = s
                    End If
                    pivot = SCALE1(i) * Abs(s)
                    If (pivot >= PIVOTMAX) Then
                        imax = i
                        PIVOTMAX = pivot
                    End If
                Next i

                If (j <> imax) Then
                    For m = 1 To n
                        tmp = A(k, imax, m)
                        A(k, imax, m) = A(k, j, m)
                        A(k, j, m) = tmp
                    Next m
                    SCALE1(imax) = SCALE1(j)
                End If

                index(k, j) = imax

                If (j <> n) Then
                    ' check for singular matrix
                    If (A(k, j, j) = 0) Then
                        ier = 1
                        GoTo EXITSUB
                    End If
                    tmp = 1 / A(k, j, j)
                    For i = j + 1 To n
                        A(k, i, j) = A(k, i, j) * tmp
                    Next i
                End If
            Next j

            If (A(k, n, n) = 0) Then ier = 1
        Next k

EXITSUB:
        Erase SCALE1

    End Sub
	
	
	Public Function derivs_UVol(ByVal dpdt As Double, ByVal pressure As Double, ByVal uvol As Double, ByVal Enthalpy As Double, ByVal gamma_u As Double) As Double
		
		derivs_UVol = ((gamma_u - 1) * Enthalpy - uvol * dpdt / 1000) / ((pressure / 1000 + Atm_Pressure) * gamma_u)
		
	End Function
	
	Public Function derivs_Temp(ByVal SpecificHeat_layer As Double, ByVal MW_layer As Double, ByVal dpdt As Double, ByVal volume As Double, ByVal temp As Double, ByVal Enthalpy As Double, ByVal dmdt As Double, ByVal pressure As Double) As Double
		
		derivs_Temp = (Gas_Constant / MW_layer) * temp * (Enthalpy - SpecificHeat_layer * dmdt * temp + volume * dpdt / 1000) / (SpecificHeat_layer * volume * (pressure / 1000 + Atm_Pressure))
		
	End Function
	
	Public Function derivs_Pressure(ByVal room As Integer, ByVal EnthalpyU As Double, ByVal EnthalpyL As Double, ByVal gamma_u As Double) As Double

        derivs_Pressure = (gamma_u - 1) * (EnthalpyU + EnthalpyL) / RoomVolume(room) 'kW/m3=kJ/m3.s=kN/m2s=kPa/s bug prior to 2008.3
        'derivs_Pressure = (EnthalpyU + EnthalpyL) / RoomVolume(room) * (gamma_u - 1) * 1000 'Pa/s 

    End Function

    Public Function derivs_species_U(ByVal mass As Double, ByVal dmdt As Double, ByVal rate_Renamed As Double, ByVal prod As Double, ByVal species_U As Double, ByVal species_L As Double, ByVal mplume As Double) As Double
        On Error GoTo specieshandler
        derivs_species_U = (mplume * species_L + rate_Renamed + prod - species_U * dmdt) / mass
        Exit Function

specieshandler:
        MsgBox(ErrorToString(Err.Number) & " in derivs_species_U")
        ERRNO = Err.Number

        Exit Function

    End Function


    Public Sub mass_rate(ByVal T As Double, ByRef mrate() As Double, ByRef mrate_wall As Double, ByRef mrate_ceiling As Double, ByRef mrate_floor As Double)

        Dim i As Integer
        Dim QCeiling, QWall, QFloor As Single

        If Flashover = True And g_post = True Then

            If useCLTmodel = True And IntegralModel = True Then
                mrate(1) = MassLoss_ObjectwithCLT(1, T, Qburner, mrate_floor, mrate_wall, mrate_ceiling) 'spearpoint quintiere
            ElseIf useCLTmodel = True And KineticModel = True Then
                mrate(1) = MassLoss_ObjectwithCLT(1, T, Qburner, mrate_floor, mrate_wall, mrate_ceiling)
            Else
                mrate(1) = MassLoss_Object(1, T, Qburner, QFloor, QWall, QCeiling)
            End If

        Else
            For i = 1 To NumberObjects

                mrate(i) = MassLoss_Object(i, T, Qburner, QFloor, QWall, QCeiling)

                If i = burner_id And frmOptions1.optRCNone.Checked = False Then
                    'If Qburner > 0 Then mrate(i) = Qburner / (EnergyYield(i) * 1000) 'kg/sec
                    If Qburner > 0 Then mrate(i) = Qburner / (NewEnergyYield(i) * 1000) 'kg/sec

                    'limit the theoretical wall hrr to be no more than twice the actual heat release rate in the room
                    'in the absence of a proper pyrolysis model

                    'commented out the following three lines 28/2/2002
                    'If QWall > HeatRelease(fireroom, stepcount, 2) Then QWall = HeatRelease(fireroom, stepcount, 2)
                    'If QCeiling > HeatRelease(fireroom, stepcount, 2) Then QCeiling = HeatRelease(fireroom, stepcount, 2)
                    'If QFloor > HeatRelease(fireroom, stepcount, 2) Then QFloor = HeatRelease(fireroom, stepcount, 2)

                    If WallEffectiveHeatofCombustion(fireroom) > 0 Then mrate_wall = QWall / (WallEffectiveHeatofCombustion(fireroom) * 1000)
                    If CeilingEffectiveHeatofCombustion(fireroom) > 0 Then mrate_ceiling = QCeiling / (CeilingEffectiveHeatofCombustion(fireroom) * 1000)
                    If FloorEffectiveHeatofCombustion(fireroom) > 0 Then mrate_floor = QFloor / (FloorEffectiveHeatofCombustion(fireroom) * 1000)

                ElseIf i = 1 And useCLTmodel = True Then
                    If KineticModel = True Then
                        If CeilingEffectiveHeatofCombustion(fireroom) > 0 Then
                            mrate_ceiling = CeilingWoodMLR_tot(stepcount)  'kg/s
                            QCeiling = mrate_ceiling * CeilingEffectiveHeatofCombustion(fireroom) * 1000
                        End If
                        If WallEffectiveHeatofCombustion(fireroom) > 0 Then
                            'mrate_wall =( UwallWoodMLR_tot(stepcount) * upperwallarea+LwallWoodMLR_tot(stepcount) * lowerwallarea) * WallThickness(fireroom) / 1000 'kg/s
                        End If

                    Else
                        If WallEffectiveHeatofCombustion(fireroom) > 0 Then mrate_wall = QWall / (WallEffectiveHeatofCombustion(fireroom) * 1000)
                        If CeilingEffectiveHeatofCombustion(fireroom) > 0 Then mrate_ceiling = QCeiling / (CeilingEffectiveHeatofCombustion(fireroom) * 1000)
                    End If

                    'If mrate_wall > 1 Then Stop
                End If
            Next i
        End If

    End Sub

    Public Sub mass_rate_withfuelresponse(ByVal T As Double, ByRef mrate() As Double, ByRef mrate_wall As Double, ByRef mrate_ceiling As Double, ByRef mrate_floor As Double, ByVal O2lower As Double, ByVal ltemp As Double, ByVal mplume As Double, ByVal incidentflux As Double, ByRef burningrate As Double)

        Dim i As Integer
        Dim dummy As Double

        If Flashover = True And g_post = True Then

            mrate(1) = MassLoss_ObjectwithFuelResponse(1, T, Qburner, O2lower, ltemp, mplume, incidentflux, burningrate, 0, 0)

        Else
            If useCLTmodel = True Then 'must use kinetic submodel with CLT.

                'just need the MLR for wall and ceiling
                If useCLTmodel = True And IntegralModel = True Then

                    'mrate(1) = MassLoss_ObjectwithCLT(1, T, Qburner, mrate_floor, mrate_wall, mrate_ceiling) 'spearpoint quintiere
                ElseIf useCLTmodel = True And KineticModel = True Then
                    'mrate(1) = MassLoss_ObjectwithCLT(1, T, Qburner, mrate_floor, mrate_wall, mrate_ceiling)
                    dummy = MassLoss_surfaceonly_Kinetic(T, mrate_wall, mrate_ceiling)
                    'If mrate_wall > 0 Then Stop
                Else
                    'mrate(1) = MassLoss_Object(1, T, Qburner, QFloor, QWall, QCeiling)
                End If

            End If

            For i = 1 To NumberObjects

                mrate(i) = MassLoss_ObjectwithFuelResponse(i, T, Qburner, O2lower, ltemp, mplume, incidentflux, burningrate, mrate_wall, mrate_ceiling)


                'If i = burner_id And frmOptions1.optRCNone.Checked = False Then
                '    If Qburner > 0 Then mrate(i) = Qburner / (NewEnergyYield(i) * 1000) 'kg/sec

                '    If WallEffectiveHeatofCombustion(fireroom) > 0 Then mrate_wall = QWall1 / (WallEffectiveHeatofCombustion(fireroom) * 1000)
                '    If CeilingEffectiveHeatofCombustion(fireroom) > 0 Then mrate_ceiling = QCeiling1 / (CeilingEffectiveHeatofCombustion(fireroom) * 1000)
                '    If FloorEffectiveHeatofCombustion(fireroom) > 0 Then mrate_floor = QFloor1 / (FloorEffectiveHeatofCombustion(fireroom) * 1000)
                'End If

            Next i
            ' If T = 183 Then Stop


        End If

    End Sub

    Public Function species_mass_rate(ByVal room As Integer, ByRef Yield() As Single, ByRef WallYield() As Double, ByRef CeilingYield() As Double, ByRef FloorYield() As Double, ByRef T As Double, ByRef mrate() As Double, ByRef mrate_wall As Double, ByRef mrate_ceiling As Double, ByRef mrate_floor As Double) As Double
        '*  ====================================================================
        '*  This function returns the value of the sum of the fuel mass loss rates
        '*  multiplied by the species generation rate for all burning objects.
        '*  ====================================================================

        Dim dummy As Double
        Dim total As Double
        Dim i As Integer

        If room = fireroom Then
            For i = 1 To NumberObjects
                dummy = mrate(i) * Yield(i)
                If i = 1 Then
                    If useCLTmodel = True And KineticModel = False Then dummy = (mrate(i) - mrate_wall - mrate_ceiling - mrate_floor) * Yield(i)
                    dummy = dummy + mrate_wall * WallYield(room) + mrate_ceiling * CeilingYield(room) + mrate_floor * FloorYield(room)
                    End If
                    total = total + dummy
            Next i
        Else
            total = mrate_ceiling * CeilingYield(room)
        End If
        species_mass_rate = total 'kg-species per sec

    End Function

    Public Function derivs_species_L(ByVal species_mass As Double, ByVal dmdt As Double, ByVal prod As Double, ByVal species_L As Double, ByVal species_U As Double, ByVal mplume As Double) As Double
        On Error GoTo specieshandler
        derivs_species_L = (prod - mplume * species_L - species_L * dmdt) / species_mass
        Exit Function

specieshandler:
        MsgBox(ErrorToString(Err.Number) & " in derivs_species_L")
        ERRNO = Err.Number
        Exit Function
    End Function



    Public Sub door_mixing_flow(ByVal NProd As Integer, ByVal nelev As Integer, ByVal m As Integer, ByVal i As Integer, ByVal j As Integer, ByRef conu(,) As Double, ByRef prod2U() As Double, ByRef prod2L() As Double, ByRef prod1U() As Double, ByRef prod1L() As Double, ByRef TU() As Double, ByRef TL() As Double, ByRef cpu() As Double, ByRef cpl() As Double, ByRef e1u As Double, ByRef e1l As Double, ByRef e2u As Double, ByRef e2l As Double, ByRef r1u As Double, ByRef r1l As Double, ByRef r2l As Double, ByRef r2u As Double, ByRef ylay() As Double, ByRef UFLW2(,,) As Double, ByRef dp1m2() As Double, ByRef yelev() As Double)
        '---------------------------------------------
        ' 28 January 2008
        '---------------------------------------------
        Dim neutralplane As Double
        Dim L, h As Integer
        Dim doormix(2) As Double

        'door mixing vent
        If UFLW2(1, 1, 1) > 0 And UFLW2(1, 1, 2) < 0 Then
            'flow from the upper layer to the lower layer of room 1
            'find neutral plane
            For L = 1 To nelev
                If dp1m2(L) < gcd_Machine_Error And dp1m2(L) > -gcd_Machine_Error Then
                    neutralplane = yelev(L)
                    Exit For
                End If
            Next
            If neutralplane <= ylay(2) And neutralplane >= ylay(1) Then
                If j <> NumberRooms + 1 Then
                    doormix(1) = 0.5 * TL(2) / TU(1) * (1 - ylay(1) / (neutralplane)) * (WallLength2(i, j, m) / VentWidth(i, j, m)) ^ 0.25 * (r1l)
                Else
                    doormix(1) = 0.5 * TL(2) / TU(1) * (1 - ylay(1) / (neutralplane)) * (4) ^ 0.25 * (r1l)
                End If
                r1u = r1u - doormix(1)
                r1l = r1l + doormix(1)
                e1u = e1u - 1000 * cpu(1) * doormix(1) * TU(1)
                e1l = e1l + 1000 * cpu(1) * doormix(1) * TU(1)
                For h = 3 To NProd + 2
                    prod1U(h) = prod1U(h) - conu(h - 2, 1) * doormix(1)
                    prod1L(h) = prod1L(h) + conu(h - 2, 1) * doormix(1)
                Next h
            End If
            '%%% check the correlation i, j.??
        ElseIf UFLW2(2, 1, 1) > 0 And UFLW2(2, 1, 2) < 0 And j <= NumberRooms Then
            'flow from the upper layer to the lower layer of room 2
            'find neutral plane
            For L = 1 To nelev
                If dp1m2(L) < gcd_Machine_Error And dp1m2(L) > -gcd_Machine_Error Then
                    neutralplane = yelev(L)
                    Exit For
                End If
            Next
            If neutralplane <= ylay(1) And neutralplane >= ylay(2) Then
                doormix(2) = 0.5 * TL(1) / TU(2) * (1 - ylay(2) / (neutralplane)) * (WallLength1(i, j, m) / VentWidth(i, j, m)) ^ 0.25 * (r2l)
                r2u = r2u - doormix(2)
                r2l = r2l + doormix(2)
                e2u = e2u - 1000 * cpu(2) * doormix(2) * TU(2)
                e2l = e2l + 1000 * cpu(2) * doormix(2) * TU(2)
                For h = 3 To NProd + 2
                    prod2U(h) = prod2U(h) - conu(h - 2, 2) * doormix(2)
                    prod2L(h) = prod2L(h) + conu(h - 2, 2) * doormix(2)
                Next h
            End If
        End If

    End Sub
    Public Sub near_vent_mixing(ByVal NProd As Integer, ByVal nelev As Integer, ByVal m As Integer, ByVal i As Integer, ByVal j As Integer, ByRef conu(,) As Double, ByRef prod2U() As Double, ByRef prod2L() As Double, ByRef prod1U() As Double, ByRef prod1L() As Double, ByRef TU() As Double, ByRef TL() As Double, ByRef cpu() As Double, ByRef cpl() As Double, ByRef e1u As Double, ByRef e1l As Double, ByRef e2u As Double, ByRef e2l As Double, ByRef r1u As Double, ByRef r1l As Double, ByRef r2l As Double, ByRef r2u As Double, ByRef ylay() As Double, ByRef UFLW2(,,) As Double, ByRef dp1m2() As Double, ByRef yelev() As Double)
        '---------------------------------------------
        ' 28 April 2013
        ' based on near vent mixing correlation
        ' Y. Utiskul, “Theoretical and Experimental Study on Fully-Developed Compartment Fires” 
        ' NIST GCR 07-907, National Institute of Standards and Technology, Gaithersburg, MD (2007).
        '---------------------------------------------
        Dim neutralplane As Double
        Dim L, h As Integer
        Dim doormix(2) As Double
        Dim S As Double = VentSillHeight(i, j, m)
        Dim W As Double = VentWidth(i, j, m)
        Dim RHS As Double = 0
        Dim mass_in As Double = 0

        'slit vents
        'If FuelResponseEffects = True Then
        '    ' If ylay(1) <= 0.005 And UFLW2(1, 1, 1) > 0 Then
        '    If UFLW2(1, 1, 1) > 0 And m = 1 Then
        '        doormix(1) = 3.2 * UFLW2(1, 1, 1)
        '        'mass flows
        '        r1u = r1u - doormix(1) 'remove mass from upper layer
        '        r1l = r1l + doormix(1) 'add mass to lower layer

        '        'enthalpy flows
        '        e1u = e1u - 1000 * cpu(1) * doormix(1) * TU(1) 'remove energy from upper layer
        '        e1l = e1l + 1000 * cpu(1) * doormix(1) * TU(1) 'add energy to lower layer

        '        For h = 3 To NProd + 2
        '            prod1U(h) = prod1U(h) - conu(h - 2, 1) * doormix(1) 'remove species from upper layer
        '            prod1L(h) = prod1L(h) + conu(h - 2, 1) * doormix(1) 'add species to lower layer
        '        Next h
        '        Exit Sub
        '    End If
        'End If


        'these come from looking at the final destination of each of the vent flow slabs
        'they apply to an individual vent
        'using the appropriate vent flow deposition rules 
        'UFLW2(1, 1, 1) = mass flow to the lower layer of room 1
        'UFLW2(1, 1, 2) = mass flow to the upper layer of room 1
        'UFLW2(2, 1, 1) = mass flow to the lower layer of room 2
        'UFLW2(2, 1, 2) = mass flow to the upper layer of room 2

        'door mixing vent

        If UFLW2(1, 1, 1) > 0 And UFLW2(1, 1, 2) < 0 Then
            'mixing flow from the upper layer to the lower layer of room 1
            'find neutral plane
            For L = 1 To nelev
                If dp1m2(L) < gcd_Machine_Error And dp1m2(L) > -gcd_Machine_Error Then
                    neutralplane = yelev(L)
                    Exit For
                End If
            Next
            If neutralplane <= ylay(2) And neutralplane >= ylay(1) Then

                'If j <> NumberRooms + 1 Then
                '    'vent between rooms
                '    'doormix(1) = 0.5 * TL(2) / TU(1) * (1 - ylay(1) / (neutralplane)) * (WallLength2(i, j, m) / VentWidth(i, j, m)) ^ 0.25 * (r1l)
                'Else
                '    'vent to outside
                '    'doormix(1) = 0.5 * TL(2) / TU(1) * (1 - ylay(1) / (neutralplane)) * (4) ^ 0.25 * (r1l)
                'End If

                'we have flow from the upper layer to the lower layer of room 1
                RHS = TL(2) / TU(1) * (1 + (neutralplane - S) / W) * (neutralplane - ylay(1)) / (neutralplane - S)
                If RHS < 1.1 Then
                    doormix(1) = 1.14 * RHS * UFLW2(1, 1, 1)
                Else
                    doormix(1) = 1.28 * UFLW2(1, 1, 1)
                End If

                'mass flows
                r1u = r1u - doormix(1) 'remove mass from upper layer
                r1l = r1l + doormix(1) 'add mass to lower layer

                'enthalpy flows
                e1u = e1u - 1000 * cpu(1) * doormix(1) * TU(1) 'remove energy from upper layer
                e1l = e1l + 1000 * cpu(1) * doormix(1) * TU(1) 'add energy to lower layer

                For h = 3 To NProd + 2
                    prod1U(h) = prod1U(h) - conu(h - 2, 1) * doormix(1) 'remove species from upper layer
                    prod1L(h) = prod1L(h) + conu(h - 2, 1) * doormix(1) 'add species to lower layer
                Next h

            End If

        ElseIf UFLW2(2, 1, 1) > 0 And UFLW2(2, 1, 2) < 0 And j <= NumberRooms Then
            'mixing flow is from the upper layer to the lower layer of room 2
            'find neutral plane
            For L = 1 To nelev
                If dp1m2(L) < gcd_Machine_Error And dp1m2(L) > -gcd_Machine_Error Then
                    neutralplane = yelev(L)
                    Exit For
                End If
            Next

            If neutralplane <= ylay(1) And neutralplane >= ylay(2) Then
                'we have flow from the upper layer to the lower layer of room 2
                'doormix(2) = 0.5 * TL(1) / TU(2) * (1 - ylay(2) / (neutralplane)) * (WallLength1(i, j, m) / VentWidth(i, j, m)) ^ 0.25 * (r2l)

                RHS = TL(1) / TU(2) * (1 + (neutralplane - S) / W) * (neutralplane - ylay(2)) / (neutralplane - S)
                If RHS < 1.1 Then
                    doormix(2) = 1.14 * RHS * UFLW2(2, 1, 1)
                Else
                    doormix(2) = 1.28 * UFLW2(2, 1, 1)
                End If

                r2u = r2u - doormix(2) 'remove mass from upper layer
                r2l = r2l + doormix(2) 'add mass to lower layer

                e2u = e2u - 1000 * cpu(2) * doormix(2) * TU(2) 'remove energy from upper layer
                e2l = e2l + 1000 * cpu(2) * doormix(2) * TU(2) 'add energy to lower layer

                For h = 3 To NProd + 2
                    prod2U(h) = prod2U(h) - conu(h - 2, 2) * doormix(2) 'remove species from upper layer
                    prod2L(h) = prod2L(h) + conu(h - 2, 2) * doormix(2) 'add species to lower layer
                Next h
            End If

        ElseIf FuelResponseEffects = True Then
            'Exit Sub
            If ylay(1) < 0.99 * RoomHeight(fireroom) And UFLW2(1, 1, 1) > 0 Then 'vent flow entering lower layer
                '     If ylay(1) < 1.1 * (VentHeight(i, j, m) + S) And UFLW2(1, 1, 1) > 0 Then
                doormix(1) = 1.28 * UFLW2(1, 1, 1)
            End If

            'mass flows
            r1u = r1u - doormix(1) 'remove mass from upper layer
            r1l = r1l + doormix(1) 'add mass to lower layer

            'enthalpy flows
            e1u = e1u - 1000 * cpu(1) * doormix(1) * TU(1) 'remove energy from upper layer
            e1l = e1l + 1000 * cpu(1) * doormix(1) * TU(1) 'add energy to lower layer

            For h = 3 To NProd + 2
                prod1U(h) = prod1U(h) - conu(h - 2, 1) * doormix(1) 'remove species from upper layer
                prod1L(h) = prod1L(h) + conu(h - 2, 1) * doormix(1) 'add species to lower layer
            Next h

            'Else
            '    'Stop
            'End If

        End If
    End Sub
    Public Sub near_vent_mixing2(ByVal NProd As Integer, ByVal nelev As Integer, ByVal m As Integer, ByVal i As Integer, ByVal j As Integer, ByRef conu(,) As Double, ByRef prod2U() As Double, ByRef prod2L() As Double, ByRef prod1U() As Double, ByRef prod1L() As Double, ByRef TU() As Double, ByRef TL() As Double, ByRef cpu() As Double, ByRef cpl() As Double, ByRef e1u As Double, ByRef e1l As Double, ByRef e2u As Double, ByRef e2l As Double, ByRef r1u As Double, ByRef r1l As Double, ByRef r2l As Double, ByRef r2u As Double, ByRef ylay() As Double, ByRef UFLW2(,,) As Double, ByRef dp1m2() As Double, ByRef yelev() As Double)
        '---------------------------------------------
        ' 28 April 2013
        ' based on near vent mixing correlation
        ' Y. Utiskul, “Theoretical and Experimental Study on Fully-Developed Compartment Fires” 
        ' NIST GCR 07-907, National Institute of Standards and Technology, Gaithersburg, MD (2007).
        '---------------------------------------------
        Dim neutralplane As Double
        Dim L, h As Integer
        Dim doormix(2) As Double
        Dim S As Double = VentSillHeight(i, j, m)
        Dim W As Double = VentWidth(i, j, m)
        Dim RHS As Double = 0
        Dim mass_in As Double = 0

        'slit vents
        'If FuelResponseEffects = True Then
        '    ' If ylay(1) <= 0.005 And UFLW2(1, 1, 1) > 0 Then
        '    If UFLW2(1, 1, 1) > 0 And ylay(1) <= VentSillHeight(i, j, 1) + VentHeight(i, j, 1) Then
        'doormix(1) = 3.2 * UFLW2(1, 1, 1)
        '        'mass flows
        '        r1u = r1u - doormix(1) 'remove mass from upper layer
        '        r1l = r1l + doormix(1) 'add mass to lower layer

        '        'enthalpy flows
        '        e1u = e1u - 1000 * cpu(1) * doormix(1) * TU(1) 'remove energy from upper layer
        '        e1l = e1l + 1000 * cpu(1) * doormix(1) * TU(1) 'add energy to lower layer

        '        For h = 3 To NProd + 2
        '            prod1U(h) = prod1U(h) - conu(h - 2, 1) * doormix(1) 'remove species from upper layer
        '            prod1L(h) = prod1L(h) + conu(h - 2, 1) * doormix(1) 'add species to lower layer
        '        Next h
        '        Exit Sub
        '    End If
        'End If

        'If stepcount = 100 Then Stop

        'these come from looking at the final destination of each of the vent flow slabs
        'they apply to an individual vent
        'using the appropriate vent flow deposition rules 
        'UFLW2(1, 1, 1) = mass flow to the lower layer of room 1
        'UFLW2(1, 1, 2) = mass flow to the upper layer of room 1
        'UFLW2(2, 1, 1) = mass flow to the lower layer of room 2
        'UFLW2(2, 1, 2) = mass flow to the upper layer of room 2

        'door mixing vent
        If UFLW2(1, 1, 1) > 0 And FuelResponseEffects = False Then
            'If UFLW2(1, 1, 1) > 0 And UFLW2(1, 1, 2) < 0 Then
            'mixing flow from the upper layer to the lower layer of room 1
            'find neutral plane
            For L = 1 To nelev
                If dp1m2(L) < gcd_Machine_Error And dp1m2(L) > -gcd_Machine_Error Then
                    neutralplane = yelev(L)
                    Exit For
                End If
            Next
            If neutralplane <= ylay(2) And neutralplane >= ylay(1) Then

                'If j <> NumberRooms + 1 Then
                '    'vent between rooms
                '    'doormix(1) = 0.5 * TL(2) / TU(1) * (1 - ylay(1) / (neutralplane)) * (WallLength2(i, j, m) / VentWidth(i, j, m)) ^ 0.25 * (r1l)
                'Else
                '    'vent to outside
                '    'doormix(1) = 0.5 * TL(2) / TU(1) * (1 - ylay(1) / (neutralplane)) * (4) ^ 0.25 * (r1l)
                'End If

                'we have flow from the upper layer to the lower layer of room 1
                RHS = TL(2) / TU(1) * (1 + (neutralplane - S) / W) * (neutralplane - ylay(1)) / (neutralplane - S)
                If RHS < 1.1 Then
                    doormix(1) = 1.14 * RHS * UFLW2(1, 1, 1)
                Else
                    doormix(1) = 1.28 * UFLW2(1, 1, 1)
                End If

                'mass flows
                r1u = r1u - doormix(1) 'remove mass from upper layer
                r1l = r1l + doormix(1) 'add mass to lower layer

                'enthalpy flows
                e1u = e1u - 1000 * cpu(1) * doormix(1) * TU(1) 'remove energy from upper layer
                e1l = e1l + 1000 * cpu(1) * doormix(1) * TU(1) 'add energy to lower layer

                For h = 3 To NProd + 2
                    prod1U(h) = prod1U(h) - conu(h - 2, 1) * doormix(1) 'remove species from upper layer
                    prod1L(h) = prod1L(h) + conu(h - 2, 1) * doormix(1) 'add species to lower layer
                Next h

            End If

        ElseIf UFLW2(2, 1, 1) > 0 And FuelResponseEffects = False And j <= NumberRooms Then
            'ElseIf UFLW2(2, 1, 1) > 0 And UFLW2(2, 1, 2) < 0 And j <= NumberRooms Then
            'mixing flow is from the upper layer to the lower layer of room 2
            'find neutral plane
            For L = 1 To nelev
                If dp1m2(L) < gcd_Machine_Error And dp1m2(L) > -gcd_Machine_Error Then
                    neutralplane = yelev(L)
                    Exit For
                End If
            Next

            If neutralplane <= ylay(1) And neutralplane >= ylay(2) Then
                'we have flow from the upper layer to the lower layer of room 2
                'doormix(2) = 0.5 * TL(1) / TU(2) * (1 - ylay(2) / (neutralplane)) * (WallLength1(i, j, m) / VentWidth(i, j, m)) ^ 0.25 * (r2l)

                RHS = TL(1) / TU(2) * (1 + (neutralplane - S) / W) * (neutralplane - ylay(2)) / (neutralplane - S)
                If RHS < 1.1 Then
                    doormix(2) = 1.14 * RHS * UFLW2(2, 1, 1)
                Else
                    doormix(2) = 1.28 * UFLW2(2, 1, 1)
                End If

                r2u = r2u - doormix(2) 'remove mass from upper layer
                r2l = r2l + doormix(2) 'add mass to lower layer

                e2u = e2u - 1000 * cpu(2) * doormix(2) * TU(2) 'remove energy from upper layer
                e2l = e2l + 1000 * cpu(2) * doormix(2) * TU(2) 'add energy to lower layer

                For h = 3 To NProd + 2
                    prod2U(h) = prod2U(h) - conu(h - 2, 2) * doormix(2) 'remove species from upper layer
                    prod2L(h) = prod2L(h) + conu(h - 2, 2) * doormix(2) 'add species to lower layer
                Next h
            End If

        ElseIf FuelResponseEffects = True Then
            'find neutral plane
            For L = 1 To nelev
                If dp1m2(L) < gcd_Machine_Error And dp1m2(L) > -gcd_Machine_Error Then
                    neutralplane = yelev(L)
                    Exit For
                End If
            Next

            If neutralplane <= ylay(2) And neutralplane >= ylay(1) Then 'slit vents top/bottom 

                'we have flow from the upper layer to the lower layer of room 1
                RHS = TL(2) / TU(1) * (1 + (neutralplane - S) / W) * (neutralplane - ylay(1)) / (neutralplane - S)
                If RHS < 1.1 Then
                    doormix(1) = 1.14 * RHS * UFLW2(1, 1, 1)
                Else
                    doormix(1) = 1.28 * UFLW2(1, 1, 1)
                End If
                'doormix(1) = 3.2 * UFLW2(1, 1, 1) 'empirical ratio for slit vents

                'mass flows
                r1u = r1u - doormix(1) 'remove mass from upper layer
                r1l = r1l + doormix(1) 'add mass to lower layer

                'enthalpy flows
                e1u = e1u - 1000 * cpu(1) * doormix(1) * TU(1) 'remove energy from upper layer
                e1l = e1l + 1000 * cpu(1) * doormix(1) * TU(1) 'add energy to lower layer

                For h = 3 To NProd + 2
                    prod1U(h) = prod1U(h) - conu(h - 2, 1) * doormix(1) 'remove species from upper layer
                    prod1L(h) = prod1L(h) + conu(h - 2, 1) * doormix(1) 'add species to lower layer
                Next h

            Else
                'Stop
                'If VentHeight(i, j, m) > ylay(1) Then
                '    doormix(1) = 1.28 * UFLW2(1, 1, 1)
                'End If
                ''mass flows
                'r1u = r1u - doormix(1) 'remove mass from upper layer
                'r1l = r1l + doormix(1) 'add mass to lower layer

                ''enthalpy flows
                'e1u = e1u - 1000 * cpu(1) * doormix(1) * TU(1) 'remove energy from upper layer
                'e1l = e1l + 1000 * cpu(1) * doormix(1) * TU(1) 'add energy to lower layer

                'For h = 3 To NProd + 2
                '    prod1U(h) = prod1U(h) - conu(h - 2, 1) * doormix(1) 'remove species from upper layer
                '    prod1L(h) = prod1L(h) + conu(h - 2, 1) * doormix(1) 'add species to lower layer
                'Next h
            End If

        End If
    End Sub
    Public Sub door_entrain2(ByVal dteps As Double, ByRef Mass_Upper() As Double, ByRef vententrain(,,) As Double, ByVal vent As Integer, ByVal sroom As Integer, ByVal NProd As Short, ByRef prod1L() As Double, ByRef prod2L() As Double, ByRef prod2U() As Double, ByRef prod1U() As Double, ByRef conu(,) As Double, ByRef conl(,) As Double, ByRef TU() As Double, ByRef TL() As Double, ByRef cpu() As Double, ByRef cpl() As Double, ByVal room As Short, ByRef e1l As Double, ByRef e1u As Double, ByRef e2l As Double, ByRef e2u As Double, ByRef r2u As Double, ByRef r2l As Double, ByRef r1u As Double, ByRef r1l As Double, ByRef ylay() As Double, ByRef UFLW2(,,) As Double, ByRef ventfire As Double, ByRef mw_upper() As Double, ByRef mw_lower() As Double, ByRef weighted_hc As Double, ByVal nelev As Integer, ByRef dp1m2() As Double, ByRef yelev() As Double)
        '=====================================================================
        '   Subroutine to determine the entrainment in a vent flow between two rooms
        '   A modified form of the MacCaffrey Plume flow equations are used.
        '   revised 6/11/98
        '=====================================================================

        Dim zp, vp3, vp1, vp2, vp, term As Double
        Dim Qplume, vent_entrain, QEQ, T0 As Double
        Dim h As Integer
        Dim k, ventfire_new, ymax, gertoburn As Double

        T0 = TU(1) + TL(2) * (weighted_hc / deltaHair * conu(6, 1)) / (1 + (weighted_hc / deltaHair * conu(6, 1)))
        k = 13.4 * 1000 * conl(1, 2) / SpecificHeat_air / (1700 - T0)
        If weighted_hc > 0 Then gertoburn = k / (k - deltaHair / weighted_hc - 1) 'min global equivalence ratio for the upper layer to burn at the interface and external burning

        'flow passing through vent, to/from upper layer of room 1
        Dim Ms As Double = UFLW2(1, 1, 2)

        'entrainment in door vent (room = destination or room 2)
        Dim L As Integer
        Dim neutralplane As Short
        If room <= NumberRooms Then
            'flow is to an inside room
            'If UFLW2(1, 1, 2) < 0 And ylay(1) < ylay(2) Then
            If Ms < 0 And ylay(1) < ylay(2) Then

                'flow leaving upper layer of room 1
                'ventfires
                'T0 = TU(1) + TL(2) * (weighted_hc / deltaHair * conu(6, 1)) / (1 + (weighted_hc / deltaHair * conu(6, 1)))
                'k = 13.4 * 1000 * conl(1, 2) / SpecificHeat_air / (1700 - T0)
                'If weighted_hc > 0 Then gertoburn = k / (k - deltaHair / weighted_hc - 1) 'min global equivalence ratio for the upper layer to burn at the interface and external burning

                If UFLW2(2, 8, 2) <= 0.00001 Then 'unburned hydrocarbons entering upper layer of room 2
                    ventfire = 0
                Else
                    If GlobalER(stepcount) > gertoburn Then
                        ventfire = UFLW2(2, 8, 2) * weighted_hc * 1000 'kW theoretical heat release of vent fire
                    End If
                    If ventfire < 0.001 Then
                        ventfire = 0
                    End If
                End If

                'flow is from room 1 to room 2
                'door jet will entrain air from the lower layer of the adjacent room into the upper layer
                'enthalpy flux associated with the mass flow

tryagain1:
                QEQ = (cpu(1) * TU(1) - cpl(2) * TL(2)) * Abs(UFLW2(2, 1, 2)) + ventfire 'kw


                If QEQ < 0 Then QEQ = 0
                'virtual point sources
                If QEQ > 0 Then
                    vp1 = (90.9 * Abs(UFLW2(2, 1, 2)) / QEQ) ^ 1.76
                    vp2 = (38.5 * Abs(UFLW2(2, 1, 2)) / QEQ) ^ 1.001
                    vp3 = (8.1 * Abs(UFLW2(2, 1, 2)) / QEQ) ^ 0.528
                End If

                If vp1 > 0 And vp1 <= 0.08 Then
                    vp = vp1
                ElseIf vp2 > 0.08 And vp2 <= 0.2 Then
                    vp = vp2
                ElseIf vp3 > 0.2 Then
                    vp = vp3
                Else
                    'vp = 0
                    vp = 0.2
                End If

                'height of "plume"
                If QEQ > 0 Then
                    ymax = WallLength1(sroom, room, vent) / Tan(30 * PI / 180)
                    If ylay(2) - ylay(1) < ymax Then
                        zp = (ylay(2) - ylay(1)) / QEQ ^ 0.4 + vp
                    Else
                        'limit the entrainment height when plume (30 deg to vertical intercepts the opposite wall of the shaft)
                        zp = (ymax) / QEQ ^ 0.4 + vp
                    End If

                    term = zp
                End If

                'maccaffreys correlations
                If term >= 0 And term < 0.08 Then
                    vent_entrain = QEQ * 0.011 * term ^ 0.566
                ElseIf term >= 0.08 And term < 0.2 Then
                    vent_entrain = QEQ * 0.026 * term ^ 0.909
                Else
                    vent_entrain = QEQ * 0.124 * term ^ 1.895 'BUOYANT
                End If

                If TU(1) > TU(2) + dteps Then
                    If vent_entrain > QEQ / ((cpu(2) * TU(2) - cpl(2) * TL(2))) Then
                        vent_entrain = QEQ / ((cpu(2) * TU(2) - cpl(2) * TL(2)))
                        'based on eqn 46 NIST TN 1299
                        'not sure if its needed for door entrainment term
                    End If
                Else
                    vent_entrain = 0
                End If

                If ventfire > 0 Then
                    'ventfire_new = Oxygen_Limit(ventfire, vent_entrain, TU(2), conu(1, 2), conl(1, 2))
                    ventfire_new = O2_limit_cfast(room, Mass_Upper(room), ventfire, vent_entrain, TU(2), conu(1, 2), conl(1, 2), mw_upper(room), mw_lower(room), Qplume)
                    If Abs(ventfire_new - ventfire) / ventfire > 0.001 Then
                        ventfire = ventfire_new
                        GoTo tryagain1
                    End If
                End If

                vententrain(sroom, room, vent) = vent_entrain

                For h = 3 To NProd + 2
                    prod2U(h) = prod2U(h) + conl(h - 2, 2) * vent_entrain
                    prod2L(h) = prod2L(h) - conl(h - 2, 2) * vent_entrain
                Next h

                r2u = r2u + vent_entrain
                r2l = r2l - vent_entrain
                e2u = e2u + 1000 * cpl(2) * vent_entrain * TL(2) 'W
                e2l = e2l - 1000 * cpl(2) * vent_entrain * TL(2)

                e2u = e2u + (1 - NewRadiantLossFraction(1)) * ventfire * 1000 '19 june 2003, 2003.1

                'ElseIf r1u < 0 And ylay(2) < ylay(1) Then
            ElseIf UFLW2(1, 1, 2) < 0 And ylay(2) < ylay(1) Then
                'flow from 1-2 but no entrainment of lower layer
                'TUHC can burn in upper layer of 2
                'ventfires
                vent_entrain = 0
                'If prod2U(8) <= 0.00001 Then
                If UFLW2(2, 8, 2) <= 0.00001 Then
                    ventfire = 0
                Else
                    T0 = TU(1) + TU(2) * (weighted_hc / deltaHair * conu(6, 1)) / (1 + (weighted_hc / deltaHair * conu(6, 1)))
                    k = 13.4 * 1000 * conu(1, 2) / SpecificHeat_air / (1700 - T0)
                    If weighted_hc > 0 Then gertoburn = k / (k - deltaHair / weighted_hc - 1) 'min global equivalence ratio for the upper layer to burn at the interface and external burning
                    If GlobalER(stepcount) > gertoburn Then
                        ventfire = UFLW2(2, 8, 2) * weighted_hc * 1000 'kW theoretical heat release of vent fire
                    End If
                    If ventfire < 0.001 Then
                        ventfire = 0
                    Else
                        ventfire = O2_limit_cfast(room, Mass_Upper(room), ventfire, vent_entrain, TU(2), conu(1, 2), conl(1, 2), mw_upper(room), mw_lower(room), Qplume)
                    End If
                End If
                'e2u = e2u + (1 - RadiantLossFraction) * ventfire * 1000
                e2u = e2u + (1 - NewRadiantLossFraction(1)) * ventfire * 1000

                'ElseIf r1u > 0 And ylay(2) > ylay(1) Then
            ElseIf UFLW2(1, 1, 2) > 0 And ylay(2) > ylay(1) Then
                'flow from 2-1 but no entrainment of lower layer
                'TUHC can burn in upper layer of 1
                'ventfires
                vent_entrain = 0
                'If prod1U(8) <= 0.00001 Then
                If UFLW2(1, 8, 2) <= 0.00001 Then
                    ventfire = 0
                Else
                    T0 = TU(2) + TU(1) * (weighted_hc / deltaHair * conu(6, 2)) / (1 + (weighted_hc / deltaHair * conu(6, 2)))
                    k = 13.4 * 1000 * conu(1, 1) / SpecificHeat_air / (1700 - T0)
                    If weighted_hc > 0 Then gertoburn = k / (k - deltaHair / weighted_hc - 1) 'min global equivalence ratio for the upper layer to burn at the interface and external burning
                    If GlobalER(stepcount) > gertoburn Then
                        'ventfire = prod1U(8) * weighted_hc * 1000 'kW theoretical heat release of vent fire
                        ventfire = UFLW2(1, 8, 2) * weighted_hc * 1000 'kW theoretical heat release of vent fire
                    End If
                    If ventfire < 0.001 Then
                        ventfire = 0
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object O2_limit_cfast(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        ventfire = O2_limit_cfast(room, Mass_Upper(room), ventfire, vent_entrain, TU(1), conu(1, 1), conl(1, 1), mw_upper(room), mw_lower(room), Qplume)
                    End If
                End If
                e1u = e1u + (1 - NewRadiantLossFraction(1)) * ventfire * 1000

                'ElseIf r1u > 0 And ylay(1) > ylay(2) Then
            ElseIf UFLW2(1, 1, 2) > 0 And ylay(1) > ylay(2) Then
                'flow from room 2 to 1, entrains air from lower layer room 1 to upper layer room 1
                'ventfires
                'If prod1U(8) <= 0.00001 Then
                If UFLW2(1, 8, 2) <= 0.00001 Then
                    ventfire = 0
                Else
                    T0 = TU(2) + TL(1) * (weighted_hc / deltaHair * conu(6, 2)) / (1 + (weighted_hc / deltaHair * conu(6, 2)))
                    k = 13.4 * 1000 * conl(1, 1) / SpecificHeat_air / (1700 - T0)
                    If weighted_hc > 0 Then gertoburn = k / (k - deltaHair / weighted_hc - 1) 'min global equivalence ratio for the upper layer to burn at the interface and external burning
                    If GlobalER(stepcount) > gertoburn Then
                        'ventfire = prod1U(8) * weighted_hc * 1000 'kW theoretical heat release of vent fire
                        ventfire = UFLW2(1, 8, 2) * weighted_hc * 1000 'kW theoretical heat release of vent fire
                    End If
                End If
tryagain2:
                'flow is from room 2 to room 1
                'enthalpy flux associated with the mass flow
                QEQ = SpecificHeat_air * (TU(2) - TL(1)) * Abs(UFLW2(1, 1, 2)) + ventfire
                If QEQ > 0 Then
                    vp1 = (90.9 * Abs(UFLW2(1, 1, 2)) / QEQ) ^ 1.76
                    vp2 = (38.5 * Abs(UFLW2(1, 1, 2)) / QEQ) ^ 1.001
                    vp3 = (8.1 * Abs(UFLW2(1, 1, 2)) / QEQ) ^ 0.528
                End If

                If vp1 > 0 And vp1 <= 0.08 Then
                    vp = vp1
                ElseIf vp2 > 0.08 And vp2 <= 0.2 Then
                    vp = vp2
                ElseIf vp3 > 0.2 Then
                    vp = vp3
                Else
                    'vp = 0
                    vp = 0.2
                End If

                If QEQ > 0 Then
                    ymax = WallLength1(sroom, room, vent) / Tan(30 * PI / 180)
                    If ylay(1) - ylay(2) < ymax Then
                        zp = (ylay(1) - ylay(2)) / QEQ ^ 0.4 + vp
                    Else
                        'limit the entrainment height when plume (30 deg to vertical intercepts the opposite wall of the shaft)
                        zp = (ymax) / QEQ ^ 0.4 + vp
                    End If
                    'term = zp / QEQ ^ 0.4
                    term = zp
                End If

                If term >= 0 And term < 0.08 Then
                    vent_entrain = QEQ * 0.011 * term ^ 0.566 'flaming
                ElseIf term >= 0.08 And term < 0.2 Then
                    vent_entrain = QEQ * 0.026 * term ^ 0.909 'intermittent
                Else
                    vent_entrain = QEQ * 0.124 * term ^ 1.895 'buoyant plume
                End If

                If TU(2) > TU(1) + dteps Then
                    'If TU(1) > TL(1) + dteps Then
                    If vent_entrain > QEQ / (SpecificHeat_air * (TU(1) - TL(1))) Then
                        vent_entrain = QEQ / (SpecificHeat_air * (TU(1) - TL(1)))
                    End If
                Else
                    vent_entrain = 0
                End If

                If ventfire > 0.001 Then
                    ventfire_new = O2_limit_cfast(room, Mass_Upper(room), ventfire, vent_entrain, TU(1), conu(1, 1), conl(1, 1), mw_upper(room), mw_lower(room), Qplume)
                    If Abs(ventfire_new - ventfire) / ventfire > 0.001 Then
                        ventfire = ventfire_new
                        GoTo tryagain2
                    End If
                End If

                'VentEntrain(sroom, room, vent) = vent_entrain
                vententrain(room, sroom, vent) = vent_entrain

                r1u = r1u + vent_entrain
                r1l = r1l - vent_entrain
                e1u = e1u + 1000 * cpl(1) * vent_entrain * TL(1)
                e1l = e1l - 1000 * cpl(1) * vent_entrain * TL(1)

                e1u = e1u + (1 - NewRadiantLossFraction(1)) * ventfire * 1000
                For h = 3 To NProd + 2
                    prod1U(h) = prod1U(h) + conl(h - 2, 1) * vent_entrain 'kg species/s
                    prod1L(h) = prod1L(h) - conl(h - 2, 1) * vent_entrain
                Next h
                'ElseIf r2u > 0 And r1u = 0 And r1l < 0 Then
            ElseIf UFLW2(2, 1, 2) > 0 And UFLW2(1, 1, 2) = 0 And UFLW2(1, 1, 1) < 0 Then
                'here we have flow from the lower layer of r1 deposited into the upper layer of r2
                'find the neutral plane
                For L = 1 To nelev
                    If dp1m2(L) = 0 Then
                        neutralplane = yelev(L)
                        Exit For
                    End If
                Next
                'Stop
                QEQ = (cpl(1) * TL(1) - cpl(2) * TL(2)) * Abs(UFLW2(2, 1, 2)) 'kw
                QEQ = Round(QEQ, 5)

                If QEQ < 0 Then QEQ = 0
                'virtual point sources
                If QEQ > 0 Then
                    vp1 = (90.9 * Abs(UFLW2(2, 1, 2)) / QEQ) ^ 1.76
                    vp2 = (38.5 * Abs(UFLW2(2, 1, 2)) / QEQ) ^ 1.001
                    vp3 = (8.1 * Abs(UFLW2(2, 1, 2)) / QEQ) ^ 0.528
                End If

                If vp1 > 0 And vp1 <= 0.08 Then
                    vp = vp1
                ElseIf vp2 > 0.08 And vp2 <= 0.2 Then
                    vp = vp2
                ElseIf vp3 > 0.2 Then
                    vp = vp3
                Else
                    'vp = 0
                    vp = 0.2
                End If

                'height of "plume"
                If QEQ > 0 Then
                    ymax = WallLength1(sroom, room, vent) / Tan(30 * PI / 180)
                    If ylay(2) - neutralplane < ymax Then
                        zp = (ylay(2) - neutralplane) / QEQ ^ 0.4 + vp
                    Else
                        'limit the entrainment height when plume (30 deg to vertical intercepts the opposite wall of the shaft)
                        zp = (ymax) / QEQ ^ 0.4 + vp
                    End If
                    'term = zp / QEQ ^ 0.4 'mistake in version prior to 2000.04
                    term = zp
                    If term < 0 Then term = 0
                End If
                'If room = 2 Then Stop
                'maccaffreys correlations
                If term >= 0 And term < 0.08 Then
                    vent_entrain = QEQ * 0.011 * term ^ 0.566
                ElseIf term >= 0.08 And term < 0.2 Then
                    vent_entrain = QEQ * 0.026 * term ^ 0.909
                Else
                    vent_entrain = QEQ * 0.124 * term ^ 1.895
                End If

                'If TU(2) > TL(2) + dteps Then
                If TU(1) > TU(2) + dteps Then
                    If vent_entrain > QEQ / ((cpu(2) * TU(2) - cpl(2) * TL(2))) Then
                        vent_entrain = QEQ / ((cpu(2) * TU(2) - cpl(2) * TL(2)))
                    End If
                Else
                    vent_entrain = 0
                End If

                vententrain(sroom, room, vent) = vent_entrain

                For h = 3 To NProd + 2
                    prod2U(h) = prod2U(h) + conl(h - 2, 2) * vent_entrain
                    prod2L(h) = prod2L(h) - conl(h - 2, 2) * vent_entrain
                Next h

                r2u = r2u + vent_entrain
                r2l = r2l - vent_entrain
                e2u = e2u + 1000 * cpl(2) * vent_entrain * TL(2) 'W
                e2l = e2l - 1000 * cpl(2) * vent_entrain * TL(2)

                e2u = e2u + (1 - NewRadiantLossFraction(1)) * ventfire * 1000 '19 june 2003, 2003.1

            End If
        Else
            'ventfires to outside
            If stepcount > 1 Then
                T0 = TU(1) + TL(2) * (weighted_hc / deltaHair * conu(6, 1)) / (1 + (weighted_hc / deltaHair * conu(6, 1)))
                k = 13.4 * 1000 * conl(1, 2) / SpecificHeat_air / (1700 - T0)
                If weighted_hc > 0 Then gertoburn = k / (k - deltaHair / weighted_hc - 1) 'min global equivalence ratio for the upper layer to burn at the interface and external burning
                If GlobalER(stepcount - 1) > gertoburn Then
                    'ventfire = prod2U(8) * weighted_hc * 1000 'kW theoretical heat release of vent fire
                    ventfire = UFLW2(2, 8, 2) * weighted_hc * 1000 'kW theoretical heat release of vent fire
                    If ventfire < 0.001 Then ventfire = 0
                End If
            End If
            'If ventfire > 0 Then Stop
        End If
    End Sub

    Public Sub standard_entrain(ByVal minwidth As Single, ByVal Ms As Double, ByVal dteps As Double, ByRef Mass_Upper() As Double, ByRef vententrain(,,) As Double, ByVal vent As Integer, ByVal sroom As Integer, ByVal NProd As Short, ByRef prod1L() As Double, ByRef prod2L() As Double, ByRef prod2U() As Double, ByRef prod1U() As Double, ByRef conu(,) As Double, ByRef conl(,) As Double, ByVal TU() As Double, ByVal TL() As Double, ByVal cpu() As Double, ByVal cpl() As Double, ByVal room As Short, ByRef e1l As Double, ByRef e1u As Double, ByRef e2l As Double, ByRef e2u As Double, ByRef r2u As Double, ByRef r2l As Double, ByRef r1u As Double, ByRef r1l As Double, ByVal ylay() As Double, ByRef UFLW2(,,) As Double, ByRef ventfire As Double, ByRef mw_upper() As Double, ByRef mw_lower() As Double, ByVal weighted_hc As Double, ByVal nelev As Integer)

        Dim Qplume, vent_entrain, QEQ As Double
        Dim zp, vp3, vp1, vp2, vp, term As Double
        Dim ventfire_new, ymax As Double

tryagain1:
        QEQ = (cpu(1) * TU(1) - cpl(2) * TL(2)) * Abs(Ms) + ventfire 'kw

        If QEQ < 0 Then QEQ = 0

        'virtual point sources
        If QEQ > 0 Then
            vp1 = (90.9 * Abs(Ms) / QEQ) ^ 1.76
            vp2 = (38.5 * Abs(Ms) / QEQ) ^ 1.001
            vp3 = (8.1 * Abs(Ms) / QEQ) ^ 0.528
        End If

        If vp1 > 0 And vp1 <= 0.08 Then
            vp = vp1
        ElseIf vp2 > 0.08 And vp2 <= 0.2 Then
            vp = vp2
        ElseIf vp3 > 0.2 Then
            vp = vp3
        Else
            vp = 0.2
        End If

        'height of "plume"
        If QEQ > 0 Then
            Dim angle As Single = 30
            Dim radians As Single = angle * (PI / 180)

            ymax = minwidth / Tan(radians)
            If ylay(2) - ylay(1) < ymax Then
                zp = (ylay(2) - ylay(1)) / QEQ ^ 0.4 + vp
            Else
                'limit the entrainment height when plume (30 deg to vertical intercepts the opposite wall of the shaft)
                zp = (ymax) / QEQ ^ 0.4 + vp
            End If

            term = zp
        End If

        'maccaffreys correlations
        If term >= 0 And term < 0.08 Then
            vent_entrain = QEQ * 0.011 * term ^ 0.566
        ElseIf term >= 0.08 And term < 0.2 Then
            vent_entrain = QEQ * 0.026 * term ^ 0.909
        Else
            vent_entrain = QEQ * 0.124 * term ^ 1.895 'BUOYANT
        End If

        If TU(1) > TU(2) + dteps Then
            If vent_entrain > QEQ / ((cpu(2) * TU(2) - cpl(2) * TL(2))) Then
                vent_entrain = QEQ / ((cpu(2) * TU(2) - cpl(2) * TL(2)))
                'based on eqn 46 NIST TN 1299
            End If
        Else
            vent_entrain = 0
        End If

        If ventfire > 0 Then
            ventfire_new = O2_limit_cfast(room, Mass_Upper(room), ventfire, vent_entrain, TU(2), conu(1, 2), conl(1, 2), mw_upper(room), mw_lower(room), Qplume)
            If Abs(ventfire_new - ventfire) / ventfire > 0.001 Then
                ventfire = ventfire_new
                GoTo tryagain1
            End If
        End If

        vententrain(sroom, room, vent) = vent_entrain

        For h = 3 To NProd + 2
            prod2U(h) = prod2U(h) + conl(h - 2, 2) * vent_entrain
            prod2L(h) = prod2L(h) - conl(h - 2, 2) * vent_entrain
        Next h

        r2u = r2u + vent_entrain
        r2l = r2l - vent_entrain
        e2u = e2u + 1000 * cpl(2) * vent_entrain * TL(2) 'W
        e2l = e2l - 1000 * cpl(2) * vent_entrain * TL(2)

    End Sub
    Public Sub standard_entrain_rev(ByVal minwidth As Single, ByVal Ms As Double, ByVal dteps As Double, ByRef Mass_Upper() As Double, ByRef vententrain(,,) As Double, ByVal vent As Integer, ByVal sroom As Integer, ByVal NProd As Short, ByRef prod1L() As Double, ByRef prod2L() As Double, ByRef prod2U() As Double, ByRef prod1U() As Double, ByRef conu(,) As Double, ByRef conl(,) As Double, ByVal TU() As Double, ByVal TL() As Double, ByVal cpu() As Double, ByVal cpl() As Double, ByVal room As Short, ByRef e1l As Double, ByRef e1u As Double, ByRef e2l As Double, ByRef e2u As Double, ByRef r2u As Double, ByRef r2l As Double, ByRef r1u As Double, ByRef r1l As Double, ByVal ylay() As Double, ByRef UFLW2(,,) As Double, ByRef ventfire As Double, ByRef mw_upper() As Double, ByRef mw_lower() As Double, ByVal weighted_hc As Double, ByVal nelev As Integer)
        'when the entrainment occurs in first room
        Dim Qplume, vent_entrain, QEQ As Double
        Dim zp, vp3, vp1, vp2, vp, term As Double
        Dim ventfire_new, ymax As Double

tryagain1:
        QEQ = (cpu(2) * TU(2) - cpl(1) * TL(1)) * Abs(Ms) + ventfire 'kw

        If QEQ < 0 Then QEQ = 0

        'virtual point sources
        If QEQ > 0 Then
            vp1 = (90.9 * Abs(Ms) / QEQ) ^ 1.76
            vp2 = (38.5 * Abs(Ms) / QEQ) ^ 1.001
            vp3 = (8.1 * Abs(Ms) / QEQ) ^ 0.528
        End If

        If vp1 > 0 And vp1 <= 0.08 Then
            vp = vp1
        ElseIf vp2 > 0.08 And vp2 <= 0.2 Then
            vp = vp2
        ElseIf vp3 > 0.2 Then
            vp = vp3
        Else
            vp = 0.2
        End If

        'height of "plume"
        If QEQ > 0 Then
            Dim angle As Single = 30
            Dim radians As Single = angle * (PI / 180)

            ymax = minwidth / Tan(radians)
            If ylay(1) - ylay(2) < ymax Then
                zp = (ylay(1) - ylay(2)) / QEQ ^ 0.4 + vp
            Else
                'limit the entrainment height when plume (30 deg to vertical intercepts the opposite wall of the shaft)
                zp = (ymax) / QEQ ^ 0.4 + vp
            End If

            term = zp
        End If

        'maccaffreys correlations
        If term >= 0 And term < 0.08 Then
            vent_entrain = QEQ * 0.011 * term ^ 0.566
        ElseIf term >= 0.08 And term < 0.2 Then
            vent_entrain = QEQ * 0.026 * term ^ 0.909
        Else
            vent_entrain = QEQ * 0.124 * term ^ 1.895 'BUOYANT
        End If

        If TU(2) > TU(1) + dteps Then
            If vent_entrain > QEQ / ((cpu(1) * TU(1) - cpl(1) * TL(1))) Then
                vent_entrain = QEQ / ((cpu(1) * TU(1) - cpl(1) * TL(1)))
                'based on eqn 46 NIST TN 1299
            End If
        Else
            vent_entrain = 0
        End If

        If ventfire > 0 Then
            ventfire_new = O2_limit_cfast(sroom, Mass_Upper(sroom), ventfire, vent_entrain, TU(1), conu(1, 1), conl(1, 1), mw_upper(sroom), mw_lower(sroom), Qplume)
            If Abs(ventfire_new - ventfire) / ventfire > 0.001 Then
                ventfire = ventfire_new
                GoTo tryagain1
            End If
        End If

        vententrain(room, sroom, vent) = vent_entrain

        For h = 3 To NProd + 2
            prod1U(h) = prod1U(h) + conl(h - 2, 1) * vent_entrain
            prod1L(h) = prod1L(h) - conl(h - 2, 1) * vent_entrain
        Next h

        r1u = r1u + vent_entrain
        r1l = r1l - vent_entrain
        e1u = e1u + 1000 * cpl(1) * vent_entrain * TL(1) 'W
        e1l = e1l - 1000 * cpl(1) * vent_entrain * TL(1)

    End Sub
    Public Sub door_entrain3(ByVal dteps As Double, ByRef Mass_Upper() As Double, ByRef vententrain(,,) As Double, ByVal vent As Integer, ByVal sroom As Integer, ByVal NProd As Short, ByRef prod1L() As Double, ByRef prod2L() As Double, ByRef prod2U() As Double, ByRef prod1U() As Double, ByRef conu(,) As Double, ByRef conl(,) As Double, ByRef TU() As Double, ByRef TL() As Double, ByRef cpu() As Double, ByRef cpl() As Double, ByVal room As Short, ByRef e1l As Double, ByRef e1u As Double, ByRef e2l As Double, ByRef e2u As Double, ByRef r2u As Double, ByRef r2l As Double, ByRef r1u As Double, ByRef r1l As Double, ByRef ylay() As Double, ByRef UFLW2(,,) As Double, ByRef ventfire As Double, ByRef mw_upper() As Double, ByRef mw_lower() As Double, ByRef weighted_hc As Double, ByVal nelev As Integer, ByRef dp1m2() As Double, ByRef yelev() As Double)
        '=====================================================================
        '   Subroutine to determine the entrainment in a vent flow between two rooms
        '   plus vent-fires
        ' sroom = first room
        ' room = second room
        '=====================================================================

        Dim zp, vp3, vp1, vp2, vp, term As Double
        Dim Qplume, vent_entrain, QEQ, T0 As Double
        Dim h As Integer
        Dim k, ventfire_new, ymax, gertoburn As Double

        'flow passing through vent, entering(+)/leaving(-) upper layer of room 1
        Dim Ms As Double = UFLW2(1, 1, 2)

        'entrainment in door vent (room = destination or room 2)
        Dim L As Integer
        Dim neutralplane As Double

        If room <= NumberRooms Then
            'flow is to an inside room
            If Ms < 0 And ylay(1) < ylay(2) Then
                'flow is from room 1 to room 2 and the higher smoke layer is in room 2

                Dim YO2 As Single = conl(1, 2) ' = mass fraction of o2 in lower layer of room 2 = air stream
                Dim YF As Single = conu(6, 1) '= mass fraction of unburned fuel in upper layer of room 1 = fuel stream

                'min global equivalence ratio for the upper layer to burn at the interface and external burning
                T0 = TU(1) + TL(2) * (weighted_hc / deltaHair * YF) / (1 + (weighted_hc / deltaHair * YF))
                k = 13.4 * 1000 * YO2 / SpecificHeat_air / (1700 - T0)
                If weighted_hc > 0 Then gertoburn = k / (k - deltaHair / weighted_hc - 1)

                If Abs(UFLW2(1, 8, 2)) <= 0.00001 Then 'unburned hydrocarbons leaving upper layer of room 1
                    ventfire = 0
                Else
                    If GlobalER(stepcount) > gertoburn Then
                        ventfire = Abs(UFLW2(1, 8, 2)) * weighted_hc * 1000 'kW theoretical heat release of vent fire
                    End If
                    If ventfire < 0.001 Then
                        ventfire = 0
                    End If
                End If

                Dim minwidth = Min(RoomLength(room), RoomWidth(room))  'lesser dimension of room 2

                'door jet will entrain air from the lower layer of the adjacent room into the upper layer
                'enthalpy flux associated with the mass flow
                'standard_entrain(minwidth, Ms, dteps, Mass_Upper, vententrain, vent, sroom, NProd, prod1L, prod2L, prod2U, prod1U, conu, conl, TU, TL, cpu, cpl, room, e1l, e1u, e2l, e2u, r2u, r2l, r1u, r1l, ylay, UFLW2, ventfire, mw_upper, mw_lower, weighted_hc, nelev)

                If spillplumemodel(sroom, room, vent) = 0 Then
                    Call standard_entrain(minwidth, Ms, dteps, Mass_Upper, vententrain, vent, sroom, NProd, prod1L, prod2L, prod2U, prod1U, conu, conl, TU, TL, cpu, cpl, room, e1l, e1u, e2l, e2u, r2u, r2l, r1u, r1l, ylay, UFLW2, ventfire, mw_upper, mw_lower, weighted_hc, nelev)
                Else
                    Call spillplume_2013(minwidth, dteps, Mass_Upper, vententrain, vent, sroom, NProd, prod1L, prod2L, prod2U, prod1U, conu, conl, TU, TL, cpu, cpl, room, e1l, e1u, e2l, e2u, r2u, r2l, r1u, r1l, ylay, UFLW2, ventfire, mw_upper, mw_lower, weighted_hc, nelev, dp1m2, yelev)
                End If

                e2u = e2u + (1 - NewRadiantLossFraction(1)) * ventfire * 1000

            ElseIf Ms < 0 And ylay(2) < ylay(1) Then
                'flow from 1-2 but no entrainment of lower layer
                'TUHC can burn in upper layer of 2
                'ventfires
                vent_entrain = 0

                If Abs(UFLW2(1, 8, 2)) <= 0.00001 Then 'unburned hydrocarbons leaving upper layer of room 1
                    ventfire = 0
                Else
                    If GlobalER(stepcount) > gertoburn Then
                        'ventfire = UFLW2(2, 8, 2) * weighted_hc * 1000 'kW theoretical heat release of vent fire
                        ventfire = Abs(UFLW2(1, 8, 2)) * weighted_hc * 1000 'kW theoretical heat release of vent fire

                    End If
                    If ventfire < 0.001 Then
                        ventfire = 0
                    Else
                        ventfire = O2_limit_cfast(room, Mass_Upper(room), ventfire, vent_entrain, TU(2), conu(1, 2), conl(1, 2), mw_upper(room), mw_lower(room), Qplume)
                    End If
                End If
                e2u = e2u + (1 - NewRadiantLossFraction(1)) * ventfire * 1000

            ElseIf Ms > 0 And ylay(2) > ylay(1) Then
                'flow from 2-1 but no entrainment of lower layer
                'TUHC can burn in upper layer of 1
                'ventfires
                vent_entrain = 0

                If Abs(UFLW2(1, 8, 2)) <= 0.00001 Then
                    ventfire = 0
                Else
                    If GlobalER(stepcount) > gertoburn Then
                        ventfire = Abs(UFLW2(1, 8, 2)) * weighted_hc * 1000 'kW theoretical heat release of vent fire
                    End If
                    If ventfire < 0.001 Then
                        ventfire = 0
                    Else
                        ventfire = O2_limit_cfast(room, Mass_Upper(room), ventfire, vent_entrain, TU(1), conu(1, 1), conl(1, 1), mw_upper(room), mw_lower(room), Qplume)
                    End If
                End If
                e1u = e1u + (1 - NewRadiantLossFraction(1)) * ventfire * 1000

            ElseIf Ms > 0 And ylay(1) > ylay(2) Then
                'flow from room 2 to 1, entrains air from lower layer room 1 to upper layer room 1
                'ventfires

                If Abs(UFLW2(1, 8, 2)) <= 0.00001 Then 'unburned fuel 
                    ventfire = 0
                Else
                    'conl(1, 1) = mass fraction of o2 in lower layer of room 1 = air stream
                    'conu(6, 2) = mass fraction of unburned fuel in upper layer of room 2 = fuel stream

                    T0 = TU(2) + TL(1) * (weighted_hc / deltaHair * conu(6, 2)) / (1 + (weighted_hc / deltaHair * conu(6, 2)))
                    k = 13.4 * 1000 * conl(1, 1) / SpecificHeat_air / (1700 - T0)
                    If weighted_hc > 0 Then gertoburn = k / (k - deltaHair / weighted_hc - 1) 'min global equivalence ratio for the upper layer to burn at the interface and external burning
                    If GlobalER(stepcount) > gertoburn Then
                        ventfire = Abs(UFLW2(1, 8, 2)) * weighted_hc * 1000 'kW theoretical heat release of vent fire
                    End If
                End If
tryagain2:
                'flow is from room 2 to room 1
                'enthalpy flux associated with the mass flow
                QEQ = SpecificHeat_air * (TU(2) - TL(1)) * Abs(Ms) + ventfire
                If QEQ > 0 Then
                    vp1 = (90.9 * Abs(UFLW2(1, 1, 2)) / QEQ) ^ 1.76
                    vp2 = (38.5 * Abs(UFLW2(1, 1, 2)) / QEQ) ^ 1.001
                    vp3 = (8.1 * Abs(UFLW2(1, 1, 2)) / QEQ) ^ 0.528
                End If

                If vp1 > 0 And vp1 <= 0.08 Then
                    vp = vp1
                ElseIf vp2 > 0.08 And vp2 <= 0.2 Then
                    vp = vp2
                ElseIf vp3 > 0.2 Then
                    vp = vp3
                Else
                    vp = 0.2
                End If

                If QEQ > 0 Then
                    ymax = WallLength1(sroom, room, vent) / Tan(30 * PI / 180)
                    If ylay(1) - ylay(2) < ymax Then
                        zp = (ylay(1) - ylay(2)) / QEQ ^ 0.4 + vp
                    Else
                        'limit the entrainment height when plume (30 deg to vertical intercepts the opposite wall of the shaft)
                        zp = (ymax) / QEQ ^ 0.4 + vp
                    End If
                    term = zp
                End If

                If term >= 0 And term < 0.08 Then
                    vent_entrain = QEQ * 0.011 * term ^ 0.566 'flaming
                ElseIf term >= 0.08 And term < 0.2 Then
                    vent_entrain = QEQ * 0.026 * term ^ 0.909 'intermittent
                Else
                    vent_entrain = QEQ * 0.124 * term ^ 1.895 'buoyant plume
                End If

                If TU(2) > TU(1) + dteps Then
                    If vent_entrain > QEQ / (SpecificHeat_air * (TU(1) - TL(1))) Then
                        vent_entrain = QEQ / (SpecificHeat_air * (TU(1) - TL(1)))
                    End If
                Else
                    vent_entrain = 0
                End If

                If ventfire > 0.001 Then
                    ventfire_new = O2_limit_cfast(room, Mass_Upper(room), ventfire, vent_entrain, TU(1), conu(1, 1), conl(1, 1), mw_upper(room), mw_lower(room), Qplume)
                    If Abs(ventfire_new - ventfire) / ventfire > 0.001 Then
                        ventfire = ventfire_new
                        GoTo tryagain2
                    End If
                End If

                vententrain(room, sroom, vent) = vent_entrain

                r1u = r1u + vent_entrain
                r1l = r1l - vent_entrain
                e1u = e1u + 1000 * cpl(1) * vent_entrain * TL(1)
                e1l = e1l - 1000 * cpl(1) * vent_entrain * TL(1)

                e1u = e1u + (1 - NewRadiantLossFraction(1)) * ventfire * 1000
                For h = 3 To NProd + 2
                    prod1U(h) = prod1U(h) + conl(h - 2, 1) * vent_entrain 'kg species/s
                    prod1L(h) = prod1L(h) - conl(h - 2, 1) * vent_entrain
                Next h

            ElseIf UFLW2(2, 1, 2) > 0 And UFLW2(1, 1, 2) = 0 And UFLW2(1, 1, 1) < 0 Then
                'here we have flow from the lower layer of r1 deposited into the upper layer of r2
                'find the neutral plane
                For L = 1 To nelev
                    If dp1m2(L) = 0 Then
                        neutralplane = yelev(L)
                        Exit For
                    End If
                Next

                QEQ = (cpl(1) * TL(1) - cpl(2) * TL(2)) * Abs(UFLW2(2, 1, 2)) 'kw
                QEQ = Round(QEQ, 5)

                If QEQ < 0 Then QEQ = 0
                'virtual point sources
                If QEQ > 0 Then
                    vp1 = (90.9 * Abs(UFLW2(2, 1, 2)) / QEQ) ^ 1.76
                    vp2 = (38.5 * Abs(UFLW2(2, 1, 2)) / QEQ) ^ 1.001
                    vp3 = (8.1 * Abs(UFLW2(2, 1, 2)) / QEQ) ^ 0.528
                End If

                If vp1 > 0 And vp1 <= 0.08 Then
                    vp = vp1
                ElseIf vp2 > 0.08 And vp2 <= 0.2 Then
                    vp = vp2
                ElseIf vp3 > 0.2 Then
                    vp = vp3
                Else
                    'vp = 0
                    vp = 0.2
                End If

                'height of "plume"
                If QEQ > 0 Then
                    ymax = WallLength1(sroom, room, vent) / Tan(30 * PI / 180)
                    If ylay(2) - neutralplane < ymax Then
                        zp = (ylay(2) - neutralplane) / QEQ ^ 0.4 + vp
                    Else
                        'limit the entrainment height when plume (30 deg to vertical intercepts the opposite wall of the shaft)
                        zp = (ymax) / QEQ ^ 0.4 + vp
                    End If
                    'term = zp / QEQ ^ 0.4 'mistake in version prior to 2000.04
                    term = zp
                    If term < 0 Then term = 0
                End If

                'maccaffreys correlations
                If term >= 0 And term < 0.08 Then
                    vent_entrain = QEQ * 0.011 * term ^ 0.566
                ElseIf term >= 0.08 And term < 0.2 Then
                    vent_entrain = QEQ * 0.026 * term ^ 0.909
                Else
                    vent_entrain = QEQ * 0.124 * term ^ 1.895
                End If

                'If TU(2) > TL(2) + dteps Then
                If TU(1) > TU(2) + dteps Then
                    If vent_entrain > QEQ / ((cpu(2) * TU(2) - cpl(2) * TL(2))) Then
                        vent_entrain = QEQ / ((cpu(2) * TU(2) - cpl(2) * TL(2)))
                    End If
                Else
                    vent_entrain = 0
                End If

                vententrain(sroom, room, vent) = vent_entrain

                For h = 3 To NProd + 2
                    prod2U(h) = prod2U(h) + conl(h - 2, 2) * vent_entrain
                    prod2L(h) = prod2L(h) - conl(h - 2, 2) * vent_entrain
                Next h

                r2u = r2u + vent_entrain
                r2l = r2l - vent_entrain
                e2u = e2u + 1000 * cpl(2) * vent_entrain * TL(2) 'W
                e2l = e2l - 1000 * cpl(2) * vent_entrain * TL(2)

                e2u = e2u + (1 - NewRadiantLossFraction(1)) * ventfire * 1000 '19 june 2003, 2003.1



            End If
        Else
            'If Flashover = True Then Stop

            'ventfires to outside
            If stepcount > 1 Then
                T0 = TU(1) + TL(2) * (weighted_hc / deltaHair * conu(6, 1)) / (1 + (weighted_hc / deltaHair * conu(6, 1)))
                k = 13.4 * 1000 * conl(1, 2) / SpecificHeat_air / (1700 - T0)
                If weighted_hc > 0 Then gertoburn = k / (k - deltaHair / weighted_hc - 1) 'min global equivalence ratio for the upper layer to burn at the interface and external burning

                If GlobalER(stepcount - 1) > gertoburn Then
                    ventfire = Abs(UFLW2(1, 8, 2)) * weighted_hc * 1000 'kW theoretical heat release of vent fire
                    If ventfire < 0.001 Then ventfire = 0
                End If
            End If

        End If

    End Sub
    Public Sub door_entrain4(ByVal dteps As Double, ByRef Mass_Upper() As Double, ByRef vententrain(,,) As Double, ByVal vent As Integer, ByVal sroom As Integer, ByVal NProd As Short, ByRef prod1L() As Double, ByRef prod2L() As Double, ByRef prod2U() As Double, ByRef prod1U() As Double, ByRef conu(,) As Double, ByRef conl(,) As Double, ByRef TU() As Double, ByRef TL() As Double, ByRef cpu() As Double, ByRef cpl() As Double, ByVal room As Short, ByRef e1l As Double, ByRef e1u As Double, ByRef e2l As Double, ByRef e2u As Double, ByRef r2u As Double, ByRef r2l As Double, ByRef r1u As Double, ByRef r1l As Double, ByRef ylay() As Double, ByRef UFLW2(,,) As Double, ByRef ventfire As Double, ByRef mw_upper() As Double, ByRef mw_lower() As Double, ByRef weighted_hc As Double, ByVal nelev As Integer, ByRef dp1m2() As Double, ByRef yelev() As Double)
        '=====================================================================
        '   Subroutine to determine the entrainment in a vent flow between two rooms
        '   plus vent-fires
        ' sroom = first room
        ' room = second room
        '=====================================================================

        Dim zp, vp3, vp1, vp2, vp, term, ymax As Double
        Dim Qplume, vent_entrain, QEQ, T0 As Double
        Dim h As Integer
        Dim k, gertoburn As Double
        Dim L As Integer
        Dim neutralplane As Double

        'flow passing through vent, entering(+)/leaving(-) upper layer of room 1
        Dim Ms As Double = UFLW2(1, 1, 2)

        If room <= NumberRooms Then
            'flow is to an inside room
            If Ms < 0 And ylay(1) < ylay(2) Then
                'flow is from room 1 to room 2 and the higher smoke layer is in room 2
                'get entrainment into the upper layer of room 2

                Dim YO2 As Single = conl(1, 2) ' = mass fraction of o2 in lower layer of room 2 = air stream
                Dim YF As Single = conu(6, 1) '= mass fraction of unburned fuel in upper layer of room 1 = fuel stream

                'min global equivalence ratio for the upper layer to burn at the interface and external burning
                T0 = TU(1) + TL(2) * (weighted_hc / deltaHair * YF) / (1 + (weighted_hc / deltaHair * YF))
                k = 13.4 * 1000 * YO2 / SpecificHeat_air / (1700 - T0)
                If weighted_hc > 0 Then gertoburn = k / (k - deltaHair / weighted_hc - 1)

                'check this next bit?
                If Abs(UFLW2(1, 8, 2)) > 0.00001 Then '<0 = unburned hydrocarbons leaving upper layer of room 1
                    ventfire = 0
                Else
                    If GlobalER(stepcount) > gertoburn Then 'this only applies to vents from room of origin?
                        'ventfire in room 2
                        ventfire = Abs(UFLW2(1, 8, 2)) * weighted_hc * 1000 'kW theoretical heat release of vent fire
                    End If
                    If ventfire < 0.001 Then
                        ventfire = 0
                    End If
                End If

                Dim minwidth = Min(RoomLength(room), RoomWidth(room))  'lesser dimension of room 2

                'door jet will entrain air from the lower layer of the adjacent room into the upper layer
                'enthalpy flux associated with the mass flow

                If spillplumemodel(sroom, room, vent) = 0 Then
                    Call standard_entrain(minwidth, Ms, dteps, Mass_Upper, vententrain, vent, sroom, NProd, prod1L, prod2L, prod2U, prod1U, conu, conl, TU, TL, cpu, cpl, room, e1l, e1u, e2l, e2u, r2u, r2l, r1u, r1l, ylay, UFLW2, ventfire, mw_upper, mw_lower, weighted_hc, nelev)
                Else
                    Call spillplume_2014(minwidth, dteps, Mass_Upper, vententrain, vent, sroom, NProd, prod1L, prod2L, prod2U, prod1U, conu, conl, TU, TL, cpu, cpl, room, e1l, e1u, e2l, e2u, r2u, r2l, r1u, r1l, ylay, UFLW2, ventfire, mw_upper, mw_lower, weighted_hc, nelev, dp1m2, yelev)
                End If

                e2u = e2u + (1 - NewRadiantLossFraction(1)) * ventfire * 1000


            ElseIf Ms > 0 And ylay(1) > ylay(2) Then
                'flow is from room 2 to room 1 and the higher smoke layer is in room 1
                'get entrainment into the upper layer of room 1

                Dim YO2 As Single = conl(1, 1) ' = mass fraction of o2 in lower layer of room 1 = air stream
                Dim YF As Single = conu(6, 2) '= mass fraction of unburned fuel in upper layer of room 2 = fuel stream

                'min global equivalence ratio for the upper layer to burn at the interface and external burning
                T0 = TU(2) + TL(1) * (weighted_hc / deltaHair * YF) / (1 + (weighted_hc / deltaHair * YF))
                k = 13.4 * 1000 * YO2 / SpecificHeat_air / (1700 - T0)
                If weighted_hc > 0 Then gertoburn = k / (k - deltaHair / weighted_hc - 1)

                'check this
                If Abs(UFLW2(2, 8, 2)) > 0.00001 Then '<0= unburned hydrocarbons leaving upper layer of room 2
                    ventfire = 0
                Else
                    If GlobalER(stepcount) > gertoburn Then
                        'ventfire in room 1
                        ventfire = Abs(UFLW2(2, 8, 2)) * weighted_hc * 1000 'kW theoretical heat release of vent fire
                    End If
                    If ventfire < 0.001 Then
                        ventfire = 0
                    End If
                End If

                Dim minwidth = Min(RoomLength(sroom), RoomWidth(sroom))  'lesser dimension of room 1

                'door jet will entrain air from the lower layer of the adjacent room into the upper layer
                'enthalpy flux associated with the mass flow

                If spillplumemodel(sroom, room, vent) = 0 Then
                    Call standard_entrain_rev(minwidth, Ms, dteps, Mass_Upper, vententrain, vent, sroom, NProd, prod1L, prod2L, prod2U, prod1U, conu, conl, TU, TL, cpu, cpl, room, e1l, e1u, e2l, e2u, r2u, r2l, r1u, r1l, ylay, UFLW2, ventfire, mw_upper, mw_lower, weighted_hc, nelev)
                Else
                    Call spillplume_2014(minwidth, dteps, Mass_Upper, vententrain, vent, sroom, NProd, prod1L, prod2L, prod2U, prod1U, conu, conl, TU, TL, cpu, cpl, room, e1l, e1u, e2l, e2u, r2u, r2l, r1u, r1l, ylay, UFLW2, ventfire, mw_upper, mw_lower, weighted_hc, nelev, dp1m2, yelev)
                End If

                e1u = e1u + (1 - NewRadiantLossFraction(1)) * ventfire * 1000

            ElseIf Ms < 0 And ylay(2) < ylay(1) Then
                ''flow from 1-2 but no entrainment of lower layer

                'TUHC can burn in upper layer of 2
                'ventfires
                vent_entrain = 0

                If Abs(UFLW2(1, 8, 2)) <= 0.00001 Then 'unburned hydrocarbons leaving upper layer of room 1
                    ventfire = 0
                Else
                    If GlobalER(stepcount) > gertoburn Then
                        ventfire = Abs(UFLW2(1, 8, 2)) * weighted_hc * 1000 'kW theoretical heat release of vent fire
                    End If
                    If ventfire < 0.001 Then
                        ventfire = 0
                    Else
                        ventfire = O2_limit_cfast(room, Mass_Upper(room), ventfire, vent_entrain, TU(2), conu(1, 2), conl(1, 2), mw_upper(room), mw_lower(room), Qplume)
                    End If
                End If

                e2u = e2u + (1 - NewRadiantLossFraction(1)) * ventfire * 1000

            ElseIf Ms > 0 And ylay(2) > ylay(1) Then
                'flow from 2-1 but no entrainment of lower layer
                'TUHC can burn in upper layer of 1
                'ventfires
                vent_entrain = 0

                If Abs(UFLW2(1, 8, 2)) <= 0.00001 Then
                    ventfire = 0
                Else
                    If GlobalER(stepcount) > gertoburn Then
                        ventfire = Abs(UFLW2(1, 8, 2)) * weighted_hc * 1000 'kW theoretical heat release of vent fire
                    End If
                    If ventfire < 0.001 Then
                        ventfire = 0
                    Else
                        ventfire = O2_limit_cfast(room, Mass_Upper(room), ventfire, vent_entrain, TU(1), conu(1, 1), conl(1, 1), mw_upper(room), mw_lower(room), Qplume)
                    End If
                End If

                e1u = e1u + (1 - NewRadiantLossFraction(1)) * ventfire * 1000

            ElseIf UFLW2(2, 1, 2) > 0 And Ms = 0 And UFLW2(1, 1, 1) < 0 Then
                'here we have flow from the lower layer of r1 deposited into the upper layer of r2
                'find the neutral plane
                For L = 1 To nelev
                    If dp1m2(L) = 0 Then
                        neutralplane = yelev(L)
                        Exit For
                    End If
                Next

                QEQ = (cpl(1) * TL(1) - cpl(2) * TL(2)) * Abs(UFLW2(2, 1, 2)) 'kw
                QEQ = Round(QEQ, 5)

                If QEQ < 0 Then QEQ = 0
                'virtual point sources
                If QEQ > 0 Then
                    vp1 = (90.9 * Abs(UFLW2(2, 1, 2)) / QEQ) ^ 1.76
                    vp2 = (38.5 * Abs(UFLW2(2, 1, 2)) / QEQ) ^ 1.001
                    vp3 = (8.1 * Abs(UFLW2(2, 1, 2)) / QEQ) ^ 0.528
                End If

                If vp1 > 0 And vp1 <= 0.08 Then
                    vp = vp1
                ElseIf vp2 > 0.08 And vp2 <= 0.2 Then
                    vp = vp2
                ElseIf vp3 > 0.2 Then
                    vp = vp3
                Else
                    'vp = 0
                    vp = 0.2
                End If

                'height of "plume"
                If QEQ > 0 Then
                    ymax = WallLength1(sroom, room, vent) / Tan(30 * PI / 180)
                    If ylay(2) - neutralplane < ymax Then
                        zp = (ylay(2) - neutralplane) / QEQ ^ 0.4 + vp
                    Else
                        'limit the entrainment height when plume (30 deg to vertical intercepts the opposite wall of the shaft)
                        zp = (ymax) / QEQ ^ 0.4 + vp
                    End If
                    'term = zp / QEQ ^ 0.4 'mistake in version prior to 2000.04
                    term = zp
                    If term < 0 Then term = 0
                End If

                'maccaffreys correlations
                If term >= 0 And term < 0.08 Then
                    vent_entrain = QEQ * 0.011 * term ^ 0.566
                ElseIf term >= 0.08 And term < 0.2 Then
                    vent_entrain = QEQ * 0.026 * term ^ 0.909
                Else
                    vent_entrain = QEQ * 0.124 * term ^ 1.895
                End If

                'If TU(2) > TL(2) + dteps Then
                If TU(1) > TU(2) + dteps Then
                    If vent_entrain > QEQ / ((cpu(2) * TU(2) - cpl(2) * TL(2))) Then
                        vent_entrain = QEQ / ((cpu(2) * TU(2) - cpl(2) * TL(2)))
                    End If
                Else
                    vent_entrain = 0
                End If

                vententrain(sroom, room, vent) = vent_entrain

                For h = 3 To NProd + 2
                    prod2U(h) = prod2U(h) + conl(h - 2, 2) * vent_entrain
                    prod2L(h) = prod2L(h) - conl(h - 2, 2) * vent_entrain
                Next h

                r2u = r2u + vent_entrain
                r2l = r2l - vent_entrain
                e2u = e2u + 1000 * cpl(2) * vent_entrain * TL(2) 'W
                e2l = e2l - 1000 * cpl(2) * vent_entrain * TL(2)

                e2u = e2u + (1 - NewRadiantLossFraction(1)) * ventfire * 1000 '19 june 2003, 2003.1

            End If
        Else
            'If Flashover = True Then Stop

            'ventfires to outside
            If stepcount > 1 Then
                T0 = TU(1) + TL(2) * (weighted_hc / deltaHair * conu(6, 1)) / (1 + (weighted_hc / deltaHair * conu(6, 1)))
                k = 13.4 * 1000 * conl(1, 2) / SpecificHeat_air / (1700 - T0)
                If weighted_hc > 0 Then gertoburn = k / (k - deltaHair / weighted_hc - 1) 'min global equivalence ratio for the upper layer to burn at the interface and external burning

                If GlobalER(stepcount - 1) > gertoburn Then
                    ventfire = Abs(UFLW2(1, 8, 2)) * weighted_hc * 1000 'kW theoretical heat release of vent fire
                    If ventfire < 0.001 Then ventfire = 0
                End If
            End If

        End If

    End Sub
    Public Function TUHC_Rate(ByVal QA As Double, ByVal totalmassloss As Double, ByVal hc As Double) As Double
        '===========================================================
        '   Function to determine the rate of change in concentration
        '   of total unburned fuel generated by fire
        '   20 November 1998
        '===========================================================
        If QA < 0 Then QA = 0
        On Error GoTo errorhandler
        If hc = 0 Then GoTo errorhandler
        TUHC_Rate = totalmassloss - QA / (hc * 1000) 'kg/s

        Exit Function
errorhandler:
        TUHC_Rate = 0
    End Function

    Public Sub Do_Stuff(ByRef i As Integer)
        '==============================================================================
        '   misc calculations
        '   called once per timestep
        '   C Wade 7/8/98
        '==============================================================================

        Dim stemp, ftemp, weighted_hc As Double
        Dim count As Integer
        Dim heatreleased As Double
        Dim lowerabsorb, upperabsorb, mplume As Double
        Dim j, id As Integer
        Dim sum, Qplume As Double
        Dim massloss() As Double
        Dim mrate_ceiling, mrate_wall, mrate_floor As Double
        Dim mrate() As Double
        Dim matc(4, 1) As Double
        ReDim massloss(NumberObjects)
        ReDim mrate(NumberObjects)
        Dim Mass_Upper, mw_upper, mw_lower, fuelmassloss As Double

        'effective molecular weight of the layers for fireroom
        mw_upper = MolecularWeightCO * COMassFraction(fireroom, stepcount, 1) + MolecularWeightCO2 * CO2MassFraction(fireroom, stepcount, 1) + MolecularWeightH2O * H2OMassFraction(fireroom, stepcount, 1) + MolecularWeightHCN * HCNMassFraction(fireroom, stepcount, 1) + MolecularWeightO2 * O2MassFraction(fireroom, stepcount, 1) + MolecularWeightN2 * (1 - O2MassFraction(fireroom, stepcount, 1) - COMassFraction(fireroom, stepcount, 1) - CO2MassFraction(fireroom, stepcount, 1) - H2OMassFraction(fireroom, stepcount, 1) - HCNMassFraction(fireroom, stepcount, 1))
        mw_lower = MolecularWeightCO * COMassFraction(fireroom, stepcount, 2) + MolecularWeightCO2 * CO2MassFraction(fireroom, stepcount, 2) + MolecularWeightH2O * H2OMassFraction(fireroom, stepcount, 2) + MolecularWeightHCN * HCNMassFraction(fireroom, stepcount, 2) + MolecularWeightO2 * O2MassFraction(fireroom, stepcount, 2) + MolecularWeightN2 * (1 - O2MassFraction(fireroom, stepcount, 2) - COMassFraction(fireroom, stepcount, 2) - CO2MassFraction(fireroom, stepcount, 2) - H2OMassFraction(fireroom, stepcount, 2) - HCNMassFraction(fireroom, stepcount, 2))

        count = 0
        'radiation flux on floor
        If i > 1 Then Target(fireroom, i) = Target(fireroom, i - 1) 'kw/m2 - for use in plume equation

        'mass of upper layer in fireroom
        Mass_Upper = UpperVolume(fireroom, i) * (Atm_Pressure + RoomPressure(fireroom, i) / 1000) / (Gas_Constant / mw_upper) / uppertemp(fireroom, i)

        'If j = fireroom Then
        If Flashover = True And g_post = True Then
            Call mass_rate(tim(i, 1), mrate, mrate_wall, mrate_ceiling, mrate_floor)
            sum = 0
            fuelmassloss = mrate(1)
            'sum = mrate(1) * HoC_fuel 'kg/s.kJ/g
            sum = mrate(1) * NewHoC_fuel 'kg/s.kJ/g
            ftemp = fuelmassloss 'kg/sec
            stemp = sum 'kg/s.kJ/g
            'fuelmassloss = fuelmassloss + mrate_wall + mrate_ceiling + mrate_floor 'kg/sec
            'sum = sum + mrate_wall * WallEffectiveHeatofCombustion(fireroom) + mrate_ceiling * FloorEffectiveHeatofCombustion(fireroom) + mrate_floor * FloorEffectiveHeatofCombustion(fireroom)
        Else
            Call mass_rate(tim(i, 1), mrate, mrate_wall, mrate_ceiling, mrate_floor)
            fuelmassloss = 0
            sum = 0
            ftemp = 0 'kg/sec
            stemp = 0
            For id = 1 To NumberObjects
                fuelmassloss = fuelmassloss + mrate(id)
                'sum = sum + mrate(id) * EnergyYield(id) 'kg/s.kJ/g
                sum = sum + mrate(id) * NewEnergyYield(id) 'kg/s.kJ/g
                ftemp = ftemp + fuelmassloss 'kg/sec
                stemp = stemp + sum 'kg/s.kJ/g
                If id = 1 Then
                    fuelmassloss = fuelmassloss + mrate_wall + mrate_ceiling + mrate_floor 'kg/sec
                    sum = sum + mrate_wall * WallEffectiveHeatofCombustion(fireroom) + mrate_ceiling * CeilingEffectiveHeatofCombustion(fireroom) + mrate_floor * FloorEffectiveHeatofCombustion(fireroom)
                End If
            Next id
        End If
        If fuelmassloss > 0 Then weighted_hc = CDec(sum / fuelmassloss) 'MJ/kg or kJ/g
        If ftemp > 0 Then stemp = CDec(stemp / ftemp) 'kJ/g
        'End If

        'determine theoretical heat release rate
        If TUHC(fireroom, i, 1) > 0.00001 Then
            'why did i need to try and burn fuel in the upper layer?   15/11/13
            'HeatRelease(fireroom, i, 1) = Composite_HRR(tim(i, 1)) + TUHC(fireroom, i, 1) * Mass_Upper * weighted_hc * 1000 * Timestep 'could burn upper layer tuhc in 1 second
            HeatRelease(fireroom, i, 1) = Composite_HRR(tim(i, 1))
        Else
            HeatRelease(fireroom, i, 1) = Composite_HRR(tim(i, 1))
        End If

        'determine mass flow in the plume based on theoretical HRR
        mplume = Mass_Plume_2012(tim(i, 1), layerheight(fireroom, i), HeatRelease(fireroom, stepcount, 1), uppertemp(fireroom, i), lowertemp(fireroom, i))

        'determine oxygen limited heat release rate
        HeatRelease(fireroom, i, 2) = O2_limit_cfast(fireroom, Mass_Upper, HeatRelease(fireroom, i, 1), mplume, uppertemp(fireroom, i), O2MassFraction(fireroom, i, 1), O2MassFraction(fireroom, i, 2), mw_upper, mw_lower, Qplume)
        'HeatRelease(fireroom, i, 2) = CDbl(VB6.Format(HeatRelease(fireroom, i, 2), "0.000"))
        If Qplume > 0 Then
            mplume = Mass_Plume_2012(tim(i, 1), layerheight(fireroom, i), Qplume, uppertemp(fireroom, i), lowertemp(fireroom, i))
        Else
            mplume = 0
        End If

        count = 1
        heatreleased = HeatRelease(fireroom, i, 1) 'max theoretical hrr from the fuel

        If heatreleased > 0 Then
            Do While Abs(heatreleased - HeatRelease(fireroom, i, 2)) / heatreleased > 0.001

                'fire is ventilation-limited
                heatreleased = HeatRelease(fireroom, i, 2)
                'heatreleased = Qplume
                If heatreleased = 0 Then Exit Do

                'Mass flow in the plume
                'massplumeflow(i, fireroom) = Mass_Plume(ByVal tim(i, 1), ByVal layerheight(fireroom, i), heatreleased, ByVal uppertemp(fireroom, i), ByVal lowertemp(fireroom, i))
                mplume = Mass_Plume_2012(tim(i, 1), layerheight(fireroom, i), Qplume, uppertemp(fireroom, i), lowertemp(fireroom, i))

                'recalculate oxygen limit using new plume flow
                'HeatRelease(i, 2) = Oxygen_Limit(heatreleased, massplumeflow(i, 1), uppertemp(fireroom, i), O2MassFraction(fireroom, i, 1), O2MassFraction(fireroom, i, 2))
                HeatRelease(fireroom, i, 2) = O2_limit_cfast(fireroom, Mass_Upper, heatreleased, mplume, uppertemp(fireroom, i), O2MassFraction(fireroom, i, 1), O2MassFraction(fireroom, i, 2), mw_upper, mw_lower, Qplume)
                'HeatRelease(fireroom, i, 2) = CDbl(VB6.Format(HeatRelease(fireroom, i, 2), "0.000"))

                count = count + 1
                'Debug.Print "count", count
                If count > 50 Then
                    Exit Do
                End If
            Loop
        End If

        massplumeflow(i, fireroom) = mplume

        If VentilationLimitFlag = False Then
            If tim(i - 1, 1) = 0 Then
                alphaTfitted(4, itcounter - 1) = SimTime + 1 'initialise time ventilation limit reached to long value
            End If

            If HeatRelease(fireroom, i, 1) - HeatRelease(fireroom, i, 2) > 0.1 Then 'compare free burning HRR with oxygen constrained value

                Dim vltemp As Double = Min(HeatRelease(fireroom, i - 1, 2), HeatRelease(fireroom, i, 2)) 'use the lesser of the two (over two successive timesteps)
                Dim vl As Single = vltemp

                If alphaTfitted(3, itcounter - 1) > vl Then vl = alphaTfitted(3, itcounter - 1)

                VentilationLimitFlag = True

                Dim Message As String = CStr(tim(i - 1, 1)) & " sec. Ventilation Limit " & Format(vl, "0") & " kW."
                If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                'adjust peak hrr to be no more than 1.5 x ventilation limit as per VM2
                If VM2 = True Then PeakHRR = Min(originalpeakhrr, 1.5 * vl)

                alphaTfitted(3, itcounter - 1) = vl 'value of ventilation limit
                alphaTfitted(4, itcounter - 1) = tim(i, 1) 'time ventilation limit reached

            End If
        End If

        'calculate energy emitted/absorbed by the gas layers
        'Call Heat_Losses(tim(i), LayerHeight(i), uppertemp(i), lowertemp(i), Upperwalltemp(i), CeilingTemp(i), LowerWallTemp(i), FloorTemp(i), HeatRelease(i, 2), upperabsorb, lowerabsorb, MatC(), CO2MassFraction(i, 1), H2OMassFraction(i, 1), CO2MassFraction(i, 2), H2OMassFraction(i, 2))
        For j = 1 To NumberRooms
            If j = fireroom Then
                Call heat_losses_stiff(j, tim(i, 1), layerheight(j, i), uppertemp(j, i), lowertemp(j, i), Upperwalltemp(j, i), CeilingTemp(j, i), LowerWallTemp(j, i), FloorTemp(j, i), HeatRelease(fireroom, i, 2), upperabsorb, lowerabsorb, matc, CO2MassFraction(j, i, 1), H2OMassFraction(j, i, 1), CO2MassFraction(j, i, 2), H2OMassFraction(j, i, 2), UpperVolume(j, i))
            Else
                Call heat_losses_stiff(j, tim(i, 1), layerheight(j, i), uppertemp(j, i), lowertemp(j, i), Upperwalltemp(j, i), CeilingTemp(j, i), LowerWallTemp(j, i), FloorTemp(j, i), 0, upperabsorb, lowerabsorb, matc, CO2MassFraction(j, i, 1), H2OMassFraction(j, i, 1), CO2MassFraction(j, i, 2), H2OMassFraction(j, i, 2), UpperVolume(j, i))
            End If

            'radiation flux on floor
            Target(j, i) = matc(4, 1) 'kw/m2
            absorb(j, i, 1) = upperabsorb 'kW
            absorb(j, i, 2) = lowerabsorb 'kW

            If calcFRR = True Then
                'calc energy dose for fire severity calc
                'energy(j, i) = energy(j, i - 1) + Timestep * StefanBoltzmann / 1000 * uppertemp(j, i) ^ 4 'Ws/m2
                energy(j, i) = energy(j, i - 1) + Timestep * StefanBoltzmann / 1000 * QCeilingAST(j, 2, i) ^ 4 'Ws/m2
                'kJ/m2

                'NHL harmathy furnace
                If NHL(0, j, i) > 0 Then
                    NHL(1, j, i) = NHL(1, j, i - 1) + (NHL(0, j, i) + NHL(0, j, i - 1)) / 2 * Timestep
                End If
            End If

        Next j

        chemistry() 'moved to do_stuff and run only once per timestep

    End Sub
    Public Sub Do_Stuff_fuelresponse(ByRef i As Integer)
        '==============================================================================
        '   misc calculations
        '   called once per timestep
        '==============================================================================

        Dim stemp, ftemp, weighted_hc As Double
        Dim count As Integer
        Dim heatreleased As Double
        Dim lowerabsorb, upperabsorb, mplume As Double
        Dim j, id As Integer
        Dim sum, Qplume, burningrate As Double
        Dim massloss() As Double
        Dim mrate_ceiling, mrate_wall, mrate_floor As Double
        Dim QWall, QCeiling, QFloor As Double '05012018
        Dim mrate() As Double
        Dim matc(4, 1) As Double
        ReDim massloss(NumberObjects)
        ReDim mrate(NumberObjects)
        Dim Mass_Upper, mw_upper, mw_lower, fuelmassloss As Double
        Dim o2lower As Double = O2MassFraction(fireroom, i, 2)
        Dim ltemp As Double = lowertemp(fireroom, i)

        'effective molecular weight of the layers for fireroom
        mw_upper = MolecularWeightCO * COMassFraction(fireroom, stepcount, 1) + MolecularWeightCO2 * CO2MassFraction(fireroom, stepcount, 1) + MolecularWeightH2O * H2OMassFraction(fireroom, stepcount, 1) + MolecularWeightHCN * HCNMassFraction(fireroom, stepcount, 1) + MolecularWeightO2 * O2MassFraction(fireroom, stepcount, 1) + MolecularWeightN2 * (1 - O2MassFraction(fireroom, stepcount, 1) - COMassFraction(fireroom, stepcount, 1) - CO2MassFraction(fireroom, stepcount, 1) - H2OMassFraction(fireroom, stepcount, 1) - HCNMassFraction(fireroom, stepcount, 1))
        mw_lower = MolecularWeightCO * COMassFraction(fireroom, stepcount, 2) + MolecularWeightCO2 * CO2MassFraction(fireroom, stepcount, 2) + MolecularWeightH2O * H2OMassFraction(fireroom, stepcount, 2) + MolecularWeightHCN * HCNMassFraction(fireroom, stepcount, 2) + MolecularWeightO2 * O2MassFraction(fireroom, stepcount, 2) + MolecularWeightN2 * (1 - O2MassFraction(fireroom, stepcount, 2) - COMassFraction(fireroom, stepcount, 2) - CO2MassFraction(fireroom, stepcount, 2) - H2OMassFraction(fireroom, stepcount, 2) - HCNMassFraction(fireroom, stepcount, 2))
        'mplume = 0
        mplume = massplumeflow(i - 1, fireroom)
        count = 0

        'mass of upper layer in fireroom
        Mass_Upper = UpperVolume(fireroom, i) * (Atm_Pressure + RoomPressure(fireroom, i) / 1000) / (Gas_Constant / mw_upper) / uppertemp(fireroom, i)

        Call mass_rate_withfuelresponse(tim(i, 1), mrate, mrate_wall, mrate_ceiling, mrate_floor, o2lower, ltemp, mplume, Target(fireroom, i - 1), burningrate)
        'Call mass_rate_withfuelresponse(tim(i, 1), mrate, mrate_wall, mrate_ceiling, mrate_floor, o2lower, ltemp, mplume, -QFloor(fireroom, i - 1), burningrate)

        If Flashover = True And g_post = True Then 'using wood crib post flashover model
            sum = 0
            fuelmassloss = mrate(1)
            sum = mrate(1) * NewHoC_fuel 'kg/s.kJ/g
            ftemp = fuelmassloss 'kg/sec
            stemp = sum 'kg/s.kJ/g
        Else
            fuelmassloss = 0
            sum = 0
            ftemp = 0 'kg/sec
            stemp = 0

            For id = 1 To NumberObjects
                fuelmassloss = fuelmassloss + mrate(id)
                sum = sum + mrate(id) * NewEnergyYield(id) 'kg/s.kJ/g
                ftemp = ftemp + fuelmassloss 'kg/sec
                stemp = stemp + sum 'kg/s.kJ/g
                If id = 1 Then
                    fuelmassloss = fuelmassloss + mrate_wall + mrate_ceiling + mrate_floor 'kg/sec
                    sum = sum + mrate_wall * WallEffectiveHeatofCombustion(fireroom) + mrate_ceiling * CeilingEffectiveHeatofCombustion(fireroom) + mrate_floor * FloorEffectiveHeatofCombustion(fireroom)
                End If
            Next id
        End If

        If fuelmassloss > 0 Then weighted_hc = CDec(sum / fuelmassloss) 'MJ/kg or kJ/g
        If ftemp > 0 Then stemp = CDec(stemp / ftemp) 'kJ/g

        'well ventilated, theoretical
        'HeatRelease(fireroom, i, 1) = 1000 * EnergyYield(1) * fuelmassloss 'kW 'CW 9/4/2019
        HeatRelease(fireroom, i, 1) = 1000 * weighted_hc * fuelmassloss 'kW 'CW 9/4/2019

        HeatRelease(fireroom, i, 2) = 1000 * EnergyYield(1) * burningrate 'kW

        'determine mass flow in the plume
        mplume = Mass_Plume_2012(tim(i, 1), layerheight(fireroom, i), HeatRelease(fireroom, i, 2), uppertemp(fireroom, i), lowertemp(fireroom, i))


        Call mass_rate_withfuelresponse(tim(i, 1), mrate, mrate_wall, mrate_ceiling, mrate_floor, O2MassFraction(fireroom, i, 2), lowertemp(fireroom, i), mplume, Target(fireroom, i - 1), burningrate)
        'Call mass_rate_withfuelresponse(tim(i, 1), mrate, mrate_wall, mrate_ceiling, mrate_floor, O2MassFraction(fireroom, i, 2), lowertemp(fireroom, i), mplume, -QFloor(fireroom, i - 1), burningrate)

        HeatRelease(fireroom, i, 2) = 1000 * EnergyYield(1) * burningrate 'kW
        'End If

        count = 1

        If useCLTmodel = True And KineticModel = True Then

            HeatRelease(fireroom, i, 1) = 1000 * EnergyYield(1) * mrate(1) + 1000 * WallEffectiveHeatofCombustion(fireroom) * mrate_wall + 1000 * CeilingEffectiveHeatofCombustion(fireroom) * mrate_ceiling 'kW

            mplume = Mass_Plume_2012(tim(i, 1), layerheight(fireroom, i), HeatRelease(fireroom, i, 1), uppertemp(fireroom, i), lowertemp(fireroom, i))

            Call mass_rate_withfuelresponse(tim(i, 1), mrate, mrate_wall, mrate_ceiling, mrate_floor, O2MassFraction(fireroom, i, 2), lowertemp(fireroom, i), mplume, Target(fireroom, i - 1), burningrate)

            HeatRelease(fireroom, i, 2) = 1000 * EnergyYield(1) * burningrate 'kW

            'HeatRelease(fireroom, i, 2) = O2_limit_cfast(fireroom, Mass_Upper, HeatRelease(fireroom, i, 1), mplume, uppertemp(fireroom, i), O2MassFraction(fireroom, i, 1), O2MassFraction(fireroom, i, 2), mw_upper, mw_lower, Qplume)

            mplume = Mass_Plume_2012(tim(i, 1), layerheight(fireroom, i), HeatRelease(fireroom, i, 2), uppertemp(fireroom, i), lowertemp(fireroom, i))

            heatreleased = HeatRelease(fireroom, i, 1)

            Do While Abs(heatreleased - HeatRelease(fireroom, i, 2)) / heatreleased > 0.001
                'If tim(i, 1) = 216 Then Stop

                'fire is ventilation-limited
                heatreleased = HeatRelease(fireroom, i, 2)

                If heatreleased = 0 Then Exit Do

                'Mass flow in the plume
                mplume = Mass_Plume_2012(tim(i, 1), layerheight(fireroom, i), heatreleased, uppertemp(fireroom, i), lowertemp(fireroom, i))

                Call mass_rate_withfuelresponse(tim(i, 1), mrate, mrate_wall, mrate_ceiling, mrate_floor, O2MassFraction(fireroom, i, 2), lowertemp(fireroom, i), mplume, Target(fireroom, i - 1), burningrate)

                HeatRelease(fireroom, i, 2) = 1000 * EnergyYield(1) * burningrate 'kW

                'recalculate oxygen limit using new plume flow
                'HeatRelease(fireroom, i, 2) = O2_limit_cfast(fireroom, Mass_Upper, HeatRelease(fireroom, i, 1), mplume, uppertemp(fireroom, i), O2MassFraction(fireroom, i, 1), O2MassFraction(fireroom, i, 2), mw_upper, mw_lower, Qplume)

                count = count + 1
                If count > 100 Then
                    Exit Do
                End If
            Loop

        Else

                Do While Abs(heatreleased - HeatRelease(fireroom, i, 2)) / heatreleased > 0.001

                'fire is ventilation-limited
                heatreleased = HeatRelease(fireroom, i, 2)

                If heatreleased = 0 Then Exit Do

                'Mass flow in the plume
                mplume = Mass_Plume_2012(tim(i, 1), layerheight(fireroom, i), heatreleased, uppertemp(fireroom, i), lowertemp(fireroom, i))

                Call mass_rate_withfuelresponse(tim(i, 1), mrate, mrate_wall, mrate_ceiling, mrate_floor, O2MassFraction(fireroom, i, 2), lowertemp(fireroom, i), mplume, Target(fireroom, i - 1), burningrate)
                'Call mass_rate_withfuelresponse(tim(i, 1), mrate, mrate_wall, mrate_ceiling, mrate_floor, O2MassFraction(fireroom, i, 2), lowertemp(fireroom, i), mplume, -QFloor(fireroom, i - 1), burningrate)
                HeatRelease(fireroom, i, 2) = 1000 * EnergyYield(1) * burningrate 'kW

                count = count + 1
                If count > 50 Then
                    Exit Do
                End If
            Loop
        End If


        massplumeflow(i, fireroom) = mplume

        If VentilationLimitFlag = False Then
            If tim(i - 1, 1) = 0 Then
                alphaTfitted(4, itcounter - 1) = SimTime + 1 'initialise time ventilation limit reached to long value
            End If

            If HeatRelease(fireroom, i, 1) - HeatRelease(fireroom, i, 2) > 0.1 Then 'compare free burning HRR with oxygen constrained value

                Dim vltemp As Double = Min(HeatRelease(fireroom, i - 1, 2), HeatRelease(fireroom, i, 2)) 'use the lesser of the two (over two successive timesteps)
                Dim vl As Single = vltemp

                If alphaTfitted(3, itcounter - 1) > vl Then vl = alphaTfitted(3, itcounter - 1)

                VentilationLimitFlag = True

                Dim Message As String = CStr(tim(i - 1, 1)) & " sec. Ventilation Limit " & Format(vl, "0") & " kW."
                If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                'adjust peak hrr to be no more than 1.5 x ventilation limit as per VM2
                If VM2 = True Then PeakHRR = Min(originalpeakhrr, 1.5 * vl)

                alphaTfitted(3, itcounter - 1) = vl 'value of ventilation limit
                alphaTfitted(4, itcounter - 1) = tim(i, 1) 'time ventilation limit reached

            End If
        End If

        'calculate energy emitted/absorbed by the gas layers
        For j = 1 To NumberRooms
            If j = fireroom Then
                Call heat_losses_stiff(j, tim(i, 1), layerheight(j, i), uppertemp(j, i), lowertemp(j, i), Upperwalltemp(j, i), CeilingTemp(j, i), LowerWallTemp(j, i), FloorTemp(j, i), HeatRelease(fireroom, i, 2), upperabsorb, lowerabsorb, matc, CO2MassFraction(j, i, 1), H2OMassFraction(j, i, 1), CO2MassFraction(j, i, 2), H2OMassFraction(j, i, 2), UpperVolume(j, i))
            Else
                Call heat_losses_stiff(j, tim(i, 1), layerheight(j, i), uppertemp(j, i), lowertemp(j, i), Upperwalltemp(j, i), CeilingTemp(j, i), LowerWallTemp(j, i), FloorTemp(j, i), 0, upperabsorb, lowerabsorb, matc, CO2MassFraction(j, i, 1), H2OMassFraction(j, i, 1), CO2MassFraction(j, i, 2), H2OMassFraction(j, i, 2), UpperVolume(j, i))
            End If

            'radiation flux on floor
            Target(j, i) = matc(4, 1) 'kw/m2
            absorb(j, i, 1) = upperabsorb 'kW
            absorb(j, i, 2) = lowerabsorb 'kW

            If calcFRR = True Then
                'calc energy dose for fire severity calc
                energy(j, i) = energy(j, i - 1) + Timestep * StefanBoltzmann / 1000 * QCeilingAST(j, 2, i) ^ 4 'Ws/m2
                'kJ/m2

                'NHL harmathy furnace
                If NHL(0, j, i) > 0 Then
                    NHL(1, j, i) = NHL(1, j, i - 1) + (NHL(0, j, i) + NHL(0, j, i - 1)) / 2 * Timestep
                End If
            End If

        Next j

        chemistry() 'moved to do_stuff and run only once per timestep

    End Sub

    Private Sub diff_eqns(ByVal X As Double, ByRef Y(,) As Double, ByRef DYDX(,) As Double)
        '============================================================================
        ' Differential equations subroutine - solving the coupled equations
        ' Input
        '  X   = current X value
        '  Y() = current integration vector
        '        (N rows by 1 column)
        ' Output
        '  DYDX() = system of differential equations evaluated
        '           at X, Y() (N rows by 1 column)
        ' C Wade 7/8/98
        ' 1 feb 2008 Amanda
        '============================================================================
        Dim r2u, deltaEU, sdot, deltaEL, beta, r2l As Double
        Dim response As Short
        Dim k, i, h, j As Integer
        Dim dMudt, dMldt As Double
        Dim doormix() As Double
        Dim Enthalpy_L, denl, denu, Enthalpy_U, sum As Double
        Dim layerz() As Double
        Dim massplume() As Double
        Dim rate_Renamed As Double
        Dim mrate() As Double
        Dim Enthalpy(1, 1) As Double
        Dim products(1, 1, 1) As Double
        Dim fuelmassloss() As Double
        Dim ventflow(1, 1) As Double
        Dim Flow_to_Lower, Flow_to_Upper, F As Double
        Dim FlowOut() As Double
        Dim Flowtoroom() As Double
        Dim id As Integer
        Dim weighted_hc, absorb_l, absorb_u, hrr, mrate_floor As Double
        Dim hrrlimit() As Double
        Dim massplumeX() As Double
        Dim ventfire_sum() As Double
        Dim fanflowout As Double
        Dim mrate_wall As Double
        Dim mrate_ceiling As Double
        Dim vent As Integer
        Dim matc(,) As Double
        Dim Qplume, hrrnextroom, M_wall, burningrate As Double
        Dim mw_upper() As Double
        Dim mw_lower() As Double
        Dim species_to_upper(9) As Double
        Dim species_to_lower(9) As Double
        Dim genrate As Double
        Dim GER As Double
        Dim CJTemp, CJMax As Double
        Dim vententrain(1, 1, 1) As Double
        Dim stemp, ftemp As Double
        Dim Mass_Upper() As Double
        Dim Mass_Lower() As Double
        Dim M_wall_u, frac_to_lower, Mwu, Mwl, E_wall, M_wall_l As Double
        Dim cp_u() As Double
        Dim cp_l() As Double
        Dim gamma_l, gamma_u, species_mass As Double
        Dim wall_spec(7) As Double

        ReDim cp_u(NumberRooms)
        ReDim cp_l(NumberRooms)
        ReDim Flowtoroom(NumberRooms)
        ReDim FlowOut(NumberRooms)
        ReDim Mass_Lower(NumberRooms)
        ReDim mrate(NumberObjects)
        ReDim layerz(NumberRooms)
        ReDim doormix(NumberRooms)
        ReDim ventfire_sum(NumberRooms + 1)
        ReDim matc(4, 1)
        ReDim hrrlimit(NumberRooms)
        ReDim mw_upper(NumberRooms)
        ReDim mw_lower(NumberRooms)
        ReDim fuelmassloss(NumberRooms)
        ReDim massplumeX(NumberRooms)
        ReDim massplume(NumberRooms)
        ReDim Mass_Upper(NumberRooms)
        If MaxNumberVents > 0 Then
            ReDim vententrain(NumberRooms, NumberRooms, MaxNumberVents)
        End If
        Static xflag As Short
        Static xlast As Double

        If X = xlast Then
            xflag = 1
        Else
            xflag = 0
        End If

        On Error GoTo errorhandler
        ' If X = 180 Then Stop
        i = stepcount

        For j = 1 To NumberRooms
            If Y(j, 2) > 1300 + 273 Then
                Y(j, 2) = 1300 + 273
            End If
            If Y(j, 2) < 50 Or Y(j, 2) > 3000 Then
                Y(j, 2) = uppertemp(j, stepcount)
            End If
            If Y(j, 3) < 50 Or Y(j, 3) > 3000 Then
                Y(j, 3) = lowertemp(j, stepcount)
            End If

            'check upper volume within valid range
            If TwoZones(j) = False Then
                'Y(j, 1) = RoomVolume(j)
                Y(j, 1) = RoomVolume(j) - layerheight(j, i) * RoomLength(j) * RoomWidth(j) 'qqq
                layerz(j) = layerheight(j, i)
            End If

            'put some limits on variable values
            If Y(j, 1) < 0.0001 * RoomVolume(j) Then
                Y(j, 1) = 0.0001 * RoomVolume(j)
            End If

            'If Y(j, 1) > RoomVolume(j) - Smallvalue Then
            '    Y(j, 1) = RoomVolume(j) - Smallvalue
            'End If
            If Y(j, 1) > RoomVolume(j) - 0.005 * RoomFloorArea(j) Then 'consistent with min layer height of 0.005m
                Y(j, 1) = RoomVolume(j) - 0.005 * RoomFloorArea(j)
            End If


            If Y(j, 12) < Smallvalue Then Y(j, 12) = Smallvalue
            If Y(j, 13) < Smallvalue Then Y(j, 13) = Smallvalue
            If Y(j, 14) < Smallvalue Then Y(j, 14) = Smallvalue
            If Y(j, 8) < Smallvalue Then Y(j, 8) = Smallvalue
            If Y(j, 9) < Smallvalue Then Y(j, 9) = Smallvalue
            If Y(j, 10) < CO2Ambient Then Y(j, 10) = CO2Ambient
            If Y(j, 11) < CO2Ambient Then Y(j, 11) = CO2Ambient
            If Y(j, 16) < H2OMassFraction(j, 1, 1) Then Y(j, 16) = H2OMassFraction(j, 1, 1)
            If Y(j, 17) < H2OMassFraction(j, 1, 2) Then Y(j, 17) = H2OMassFraction(j, 1, 2)
            If Y(j, 5) > O2Ambient Then Y(j, 5) = O2Ambient 'upper oxygen not higher than ambient
            If Y(j, 15) > O2Ambient Then
                Y(j, 15) = O2Ambient 'lower oxygen not higher than ambient
            End If

            If Y(j, 5) < 0.001 Then Y(j, 5) = 0.001 'upper oxygen not less than 0.1%
            If Y(j, 15) < 0.001 Then Y(j, 15) = 0.001 'lower oxygen not less than 0.1%
            If Y(j, 6) < TUHC(j, 1, 1) Then Y(j, 6) = TUHC(j, 1, 1) 'upper tuhc not less than 0.01%
            If Y(j, 7) < TUHC(j, 1, 2) Then Y(j, 7) = TUHC(j, 1, 2) 'lower tuhc not less than 0.01%

            If TwoZones(j) = True Then
                Call Layer_Height(Y(j, 1), layerz(j), j)
            End If

            'effective molecular weight of the layers for fireroom
            mw_upper(j) = MolecularWeightCO * Y(j, 8) + MolecularWeightCO2 * Y(j, 10) + MolecularWeightH2O * Y(j, 16) + MolecularWeightO2 * Y(j, 5) + MolecularWeightN2 * (1 - Y(j, 5) - Y(j, 16) - Y(j, 8) - Y(j, 10))
            If mw_upper(j) < 0 Then
                Exit Sub
            End If
            mw_lower(j) = MolecularWeightCO * Y(j, 9) + MolecularWeightCO2 * Y(j, 11) + MolecularWeightH2O * Y(j, 17) + MolecularWeightO2 * Y(j, 15) + MolecularWeightN2 * (1 - Y(j, 15) - Y(j, 17) - Y(j, 9) - Y(j, 11))
            If mw_lower(j) < 0 Then
                Exit Sub
            End If
            'specific heats
            cp_u(j) = specific_heat_upper(j, Y, Y(j, 2)) 'specific_heat("CO", Y(j, 2)) * Y(j, 8) + specific_heat("CO2", Y(j, 2)) * Y(j, 10) + specific_heat("H2O", Y(j, 2)) * Y(j, 16) + specific_heat("O2", Y(j, 2)) * Y(j, 5) + specific_heat("N2", Y(j, 2)) * (1 - Y(j, 5) - Y(j, 16) - Y(j, 8) - Y(j, 10))
            cp_l(j) = specific_heat_lower(j, Y, Y(j, 3)) 'specific_heat("CO", Y(j, 3)) * Y(j, 9) + specific_heat("CO2", Y(j, 3)) * Y(j, 11) + specific_heat("H2O", Y(j, 3)) * Y(j, 17) + specific_heat("O2", Y(j, 3)) * Y(j, 15) + specific_heat("N2", Y(j, 3)) * (1 - Y(j, 15) - Y(j, 17) - Y(j, 9) - Y(j, 11))
            'changed y(j,2) to Y(j,3) in call in line above 'CW

            gamma_u = gamma
            gamma_l = gamma

            If j = fireroom Then


                If FuelResponseEffects = True Then
                    'massplumeX(j) = Mass_Plume_2012(X, layerz(j), HeatRelease(j, stepcount, 1), Y(j, 2), Y(j, 3)) '18/12/18 - switched to line below
                    massplumeX(j) = Mass_Plume_2012(X, layerz(j), HeatRelease(j, stepcount, 2), Y(j, 2), Y(j, 3))
                    Call mass_rate_withfuelresponse(X, mrate, mrate_wall, mrate_ceiling, mrate_floor, Y(j, 15), Y(j, 3), massplumeX(j), Target(fireroom, i), burningrate)
                    'Call mass_rate_withfuelresponse(X, mrate, mrate_wall, mrate_ceiling, mrate_floor, Y(j, 15), Y(j, 3), massplumeX(j), -QFloor(fireroom, i), burningrate)

                    hrrlimit(fireroom) = 1000 * EnergyYield(1) * burningrate

                    Call hrr_estimate_fuelresponse(fireroom, Mass_Upper(fireroom), massplumeX(fireroom), hrrlimit(fireroom), X, layerz(fireroom), Y(fireroom, 2), Y(fireroom, 3), Y(fireroom, 5), Y(fireroom, 15), mw_upper(fireroom), mw_lower(fireroom), Y(fireroom, 6), weighted_hc, Target(fireroom, i))
                    'Call hrr_estimate_fuelresponse(fireroom, Mass_Upper(fireroom), massplumeX(fireroom), hrrlimit(fireroom), X, layerz(fireroom), Y(fireroom, 2), Y(fireroom, 3), Y(fireroom, 5), Y(fireroom, 15), mw_upper(fireroom), mw_lower(fireroom), Y(fireroom, 6), weighted_hc, -QFloor(fireroom, i))

                Else
                    Call mass_rate(X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
                End If

                fuelmassloss(j) = 0
                sum = 0
                ftemp = 0 'kg/sec
                stemp = 0
                If Flashover = True And g_post = True Then
                    fuelmassloss(j) = mrate(1)
                    sum = mrate(1) * NewHoC_fuel 'kg/s.kJ/g
                Else
                    For id = 1 To NumberObjects
                        fuelmassloss(j) = fuelmassloss(j) + mrate(id)
                        sum = sum + mrate(id) * NewEnergyYield(id) 'kg/s.kJ/g
                        ftemp = ftemp + fuelmassloss(j) 'kg/sec
                        stemp = stemp + sum 'kg/s.kJ/g
                        If id = 1 Then
                            fuelmassloss(j) = fuelmassloss(j) + mrate_wall + mrate_ceiling + mrate_floor 'kg/sec
                            sum = sum + mrate_wall * WallEffectiveHeatofCombustion(j) + mrate_ceiling * CeilingEffectiveHeatofCombustion(j) + mrate_floor * FloorEffectiveHeatofCombustion(j)
                        End If
                    Next id
                End If
                If fuelmassloss(j) > 0 Then weighted_hc = sum / fuelmassloss(j) 'MJ/kg or kJ/g
                If ftemp > 0 Then stemp = stemp / ftemp 'kJ/g
            End If

            'mass of the upper layer (kg)
            Mass_Lower(j) = (RoomVolume(j) - Y(j, 1)) * (Atm_Pressure + Y(j, 4) / 1000) / (Gas_Constant / mw_lower(j)) / Y(j, 3)
            Mass_Upper(j) = Y(j, 1) * (Atm_Pressure + Y(j, 4) / 1000) / (Gas_Constant / mw_upper(j)) / Y(j, 2)
        Next j

        If FuelResponseEffects = True Then
            'hrrlimit(fireroom) = HeatRelease(fireroom, stepcount, 1)
            'Call hrr_estimate_fuelresponse(fireroom, Mass_Upper(fireroom), massplumeX(fireroom), hrrlimit(fireroom), X, layerz(fireroom), Y(fireroom, 2), Y(fireroom, 3), Y(fireroom, 5), Y(fireroom, 15), mw_upper(fireroom), mw_lower(fireroom), Y(fireroom, 6), weighted_hc, 0)
        Else
            Call hrr_estimate(fireroom, Mass_Upper(fireroom), massplumeX(fireroom), hrrlimit(fireroom), X, layerz(fireroom), Y(fireroom, 2), Y(fireroom, 3), Y(fireroom, 5), Y(fireroom, 15), mw_upper(fireroom), mw_lower(fireroom), Y(fireroom, 6), weighted_hc)
        End If
        'Call hrr_estimate(fireroom, Mass_Upper(fireroom), massplumeX(fireroom), hrrlimit(fireroom), X, layerz(fireroom), Y(fireroom, 2), Y(fireroom, 3), Y(fireroom, 5), Y(fireroom, 15), mw_upper(fireroom), mw_lower(fireroom), Y(fireroom, 6), weighted_hc)

        'calculate wall vent flows
        'If X = 183 Then Stop
        'If Y(2, 2) = 288 Then Stop

        Call ventflows(Mass_Upper, vententrain, X, Y, ventflow, layerz, Enthalpy, products, ventfire_sum, mw_upper, mw_lower, weighted_hc, FlowOut, Flowtoroom, cp_u, cp_l)
        If flagstop = 1 Then Exit Sub

        For j = 1 To NumberRooms 'for each room
            denu = Mass_Upper(j) / Y(j, 1)
            denl = Mass_Lower(j) / (RoomVolume(j) - Y(j, 1))
            If j <> fireroom Then
                massplumeX(j) = 0

                'if the floor has ignited then we need to treat as an axisymetrical plume
                If FloorIgniteFlag(j) = 1 Then 'assume all hrr is from the floor
                    If HeatRelease(j, stepcount, 1) > 0 Then massplumeX(j) = Mass_Plume_nextroom(X, layerz(j), HeatRelease(j, stepcount, 1), Y(j, 2), Y(j, 3))
                End If

                If fireroom < j Then
                    For vent = 1 To NumberVents(fireroom, j)
                        massplumeX(j) = massplumeX(j) + vententrain(fireroom, j, vent)
                    Next vent
                Else
                    For vent = 1 To NumberVents(j, fireroom)
                        massplumeX(j) = massplumeX(j) + vententrain(fireroom, j, vent)
                    Next vent
                End If
                If HeatRelease(j, stepcount, 1) + ventfire_sum(j) > 0 Then
                    hrrnextroom = O2_limit_cfast(j, Mass_Upper(j), HeatRelease(j, stepcount, 1) + ventfire_sum(j), massplumeX(j), Y(j, 2), Y(j, 5), Y(j, 15), mw_upper(j), mw_lower(j), Qplume)
                End If
            End If

            Flow_to_Upper = 0
            Flow_to_Lower = 0
            Enthalpy_U = 0
            Enthalpy_L = 0
            For h = 3 To 9
                species_to_upper(h) = 0
                species_to_lower(h) = 0
            Next h

            'flow to upper layer of room j
            Flow_to_Upper = ventflow(j, 2)
            'flow to lower layer of room j
            Flow_to_Lower = ventflow(j, 1)

            If j = fireroom Then 'Mass flow in the plume
                massplume(j) = massplumeX(j)
                If massplume(j) > 0.0001 Then

                    GER = weighted_hc * fuelmassloss(j) / (13.1 * massplume(j) * Y(j, 15)) 'plume GER

                Else
                    GER = 0
                End If
            Else
                massplume(j) = 0
                If IgniteNextRoom = True And X > CeilingIgniteTime(j) Then
                    fuelmassloss(j) = (hrrnextroom - ventfire_sum(j)) / (CeilingEffectiveHeatofCombustion(j) * 1000)
                    If fuelmassloss(j) < 0 Then fuelmassloss(j) = 0
                Else
                    fuelmassloss(j) = 0
                End If
                If massplumeX(j) > 0 Then
                    GER = weighted_hc * fuelmassloss(j) / (deltaHair * massplumeX(j))
                Else
                    GER = 0
                End If
            End If

            If GER > 0.5 And j = fireroom Then
                If frmOptions1.chkModGER.CheckState = System.Windows.Forms.CheckState.Checked Then
                    For k = 1 To NumberObjects
                        NewEnergyYield(k) = EnergyYield(k) * (1 - (0.97 / (Exp((GER / 2.15) ^ (-1.2))))) 'SFPE h/b 3rd ed page 3-118 eqn 40
                        NewEnergyYield(k) = EnergyYield(k)
                        NewRadiantLossFraction(k) = 1 - ((1 - ObjectRLF(k)) * (1 - 1 / Exp((GER / 1.38) ^ (-2.8))) / (1 - (0.97) / (Exp((GER / 2.15) ^ (-1.2)))))

                    Next k
                    'If Flashover = False Then NewHoC_fuel = HoC_fuel * (1 - (0.97 / (Exp((GER / 2.15) ^ (-1.2)))))
                    NewHoC_fuel = HoC_fuel * (1 - (0.97 / (Exp((GER / 2.15) ^ (-1.2)))))
                Else
                    For k = 1 To NumberObjects
                        NewEnergyYield(k) = EnergyYield(k)
                        NewRadiantLossFraction(k) = ObjectRLF(k)
                    Next k
                    NewHoC_fuel = HoC_fuel

                End If
            ElseIf j = fireroom Then
                For k = 1 To NumberObjects
                    NewEnergyYield(k) = EnergyYield(k)
                    NewRadiantLossFraction(k) = ObjectRLF(k)
                Next k
                NewHoC_fuel = HoC_fuel

            End If

            If i < NumberTimeSteps + 1 Then
                If j = fireroom Then
                    Call heat_losses_stiff2(j, X, layerz(j), Y(j, 2), Y(j, 3), Upperwalltemp(j, i + 1), CeilingTemp(j, i + 1), LowerWallTemp(j, i + 1), FloorTemp(j, i + 1), hrrlimit(fireroom), absorb_u, absorb_l, matc, Y(j, 10), Y(j, 16), Y(j, 11), Y(j, 17), Y(j, 1))
                Else
                    Call heat_losses_stiff2(j, X, layerz(j), Y(j, 2), Y(j, 3), Upperwalltemp(j, i + 1), CeilingTemp(j, i + 1), LowerWallTemp(j, i + 1), FloorTemp(j, i + 1), hrrnextroom, absorb_u, absorb_l, matc, Y(j, 10), Y(j, 16), Y(j, 11), Y(j, 17), Y(j, 1))
                End If
            End If
            Enthalpy_U = Enthalpy(j, 2) / 1000 'kw
            Enthalpy_L = Enthalpy(j, 1) / 1000 'kw

            For h = 3 To 9
                species_to_upper(h) = products(h, j, 2) 'kg species/sec in ventflows
                species_to_lower(h) = products(h, j, 1) 'kg species/sec in ventflows
            Next h

            If j = fireroom Then
                Enthalpy_U = Enthalpy_U + massplume(j) * cp_l(j) * Y(j, 3) + absorb_u + (1 - NewRadiantLossFraction(1)) * hrrlimit(fireroom) + fuelmassloss(j) * SpecificHeat_fuel * Flametemp
                Enthalpy_L = Enthalpy_L - massplume(j) * cp_l(j) * Y(j, 3) + absorb_l
            Else
                If IgniteNextRoom = True And X > CeilingIgniteTime(j) Then
                    Enthalpy_U = Enthalpy_U + absorb_u + (1 - NewRadiantLossFraction(1)) * hrrnextroom
                Else
                    Enthalpy_U = Enthalpy_U + absorb_u
                End If
                Enthalpy_L = Enthalpy_L + absorb_l
            End If

            frac_to_lower = 0
            M_wall = 0
            M_wall_u = 0
            M_wall_l = 0

            If j = fireroom And nowallflow = 0 Then
                'wall flows
                If Y(j, 2) > Upperwalltemp(j, i) Then
                    Mwu = wall_flow_momentum(Mass_Upper(j) / Y(j, 1), X, j, Y(j, 2), Upperwalltemp(j, i), layerz(j), RoomHeight(j) - layerz(j))
                    'flow from upper layer to lower layer
                    M_wall_u = wall_flow(Mass_Upper(j) / Y(j, 1), X, j, Y(j, 2), Upperwalltemp(j, i), layerz(j), RoomHeight(j) - layerz(j)) 'kg/s
                Else
                    Mwu = 0
                    M_wall_u = 0
                End If

                If Y(j, 3) < LowerWallTemp(j, i) Then
                    Mwl = wall_flow_momentum(Mass_Lower(j) / (RoomVolume(j) - Y(j, 1)), X, j, Y(j, 3), LowerWallTemp(j, i), layerz(j), layerz(j))
                    'flow from lower layer to upper layer
                    M_wall_l = wall_flow(Mass_Lower(j) / (RoomVolume(j) - Y(j, 1)), X, j, Y(j, 3), LowerWallTemp(j, i), layerz(j), layerz(j))
                Else
                    Mwl = 0
                    M_wall_l = 0
                End If

                'remove flow from source
                species_to_upper(3) = species_to_upper(3) - M_wall_u * Y(j, 5) 'oxygen
                species_to_upper(4) = species_to_upper(4) - M_wall_u * Y(j, 8) 'co
                species_to_upper(5) = species_to_upper(5) - M_wall_u * Y(j, 10) 'co2
                species_to_upper(6) = species_to_upper(6) - M_wall_u * Y(j, 12) 'soot
                species_to_upper(7) = species_to_upper(7) - M_wall_u * Y(j, 16) 'h20
                species_to_upper(8) = species_to_upper(8) - M_wall_u * Y(j, 6) 'tuhc
                species_to_upper(9) = species_to_upper(9) - M_wall_u * Y(j, 18) 'hcn

                species_to_lower(3) = species_to_lower(3) - M_wall_l * Y(j, 15) 'oxygen
                species_to_lower(4) = species_to_lower(4) - M_wall_l * Y(j, 9) 'co
                species_to_lower(5) = species_to_lower(5) - M_wall_l * Y(j, 11) 'co2
                species_to_lower(6) = species_to_lower(6) - M_wall_l * Y(j, 13) 'soot
                species_to_lower(7) = species_to_lower(7) - M_wall_l * Y(j, 17) 'h20
                species_to_lower(8) = species_to_lower(8) - M_wall_l * Y(j, 7) 'tuhc
                species_to_lower(9) = species_to_lower(9) - M_wall_l * Y(j, 14) 'hcn
                Enthalpy_U = Enthalpy_U - M_wall_u * Y(j, 2) * cp_u(j)
                Enthalpy_L = Enthalpy_L - M_wall_l * Y(j, 3) * cp_l(j)

                'distribute flow to destinations
                M_wall = M_wall_u + M_wall_l
                If M_wall > 0 Then
                    frac_to_lower = Mwu / (Mwu + Mwl)

                    E_wall = M_wall_u * Y(j, 2) * cp_u(j) + M_wall_l * Y(j, 3) * cp_l(j)
                    wall_spec(1) = (M_wall_u * Y(j, 5) + M_wall_l * Y(j, 15)) / M_wall 'oxy
                    wall_spec(2) = (M_wall_u * Y(j, 8) + M_wall_l * Y(j, 9)) / M_wall
                    wall_spec(3) = (M_wall_u * Y(j, 10) + M_wall_l * Y(j, 11)) / M_wall
                    wall_spec(4) = (M_wall_u * Y(j, 12) + M_wall_l * Y(j, 13)) / M_wall
                    wall_spec(5) = (M_wall_u * Y(j, 16) + M_wall_l * Y(j, 17)) / M_wall
                    wall_spec(6) = (M_wall_u * Y(j, 6) + M_wall_l * Y(j, 7)) / M_wall
                    wall_spec(7) = (M_wall_u * Y(j, 18) + M_wall_l * Y(j, 14)) / M_wall

                    species_to_upper(3) = species_to_upper(3) + M_wall * (1 - frac_to_lower) * wall_spec(1) 'oxygen
                    species_to_upper(4) = species_to_upper(4) + M_wall * (1 - frac_to_lower) * wall_spec(2) 'co
                    species_to_upper(5) = species_to_upper(5) + M_wall * (1 - frac_to_lower) * wall_spec(3) 'co2
                    species_to_upper(6) = species_to_upper(6) + M_wall * (1 - frac_to_lower) * wall_spec(4) 'soot
                    species_to_upper(7) = species_to_upper(7) + M_wall * (1 - frac_to_lower) * wall_spec(5) 'h20
                    species_to_upper(8) = species_to_upper(8) + M_wall * (1 - frac_to_lower) * wall_spec(6) 'tuhc
                    species_to_upper(9) = species_to_upper(9) + M_wall * (1 - frac_to_lower) * wall_spec(7) 'hcn

                    species_to_lower(3) = species_to_lower(3) + M_wall * frac_to_lower * wall_spec(1) 'oxygen
                    species_to_lower(4) = species_to_lower(4) + M_wall * frac_to_lower * wall_spec(2) 'co
                    species_to_lower(5) = species_to_lower(5) + M_wall * frac_to_lower * wall_spec(3) 'co2
                    species_to_lower(6) = species_to_lower(6) + M_wall * frac_to_lower * wall_spec(4) 'soot
                    species_to_lower(7) = species_to_lower(7) + M_wall * frac_to_lower * wall_spec(5) 'h20
                    species_to_lower(8) = species_to_lower(8) + M_wall * frac_to_lower * wall_spec(6) 'tuhc
                    species_to_lower(9) = species_to_lower(9) + M_wall * frac_to_lower * wall_spec(7) 'hcn
                    Enthalpy_U = Enthalpy_U + E_wall * (1 - frac_to_lower)
                    Enthalpy_L = Enthalpy_L + E_wall * frac_to_lower
                End If

            End If

            If X = tim(i, 1) Then

                If j = fireroom Then GlobalER(i) = GER

                FuelMassLossRate(i, j) = fuelmassloss(j) 'in the case of CLT model this includes contents, wall, ceiling mass otherwise not
                If j = fireroom Then WoodBurningRate(i) = mrate_ceiling + mrate_wall + mrate_floor 'kg/s use in CLT model
                WallFlowtoUpper(j, i) = -M_wall_u + (1 - frac_to_lower) * M_wall
                WallFlowtoLower(j, i) = -M_wall_l + frac_to_lower * M_wall
                FlowToLower(j, i) = Flow_to_Lower
                FlowToUpper(j, i) = Flow_to_Upper
                UFlowToOutside(j, i) = FlowOut(j)
                FlowGER(j, i) = Flowtoroom(j)
                ventfire(NumberRooms + 1, i) = ventfire_sum(NumberRooms + 1)
                ventfire(j, i) = ventfire_sum(j)
                HeatRelease(fireroom, i, 2) = CDbl(hrrlimit(fireroom))
                If j = fireroom Then FuelHoC(i) = stemp 'excludes walls or ceiling
                If j <> fireroom Then HeatRelease(j, i, 2) = hrrnextroom
                xlast = X
            End If

            'mechanical ventilation
            Dim fanflag As Boolean = False
            If NumFans > 0 And mv_mode = True Then 'system is operational
                'can run this routine for each fan in room j
                Dim upperfans As Single = 0
                Dim lowerfans As Single = 0

                'fanflag = True
                For fx = 1 To NumFans
                    If fandata(fx, 0) = j Then 'room
                        fanflag = True
                        If fandata(fx, 4) = 0 Then 'manual op
                            Call mech_vent2(upperfans, lowerfans, fx, X, j, layerz(j), Y, denl, denu, species_to_upper, species_to_lower, Enthalpy_U, Enthalpy_L, cp_u, cp_l, massplume, M_wall_u, M_wall_l, fuelmassloss, Flow_to_Upper, Flow_to_Lower, dMudt, dMldt, M_wall, frac_to_lower)
                        ElseIf fandata(fx, 4) = 1 Then 'local sd
                            If SDFlag(j) = 1 Then
                                fandata(fx, 3) = SDTime(j) 'smoke det in room j has operated
                                Call mech_vent2(upperfans, lowerfans, fx, X, j, layerz(j), Y, denl, denu, species_to_upper, species_to_lower, Enthalpy_U, Enthalpy_L, cp_u, cp_l, massplume, M_wall_u, M_wall_l, fuelmassloss, Flow_to_Upper, Flow_to_Lower, dMudt, dMldt, M_wall, frac_to_lower)
                            End If
                        Else
                            If SDFlag(fireroom) = 1 Then
                                fandata(fx, 3) = SDTime(fireroom) 'smoke det in fireroom has operated
                                Call mech_vent2(upperfans, lowerfans, fx, X, j, layerz(j), Y, denl, denu, species_to_upper, species_to_lower, Enthalpy_U, Enthalpy_L, cp_u, cp_l, massplume, M_wall_u, M_wall_l, fuelmassloss, Flow_to_Upper, Flow_to_Lower, dMudt, dMldt, M_wall, frac_to_lower)
                            End If

                        End If
                    End If
                Next
                If fanflag = True Then
                    dMudt = massplume(j) - M_wall_u + (1 - frac_to_lower) * M_wall + fuelmassloss(j) + Flow_to_Upper - upperfans
                    dMldt = Flow_to_Lower - M_wall_l + frac_to_lower * M_wall - massplume(j) - lowerfans
                End If
            End If
            If fanflag = False Then 'no mechanical ventilation
                dMudt = massplume(j) + fuelmassloss(j) + Flow_to_Upper - M_wall_u + (1 - frac_to_lower) * M_wall
                dMldt = Flow_to_Lower - massplume(j) - M_wall_l + (frac_to_lower) * M_wall
            End If

            'mechanical ventilation OLD
            'If fanon(j) = True Then

            '    Call mech_vent(X, j, layerz(j), Y, denl, denu, species_to_upper, species_to_lower, Enthalpy_U, Enthalpy_L, cp_u, cp_l, massplume, M_wall_u, M_wall_l, fuelmassloss, Flow_to_Upper, Flow_to_Lower, dMudt, dMldt, M_wall, frac_to_lower)

            'Else 'no mechanical ventilation
            '    dMudt = massplume(j) + fuelmassloss(j) + Flow_to_Upper - M_wall_u + (1 - frac_to_lower) * M_wall
            '    dMldt = Flow_to_Lower - massplume(j) - M_wall_l + (frac_to_lower) * M_wall
            'End If

            'pressure
            DYDX(j, 4) = derivs_Pressure(j, Enthalpy_U, Enthalpy_L, gamma_u)

            'volume upper layer
            If TwoZones(j) = True Then
                DYDX(j, 1) = derivs_UVol(DYDX(j, 4), Y(j, 4), Y(j, 1), Enthalpy_U, gamma_u)
            Else
                DYDX(j, 1) = 0 'qqq
                'DYDX(j, 1) = derivs_UVol(DYDX(j, 4), Y(j, 4), RoomVolume(j), Enthalpy_U + Enthalpy_L, gamma_u)

            End If
            'If j = 8 And stepcount = 90 Then Stop

            'upper layer temperature
            DYDX(j, 2) = derivs_Temp(cp_u(j), mw_upper(j), DYDX(j, 4), Y(j, 1), Y(j, 2), Enthalpy_U, dMudt, Y(j, 4))


            'mass used in derivs_species_U
            species_mass = species_mass_U(mw_upper(j), Y(j, 4), Y(j, 1), Y(j, 2))

            'O2 concentration
            rate_Renamed = 0
            If j = fireroom Then
                rate_Renamed = O2Mass_Rate(hrrlimit(fireroom))
            Else
                If ventfire_sum(j) + HeatRelease(j, stepcount, 2) > 0 Then 'unburned fuel flowing to upper layer of room j
                    rate_Renamed = O2Mass_Rate(hrrnextroom) 'assumes oxygen taken from upper layer
                End If
            End If
            'DYDX(j, 5) = derivs_species_U(ByVal mw_upper(j), ByVal Y(j, 4), ByVal dMudt, ByVal rate, ByVal species_to_upper(3), ByVal Y(j, 5), ByVal Y(j, 15), ByVal massplume(j), ByVal Y(j, 1), ByVal Y(j, 2))
            DYDX(j, 5) = derivs_species_U(species_mass, dMudt, rate_Renamed, species_to_upper(3), Y(j, 5), Y(j, 15), massplume(j))

            'H2O concentration
            rate_Renamed = 0
            If j = fireroom Then
                If Flashover = True And frmOptions1.optPostFlashover.Checked = True Then
                    rate_Renamed = species_mass_rate_postflashover(GER, j, H2OYieldPF, WallH2OYield, CeilingH2OYield, FloorH2OYield, X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
                Else
                    If GER > 1 Then
                        rate_Renamed = species_mass_rate2(GER, j, WaterVaporYield, WallH2OYield, CeilingH2OYield, FloorH2OYield, X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
                    Else
                        rate_Renamed = species_mass_rate(j, WaterVaporYield, WallH2OYield, CeilingH2OYield, FloorH2OYield, X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
                    End If
                End If
                If fuelmassloss(j) > 0 Then rate_Renamed = rate_Renamed * (fuelmassloss(j) - genrate) / fuelmassloss(j)
            Else
                rate_Renamed = species_mass_rate(j, WaterVaporYield, WallH2OYield, CeilingH2OYield, FloorH2OYield, X, mrate, 0, fuelmassloss(j), 0)
                If fuelmassloss(j) > 0 Then rate_Renamed = rate_Renamed * (fuelmassloss(j) - genrate) / fuelmassloss(j)
            End If
            'DYDX(j, 16) = derivs_species_U(ByVal  mw_upper(j), ByVal Y(j, 4), ByVal dMudt, ByVal rate, ByVal species_to_upper(7), ByVal Y(j, 16), ByVal Y(j, 17), ByVal massplume(j), ByVal Y(j, 1), ByVal Y(j, 2))
            DYDX(j, 16) = derivs_species_U(species_mass, dMudt, rate_Renamed, species_to_upper(7), Y(j, 16), Y(j, 17), massplume(j))

            'Soot Concentration
            rate_Renamed = 0
            If j = fireroom Then
                If Flashover = True And frmOptions1.optPostFlashover.Checked = True Then
                    'using postflashover model
                    'lets adjust the soot yield for underventilated conditions
                    rate_Renamed = soot_mass_rate2(j, GER, SootYieldPF, WallSootYield, CeilingSootYield, FloorSootYield, X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
                Else
                    'lets adjust the soot yield for underventilated conditions
                    rate_Renamed = soot_mass_rate(j, GER, SootYield, WallSootYield, CeilingSootYield, FloorSootYield, X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
                End If
                'If (fuelmassloss(j) > 0 And genrate > 0) Then rate = rate * (fuelmassloss(j) - genrate) / fuelmassloss(j) 'make product of combustion not product of pyrolysis
                If X = tim(i, 1) And xflag = 0 Then
                    SootMassGen(i) = rate_Renamed 'kg-soot/s in fireroom
                End If
            Else
                rate_Renamed = species_mass_rate(j, SootYield, WallSootYield, CeilingSootYield, FloorSootYield, X, mrate, 0, fuelmassloss(j), 0)
                If fuelmassloss(j) > 0 Then rate_Renamed = rate_Renamed * (fuelmassloss(j)) / fuelmassloss(j)
            End If
            'DYDX(j, 12) = derivs_species_U(ByVal mw_upper(j), ByVal Y(j, 4), ByVal dMudt, ByVal rate, ByVal species_to_upper(6), ByVal Y(j, 12), ByVal Y(j, 13), ByVal massplume(j), ByVal Y(j, 1), ByVal Y(j, 2))
            DYDX(j, 12) = derivs_species_U(species_mass, dMudt, rate_Renamed, species_to_upper(6), Y(j, 12), Y(j, 13), massplume(j))

            'HCN Concentration
            rate_Renamed = 0
            If j = fireroom Then
                If GER > 1 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object species_mass_rate2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    rate_Renamed = species_mass_rate2(GER, j, HCNYield, WallHCNYield, CeilingHCNYield, FloorHCNYield, X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object species_mass_rate(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    rate_Renamed = species_mass_rate(j, HCNYield, WallHCNYield, CeilingHCNYield, FloorHCNYield, X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
                End If
                If fuelmassloss(j) > 0 Then rate_Renamed = rate_Renamed * (fuelmassloss(j) - genrate) / fuelmassloss(j)
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object species_mass_rate(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                rate_Renamed = species_mass_rate(j, HCNYield, WallHCNYield, CeilingHCNYield, FloorHCNYield, X, mrate, 0, fuelmassloss(j), 0)
                If fuelmassloss(j) > 0 Then rate_Renamed = rate_Renamed * (fuelmassloss(j) - genrate) / fuelmassloss(j)
            End If
            'DYDX(j, 18) = derivs_species_U(ByVal mw_upper(j), ByVal Y(j, 4), ByVal dMudt, ByVal rate, ByVal species_to_upper(9), ByVal Y(j, 18), ByVal Y(j, 14), ByVal massplume(j), ByVal Y(j, 1), ByVal Y(j, 2))
            DYDX(j, 18) = derivs_species_U(species_mass, dMudt, rate_Renamed, species_to_upper(9), Y(j, 18), Y(j, 14), massplume(j))

            'TUHC concentration
            rate_Renamed = 0
            If j = fireroom Then
                'rate at which unburned fuel generated in fire room
                rate_Renamed = TUHC_Rate(hrrlimit(j), fuelmassloss(j), weighted_hc)
            Else
                'rate at which unburned fuel consumed in adjacent rooms
                If weighted_hc > 0 Then
                    rate_Renamed = -ventfire_sum(j) / (weighted_hc * 1000) + TUHC_Rate(hrrnextroom - ventfire_sum(j), fuelmassloss(j), weighted_hc)
                End If
            End If
            'DYDX(j, 6) = derivs_species_U(ByVal mw_upper(j), ByVal Y(j, 4), ByVal dMudt, ByVal rate, ByVal species_to_upper(8), ByVal Y(j, 6), ByVal Y(j, 7), ByVal massplume(j), ByVal Y(j, 1), ByVal Y(j, 2))
            DYDX(j, 6) = derivs_species_U(species_mass, dMudt, rate_Renamed, species_to_upper(8), Y(j, 6), Y(j, 7), massplume(j))
            genrate = rate_Renamed
            If genrate < 0 Then genrate = 0

            'CO concentration
            rate_Renamed = 0
            'uses combusted mass
            If j = fireroom Then rate_Renamed = COMass_Rate2(GER, X, fuelmassloss(j) - genrate, massplume(j), Y(j, 2))
            DYDX(j, 8) = derivs_species_U(species_mass, dMudt, rate_Renamed, species_to_upper(4), Y(j, 8), Y(j, 9), massplume(j))

            'combustion chemistry
            'If j = fireroom Then
            'chemistry() 'moved to do_stuff and run only once per timestep
            'End If

            'CO2 concentration
            rate_Renamed = 0
            If j = fireroom Then
                If Flashover = True And frmOptions1.optPostFlashover.Checked = True Then
                    rate_Renamed = species_mass_rate_postflashover(GER, j, CO2YieldPF, WallCO2Yield, CeilingCO2Yield, FloorCO2Yield, X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
                Else
                    If GER > 1 Then 'adjust co2 yields for underventilated fires
                        rate_Renamed = species_mass_rate2(GER, j, CO2Yield, WallCO2Yield, CeilingCO2Yield, FloorCO2Yield, X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
                    Else
                        rate_Renamed = species_mass_rate(j, CO2Yield, WallCO2Yield, CeilingCO2Yield, FloorCO2Yield, X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
                    End If
                End If
                If fuelmassloss(j) > 0 Then rate_Renamed = rate_Renamed * (fuelmassloss(j) - genrate) / fuelmassloss(j)
            Else
                rate_Renamed = species_mass_rate(j, CO2Yield, WallCO2Yield, CeilingCO2Yield, FloorCO2Yield, X, mrate, 0, fuelmassloss(j), 0)
                If fuelmassloss(j) > 0 Then rate_Renamed = rate_Renamed * (fuelmassloss(j) - genrate) / fuelmassloss(j)
            End If
            'If stepcount = 400 Then Stop
            'DYDX(j, 10) = derivs_species_U(ByVal mw_upper(j), ByVal Y(j, 4), ByVal dMudt, ByVal rate, ByVal species_to_upper(5), ByVal Y(j, 10), ByVal Y(j, 11), ByVal massplume(j), ByVal Y(j, 1), ByVal Y(j, 2))
            DYDX(j, 10) = derivs_species_U(species_mass, dMudt, rate_Renamed, species_to_upper(5), Y(j, 10), Y(j, 11), massplume(j))

            'lower layer temperature
            DYDX(j, 3) = derivs_Temp(cp_l(j), mw_lower(j), DYDX(j, 4), RoomVolume(j) - Y(j, 1), Y(j, 3), Enthalpy_L, dMldt, Y(j, 4))

            ''detector link temperature
            ''one detector per room
            'If j = fireroom Then
            '    If DetectorType(j) > 0 Then 'only calculates ceiling jet if a detector/sprinkler is installed
            '        If cjModel = cjAlpert Then

            '            DYDX(j, 19) = Unconfined_Jet(j, X, Y(j, 2), Y(j, 19), Y(j, 3), layerz(j), CJTemp, CJMax) 'alperts unconfined ceiling, link at max position
            '        ElseIf cjModel = cjJET Then

            '            DYDX(j, 19) = JET5(j, X, layerz(j), Y(j, 2), Y(j, 3), Y(j, 19), CJTemp, CJMax) 'Davis' JET correlation, hot layer effects, link at variable position
            '        End If
            '        If X = tim(i, 1) Then
            '            CJetTemp(i, 1) = CJTemp
            '            CJetTemp(i, 2) = CJMax
            '        End If
            '    Else
            '        DYDX(j, 19) = 0
            '    End If
            '    'keep the rate of temperature rise data
            'End If

            'mass used in derivs_species_L
            species_mass = species_mass_L(mw_lower(j), Y(j, 4), Y(j, 3), RoomVolume(j) - Y(j, 1))

            'TUHC
            'DYDX(j, 7) = derivs_species_L(ByVal mw_lower(j), ByVal Y(j, 4), ByVal dMldt, ByVal species_to_lower(8), ByVal Y(j, 7), ByVal Y(j, 6), ByVal massplume(j), ByVal Y(j, 3), ByVal RoomVolume(j) - Y(j, 1))
            DYDX(j, 7) = derivs_species_L(species_mass, dMldt, species_to_lower(8), Y(j, 7), Y(j, 6), massplume(j))

            'lower layer
            'CO concentration
            'DYDX(j, 9) = derivs_species_L(ByVal mw_lower(j), ByVal Y(j, 4), ByVal dMldt, ByVal species_to_lower(4), ByVal Y(j, 9), ByVal Y(j, 8), ByVal massplume(j), ByVal Y(j, 3), ByVal RoomVolume(j) - Y(j, 1))
            DYDX(j, 9) = derivs_species_L(species_mass, dMldt, species_to_lower(4), Y(j, 9), Y(j, 8), massplume(j))

            'CO2 concentration
            'DYDX(j, 11) = derivs_species_L(ByVal mw_lower(j), ByVal Y(j, 4), ByVal dMldt, ByVal species_to_lower(5), ByVal Y(j, 11), ByVal Y(j, 10), ByVal massplume(j), ByVal Y(j, 3), ByVal RoomVolume(j) - Y(j, 1))
            DYDX(j, 11) = derivs_species_L(species_mass, dMldt, species_to_lower(5), Y(j, 11), Y(j, 10), massplume(j))

            'Soot Concentration
            'DYDX(j, 13) = derivs_species_L(ByVal mw_lower(j), ByVal Y(j, 4), ByVal dMldt, ByVal species_to_lower(6), ByVal Y(j, 13), ByVal Y(j, 12), ByVal massplume(j), ByVal Y(j, 3), ByVal RoomVolume(j) - Y(j, 1))
            DYDX(j, 13) = derivs_species_L(species_mass, dMldt, species_to_lower(6), Y(j, 13), Y(j, 12), massplume(j))

            'O2 concentration
            'DYDX(j, 15) = derivs_species_L(ByVal mw_lower(j), ByVal Y(j, 4), ByVal dMldt, ByVal species_to_lower(3), ByVal Y(j, 15), ByVal Y(j, 5), ByVal massplume(j), ByVal Y(j, 3), ByVal RoomVolume(j) - Y(j, 1))
            DYDX(j, 15) = derivs_species_L(species_mass, dMldt, species_to_lower(3), Y(j, 15), Y(j, 5), massplume(j))

            'H2O concentration
            'DYDX(j, 17) = derivs_species_L(ByVal mw_lower(j), ByVal Y(j, 4), ByVal dMldt, ByVal species_to_lower(7), ByVal Y(j, 17), ByVal Y(j, 16), ByVal massplume(j), ByVal Y(j, 3), ByVal RoomVolume(j) - Y(j, 1))
            DYDX(j, 17) = derivs_species_L(species_mass, dMldt, species_to_lower(7), Y(j, 17), Y(j, 16), massplume(j))

            'HCN concentration
            'DYDX(j, 14) = derivs_species_L(ByVal mw_lower(j), ByVal Y(j, 4), ByVal dMldt, ByVal species_to_lower(9), ByVal Y(j, 14), ByVal Y(j, 18), ByVal massplume(j), ByVal Y(j, 3), ByVal RoomVolume(j) - Y(j, 1))
            DYDX(j, 14) = derivs_species_L(species_mass, dMldt, species_to_lower(9), Y(j, 14), Y(j, 18), massplume(j))

        Next j

        'do some other stuff if you want
        Erase mrate
        Erase Enthalpy
        Erase layerz
        Erase doormix
        Erase ventflow
        Erase products
        Erase ventfire_sum
        Erase matc
        Erase mw_upper
        Erase mw_lower
        Erase species_to_upper
        Erase species_to_lower
        Erase vententrain
        Erase Mass_Upper
        Erase Mass_Lower
        System.Windows.Forms.Application.DoEvents()
        If flagstop = 1 Then Exit Sub
        gTimeX = X

        Exit Sub

errorhandler:
        response = MsgBox("Error in DIFF_EQNS. Do you want to continue?", MsgBoxStyle.YesNo)

        If response = MsgBoxResult.No Then flagstop = 1
        MsgBox(ErrorToString())
        Err.Clear()
        Exit Sub
    End Sub
    Public Function species_mass_L(ByVal mw_lower As Double, ByVal pressure As Double, ByVal temp As Double, ByVal vol As Double) As Double
        On Error GoTo specieshandler
        species_mass_L = (Atm_Pressure + pressure / 1000) * mw_lower * vol / temp / Gas_Constant
        Exit Function

specieshandler:
        MsgBox(ErrorToString(Err.Number) & " in derivs_species_L")
        ERRNO = Err.Number
        'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'System.Windows.Forms.Cursor.Current = Default_Renamed
        Exit Function
    End Function
    Public Function species_mass_U(ByVal mw_upper As Double, ByVal pressure As Double, ByVal vol As Double, ByVal temp As Double) As Double
        On Error GoTo specieshandler
        species_mass_U = (Atm_Pressure + pressure / 1000) * vol * mw_upper / (Gas_Constant * temp)
        Exit Function

specieshandler:
        MsgBox(ErrorToString(Err.Number) & " in species_mass_U")
        ERRNO = Err.Number
        'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'System.Windows.Forms.Cursor.Current = Default_Renamed
        Exit Function

    End Function

    Public Sub STIFF3(ByVal n As Integer, ByRef YI(,) As Double, ByVal x1 As Double, ByVal x2 As Double, ByVal EPS As Double, ByVal H1 As Double, ByVal HMIN As Double)

        ' Solution of stiff differential equations subroutine
        ' Input
        '  N    = number of equations
        '  YI() = initial values of integration vector
        '         ( N rows by 1 column )
        '  X1   = initial X value
        '  X2   = final X value
        '  EPS  = convergence criteria
        '  H1   = initial guess for step size
        '  HMIN = minimum step size
        ' Output
        '  YI() = final values of solution vector
        '         ( N rows by 1 column )
        ' NOTE: requires subroutines LUD, LUB, DIFFEQS and JACOBIAN

        Static Yscal(,) As Double
        Static Y(,) As Double
        Static DYDX(,) As Double
        Static j As Integer
        Static A(,,) As Double
        Static DFDX(,) As Double
        Static DFDY(,,) As Double
        Static DYSAV(,) As Double
        Static YERR(,) As Double
        Static G1(,) As Double
        Static G2(,) As Double
        Static G3(,) As Double
        Static G4(,) As Double
        Static YSAV(,) As Double
        Static index(,) As Double
        ReDim Yscal(NumberRooms + 1, n)
        ReDim Y(NumberRooms + 1, n)
        ReDim DYDX(NumberRooms + 1, n)
        ReDim A(NumberRooms, n, n)
        ReDim DFDX(NumberRooms + 1, n)
        ReDim DFDY(NumberRooms, n, n)
        ReDim DYSAV(NumberRooms + 1, n)
        ReDim YERR(NumberRooms, n)
        ReDim G1(NumberRooms, n)
        ReDim G2(NumberRooms, n)
        ReDim G3(NumberRooms, n)
        ReDim G4(NumberRooms, n)
        ReDim YSAV(NumberRooms + 1, n)
        ReDim index(NumberRooms, n)

        Static a31, gamma2, ICOEFF, a21, a32 As Double
        Static c42, c32, c21, c31, c41, c43 As Double
        Static e3, e1, b3, b1, b2, b4, e2, e4 As Double
        Static a2x, c3x, c1x, c2x, c4x, a3x As Double
        Static X, h, xsav As Double
        Static i As Integer
        Static ier As Integer
        Static errtmp, errmax, hnext As Double
        Static k As Integer
        Static eps2 As Double

        eps2 = EPS

start:
        If (ICOEFF = 0) Then
            gamma2 = 0.231
            a21 = 2
            a31 = 4.52470820736
            a32 = 4.1635287886
            c21 = -5.07167533877
            c31 = 6.02015272865
            c32 = 0.159750684673
            c41 = -1.856343618677
            c42 = -8.50538085819
            c43 = -2.08407513602
            b1 = 3.95750374663
            b2 = 4.62489238836
            b3 = 0.617477263873
            b4 = 1.282612945268
            e1 = -2.30215540292
            e2 = -3.07363448539
            e3 = 0.873280801802
            e4 = 1.282612945268
            c1x = -0.039629667752
            c2x = -0.039629667752
            c3x = 0.550778939579
            c4x = -0.05535098457
            a2x = 0.462
            a3x = 0.880208333333

            ICOEFF = 1
        End If

        ' compute step size "direction"
        h = Math.Sign(x2 - x1) * H1

        ' save initial conditions
        X = x1

        For i = 1 To n
            'Y(1, i) = YI(i)
            For j = 1 To NumberRooms '%%% added
                Y(j, i) = YI(j, i) '%%% does this need to change for multi-rooms?
            Next j '%%% added
        Next i


        Do

            Call diff_eqns(X, Y, DYDX) 'added back 4/3/07
            'Call diff_eqns2(ByVal X, Y(), DYDX(), False, 0, 0)
            If flagstop = 1 Then GoTo finish

            ' compute scaling vector
            For i = 1 To n
                For j = 1 To NumberRooms
                    Yscal(j, i) = Abs(Y(j, i)) + Abs(h * DYDX(j, i)) + 1.0E-30

                Next j
            Next i

            ' compute last step size
            If ((X + h - x2) * (X + h - x1) > 0) Then h = x2 - X

            ' perform integration
            xsav = X

            For i = 1 To n
                For j = 1 To NumberRooms
                    YSAV(j, i) = Y(j, i)
                    DYSAV(j, i) = DYDX(j, i)
                Next j
            Next i

            Call JACOBIAN(xsav, YSAV, DFDX, DFDY, n)

            Do
                For k = 1 To NumberRooms
                    For i = 1 To n
                        For j = 1 To n
                            A(k, i, j) = -DFDY(k, i, j)
                        Next j
                        A(k, i, i) = 1 / (gamma2 * h) + A(k, i, i)
                    Next i
                Next k

                Call LUD(n, A, index, ier)

                For i = 1 To n
                    For j = 1 To NumberRooms
                        G1(j, i) = DYSAV(j, i) + h * c1x * DFDX(j, i)
                    Next j
                Next i

                Call LUB(n, A, index, G1)

                For i = 1 To n
                    For j = 1 To NumberRooms
                        Y(j, i) = YSAV(j, i) + a21 * G1(j, i)
                    Next j
                Next i

                X = xsav + a2x * h

                'Call diff_eqns2(ByVal X, Y(), DYDX(), False, 0, 0)
                Call diff_eqns(X, Y, DYDX)
                If flagstop = 1 Then GoTo finish 'terminate early

                For i = 1 To n
                    For j = 1 To NumberRooms
                        G2(j, i) = DYDX(j, i) + h * c2x * DFDX(j, i) + c21 * G1(j, i) / h
                    Next j
                Next i

                Call LUB(n, A, index, G2)

                For i = 1 To n
                    For j = 1 To NumberRooms
                        Y(j, i) = YSAV(j, i) + a31 * G1(j, i) + a32 * G2(j, i)
                    Next j
                Next i

                X = xsav + a3x * h

                Call diff_eqns(X, Y, DYDX)
                'Call diff_eqns2(ByVal X, Y(), DYDX(), False, 0, 0)
                If flagstop = 1 Then GoTo finish

                For i = 1 To n
                    For j = 1 To NumberRooms
                        G3(j, i) = DYDX(j, i) + h * c3x * DFDX(j, i) + (c31 * G1(j, i) + c32 * G2(j, i)) / h
                    Next j
                Next i

                Call LUB(n, A, index, G3)

                For i = 1 To n
                    For j = 1 To NumberRooms
                        G4(j, i) = DYDX(j, i) + h * c4x * DFDX(j, i) + (c41 * G1(j, i) + c42 * G2(j, i) + c43 * G3(j, i)) / h
                    Next j
                Next i

                Call LUB(n, A, index, G4)

                For j = 1 To NumberRooms
                    For i = 1 To n
                        Y(j, i) = YSAV(j, i) + b1 * G1(j, i) + b2 * G2(j, i) + b3 * G3(j, i) + b4 * G4(j, i)
                        YERR(j, i) = e1 * G1(j, i) + e2 * G2(j, i) + e3 * G3(j, i) + e4 * G4(j, i)
                    Next i
                Next j

                X = xsav + h

                If (X = xsav) Then
                    MsgBox("Sorry, the ODE Solver cannot converge. Insignificant step size.")
                    flagstop = 1
                    Exit Sub
                End If

                errmax = 0

                For i = 1 To n
                    For j = 1 To NumberRooms
                        errtmp = Abs(YERR(j, i) / Yscal(j, i))
                        If (errtmp > errmax) Then errmax = errtmp
                    Next j
                Next i

                errmax = errmax / eps2

                If (errmax <= 1) Then
                    If (errmax > 0.1296) Then
                        hnext = 0.9 * h * errmax ^ (-0.25)
                    Else
                        hnext = 1.5 * h
                    End If
                    Exit Do
                Else
                    hnext = 0.9 * h * errmax ^ (-1 / 3)
                    If (hnext < 0.5 * h) Then hnext = 0.5 * h
                    h = hnext
                End If
            Loop

            ' check for end of integration
            If ((X - x2) * (x2 - x1) >= 0) Then
                For i = 1 To n
                    'YI(i) = Y(1, i)
                    For j = 1 To NumberRooms
                        YI(j, i) = Y(j, i) '%%% multirooms???
                    Next j
                Next i
                Exit Do
            End If

            ' check for step size too small
            If (Abs(hnext) < HMIN) Then
                For i = 1 To n

                    If stepcount = 1 Then
                        ' HeatRelease(fireroom, 1, 1) = 0.1
                    Else
                        MsgBox("Sorry, the ODE Solver cannot converge. Attempted step size is too small")
                        flagstop = 1
                        Exit Sub
                    End If


                    For j = 1 To NumberRooms
                        YI(j, i) = Y(j, i)
                    Next j
                Next i
                Exit Do
                'End If
                'GoTo start
            End If

            'set next step size
            h = hnext
        Loop

finish:
        ' erase working arrays
        Erase Yscal
        Erase DYDX
        Erase A
        Erase DFDX
        Erase DFDY
        Erase DYSAV
        Erase YERR
        Erase G1
        Erase G2
        Erase G3
        Erase G4
        Erase YSAV
        Erase index

    End Sub
	
    Public Function soot_mass_rate2(ByVal room As Integer, ByVal GER As Double, ByRef Yield As Single, ByRef WallYield() As Double, ByRef CeilingYield() As Double, ByRef FloorYield() As Double, ByRef T As Double, ByRef mrate() As Double, ByRef mrate_wall As Double, ByRef mrate_ceiling As Double, ByRef mrate_floor As Double) As Double
        '*  ====================================================================
        '*  This function returns the value of the sum of the fuel mass loss rates
        '*  multiplied by the species generation rate for all burning objects.
        '*  ====================================================================

        Dim dummy As Double
        Dim total As Single
        'Dim i As Integer

        If GER >= 0.1 Then
            SootFactor = 1 + SootAlpha / (Exp(2.5 * GER ^ (-SootEpsilon))) '13/3/2003
        Else
            SootFactor = 1
        End If

        '2007.13
        If frmOptions1.optSootman.Checked = True Then 'manual entry of pre post flashover yields
            SootFactor = 1
            If GER > 1 Then
                Yield = postSoot
            Else
                Yield = preSoot
            End If
        End If
        '-----------------------------------

        If room = fireroom Then
            'For i = 1 To NumberObjects
            If GER > 1.1 Then
                dummy = mrate(1) * Yield * SootFactor
                If useCLTmodel = True Then dummy = (mrate(1) - mrate_wall - mrate_ceiling - mrate_floor) * Yield * SootFactor
                'If i = 1 Then
                dummy = dummy + mrate_wall * WallYield(room) * SootFactor + mrate_ceiling * CeilingYield(room) * SootFactor + mrate_floor * FloorYield(room) * SootFactor
                'End If
            Else
                dummy = mrate(1) * Yield
                If useCLTmodel = True Then dummy = (mrate(1) - mrate_wall - mrate_ceiling - mrate_floor) * Yield
                'If i = 1 Then
                dummy = dummy + mrate_wall * WallYield(room) + mrate_ceiling * CeilingYield(room) + mrate_floor * FloorYield(room)
                'End If
            End If
            total = total + dummy
            'Next i
        Else
            total = mrate_ceiling * CeilingYield(room)
        End If
        soot_mass_rate2 = total 'kg-species per sec

    End Function
	
	Public Function cv_gas(ByRef gas As String, ByRef temp As Double) As Double
		'==============================================
		' look up the specific heat capacity for a given combustion product
		' c wade 25/3/2003
		'================================================
		
		Select Case gas
			Case "N2"
				If temp >= 1450 Then
					cv_gas = 0.94709 : Exit Function
				ElseIf temp >= 1350 Then 
					cv_gas = 0.93552 : Exit Function
				ElseIf temp >= 1250 Then 
					cv_gas = 0.92231 : Exit Function
				ElseIf temp >= 1150 Then 
					cv_gas = 0.9072 : Exit Function
				ElseIf temp >= 1050 Then 
					cv_gas = 0.88999 : Exit Function
				ElseIf temp >= 950 Then 
					cv_gas = 0.87052 : Exit Function
				ElseIf temp >= 850 Then 
					cv_gas = 0.84884 : Exit Function
				ElseIf temp >= 750 Then 
					cv_gas = 0.82535 : Exit Function
				ElseIf temp >= 650 Then 
					cv_gas = 0.8011 : Exit Function
				ElseIf temp >= 550 Then 
					cv_gas = 0.77807 : Exit Function
				ElseIf temp >= 450 Then 
					cv_gas = 0.75921 : Exit Function
				ElseIf temp >= 350 Then 
					cv_gas = 0.74746 : Exit Function
				Else
					cv_gas = 0.74317 : Exit Function
				End If
			Case "H2O"
				If temp >= 1150 Then
					cv_gas = 1.9682 : Exit Function
				ElseIf temp >= 1050 Then 
					cv_gas = 1.9 : Exit Function
				ElseIf temp >= 950 Then 
					cv_gas = 1.8297 : Exit Function
				ElseIf temp >= 850 Then 
					cv_gas = 1.7589 : Exit Function
				ElseIf temp >= 750 Then 
					cv_gas = 1.6892 : Exit Function
				ElseIf temp >= 650 Then 
					cv_gas = 1.6223 : Exit Function
				ElseIf temp >= 550 Then 
					cv_gas = 1.56 : Exit Function
				ElseIf temp >= 450 Then 
					cv_gas = 1.5084 : Exit Function
				ElseIf temp >= 373 Then 
					cv_gas = 1.5091 : Exit Function
				ElseIf temp > 300 Then 
					cv_gas = 3.7683 : Exit Function
				Else
					cv_gas = 4.1302 : Exit Function
				End If
			Case "CO2"
				If temp >= 1050 Then
					cv_gas = 1.0702 : Exit Function
				ElseIf temp >= 950 Then 
					cv_gas = 1.0452 : Exit Function
				ElseIf temp >= 850 Then 
					cv_gas = 1.0155 : Exit Function
				ElseIf temp >= 750 Then 
					cv_gas = 0.97997 : Exit Function
				ElseIf temp >= 650 Then 
					cv_gas = 0.93752 : Exit Function
				ElseIf temp >= 550 Then 
					cv_gas = 0.88668 : Exit Function
				ElseIf temp >= 450 Then 
					cv_gas = 0.82552 : Exit Function
				ElseIf temp >= 350 Then 
					cv_gas = 0.751 : Exit Function
				Else
					cv_gas = 0.65935 : Exit Function
				End If
			Case "CO"
				If temp >= 950 Then
					cv_gas = 0.75514 : Exit Function
				ElseIf temp >= 850 Then 
					cv_gas = 0.75186 : Exit Function
				ElseIf temp >= 750 Then 
					cv_gas = 0.74872 : Exit Function
				ElseIf temp >= 650 Then 
					cv_gas = 0.74593 : Exit Function
				ElseIf temp >= 550 Then 
					cv_gas = 0.74373 : Exit Function
				ElseIf temp >= 450 Then 
					cv_gas = 0.74228 : Exit Function
				ElseIf temp >= 350 Then 
					cv_gas = 0.74169 : Exit Function
				Else
					cv_gas = 0.74194 : Exit Function
				End If
			Case "O2"
				If temp >= 950 Then
					cv_gas = 0.82993 : Exit Function
				ElseIf temp >= 850 Then 
					cv_gas = 0.81395 : Exit Function
				ElseIf temp >= 750 Then 
					cv_gas = 0.79453 : Exit Function
				ElseIf temp >= 650 Then 
					cv_gas = 0.77103 : Exit Function
				ElseIf temp >= 550 Then 
					cv_gas = 0.74317 : Exit Function
				ElseIf temp >= 450 Then 
					cv_gas = 0.71193 : Exit Function
				ElseIf temp >= 350 Then 
					cv_gas = 0.68116 : Exit Function
				Else
					cv_gas = 0.65873 : Exit Function
				End If
			Case Else
		End Select
		MsgBox("cv not found - error in STIFFSOLV")
		
	End Function

    Private Sub ventflows(ByRef Mass_Upper() As Double, ByRef vententrain(,,) As Double, ByVal tim As Double, ByRef Y(,) As Double, ByRef ventflow(,) As Double, ByRef layerz() As Double, ByRef Enthalpy(,) As Double, ByRef products(,,) As Double, ByRef ventfire_sum() As Double, ByRef mw_upper() As Double, ByRef mw_lower() As Double, ByVal weighted_hc As Double, ByRef FlowOut() As Double, ByRef Flowtoroom() As Double, ByRef cp_u() As Double, ByRef cp_l() As Double)
        '====== ===============================================================
        '   VENT FLOW ROUTINE BASED ON PRESSURE
        '   C A WADE 21/8/98
        ' 1 feb 2008 Amanda
        '=====================================================================
        Dim DischargeCoeff As Double
        Dim SprinkCoolCoeff As Double
        Dim NProd, D, k, i, h, j, L, nelev, mxprd As Integer
        Dim xmslab() As Double
        Dim dpv1m2(9) As Double
        Dim yvelev(9) As Double
        Dim nvelev As Integer
        Dim nslab As Integer
        Dim yflor(2) As Double
        Dim yceil(2) As Double
        Dim denu(2) As Double
        Dim denl(2) As Double
        Dim r1u As Double
        Dim r1l As Double
        Dim ylay(2) As Double
        Dim pflor(2) As Double
        Dim TU(2) As Double
        Dim TL(2) As Double
        Dim r2u As Double
        Dim r2l As Double
        Dim yelev(9) As Double
        Dim dp1m2(9) As Double
        Dim pab1(9) As Double
        Dim pab2(9) As Double
        Dim yvbot, epsp, yvtop, avent As Double
        Dim conl(,) As Double
        Dim conu(,) As Double
        Dim cslab(,) As Double
        Dim pslab(,) As Double
        Dim qslab() As Double
        Dim yslab() As Double
        Dim dirs12() As Double
        Dim m, mxslab As Integer
        Dim dteps, rlam As Double
        Dim tslab() As Double
        Dim UFLW2(,,) As Double
        Dim QSUVNT() As Double
        Dim QSLVNT() As Double
        Dim e2u, e1u, neutralplane, e1l, e2l As Double
        Dim r2utemp, r1utemp, r1ltemp, r2ltemp As Double
        Dim prod1U() As Double
        Dim prod2U() As Double
        Dim prod1L() As Double
        Dim prod2L() As Double
        Dim term, vp3, vp1, vp, vp2, vent_entrain, zp As Double
        Dim doormix(2) As Double
        Dim YVent, ventfire, FU As Double
        Dim cpu(2) As Double
        Dim cpl(2) As Double
        Dim uvol(2) As Double
        Dim cvent_Renamed(,) As Double
        Dim C(,) As Double
        Dim PL(,) As Double
        Dim PU(,) As Double
        Dim PVent(,) As Double
        Dim P() As Double
        Dim XML(2) As Double
        Dim XMU(2) As Double
        Dim XMVent(2) As Double
        Dim QL(2) As Double
        Dim QU(2) As Double
        Dim QVent(2) As Double
        Dim T(2) As Double
        Dim Tvent(2) As Double
        Dim DELP1 As Double
        Dim Den(2) As Double

        Static tlast As Double
        Static FirstTime, FirstTime2, FirstTime3 As Boolean

        On Error GoTo errorhandler
        'If stepcount = 1002 Then Stop

        k = stepcount

        If gb_first_time_vent = True Then 'If tim = 0 Then
            FirstTime = True
            FirstTime2 = True
            FirstTime3 = True
        End If
        NProd = 7 'number of species
        mxprd = 7 'max number of products
        epsp = Error_Control_ventflow
        'epsp = 0.001
        'error tolerance
        mxslab = 6
        dteps = 3 'K difference in relative temps before layer will form
        rlam = 0

        ReDim conl(mxprd, NumberRooms + 1)
        ReDim conu(mxprd, NumberRooms + 1)
        ReDim cslab(mxslab, mxprd)
        ReDim pslab(mxslab, mxprd)
        ReDim xmslab(mxslab)
        ReDim yslab(mxslab)
        ReDim qslab(mxslab)
        ReDim UFLW2(2, NProd + 2, 2)
        ReDim tslab(mxslab)
        ReDim dirs12(mxslab)
        ReDim QSUVNT(2)
        ReDim QSLVNT(2)
        ReDim ventflow(NumberRooms + 1, 2)
        ReDim Enthalpy(NumberRooms + 1, 2)
        ReDim products(NProd + 2, NumberRooms + 1, 2)
        ReDim prod1U(NProd + 2)
        ReDim prod2U(NProd + 2)
        ReDim prod1L(NProd + 2)
        ReDim prod2L(NProd + 2)
        ReDim cvent_Renamed(NProd, 2)
        ReDim C(NProd, 2)
        ReDim PL(NProd, 2)
        ReDim PU(NProd, 2)
        ReDim PVent(NProd, 2)
        ReDim P(2)

        For i = 1 To NumberRooms + 1 'first room
            For j = 1 To NumberRooms + 1 'second room
                If i <> j Then
                    r1u = 0
                    r1l = 0
                    r2u = 0
                    r2l = 0
                    If i < NumberRooms + 1 And j > 1 Then 'i is an inside room
                        If i < j Then
                            If NumberVents(i, j) > 0 Then 'horizontal vents exist
                                'setup routines
                                uvol(1) = Y(i, 1)
                                yflor(1) = FloorElevation(i)
                                yceil(1) = RoomHeight(i) + FloorElevation(i)
                                ylay(1) = layerz(i) + FloorElevation(i)
                                denu(1) = (Atm_Pressure + Y(i, 4) / 1000) / (Gas_Constant / mw_upper(i)) / Y(i, 2)
                                denl(1) = (Atm_Pressure + Y(i, 4) / 1000) / (Gas_Constant / mw_lower(i)) / Y(i, 3)
                                pflor(1) = Y(i, 4)
                                TU(1) = Y(i, 2)
                                TL(1) = Y(i, 3)
                                cpu(1) = cp_u(i)
                                cpl(1) = cp_l(i)
                                conl(1, 1) = Y(i, 15) 'lower oxygen mass fraction
                                conu(1, 1) = Y(i, 5) 'upper oxygen
                                conl(2, 1) = Y(i, 9) 'lower co
                                conu(2, 1) = Y(i, 8) 'upper co
                                conl(3, 1) = Y(i, 11) 'lower co2
                                conu(3, 1) = Y(i, 10) 'upper co2
                                conl(4, 1) = Y(i, 13) 'lower soot
                                conu(4, 1) = Y(i, 12) 'upper soot
                                conl(5, 1) = Y(i, 17) 'lower water
                                conu(5, 1) = Y(i, 16) 'upper water
                                conl(6, 1) = Y(i, 7) 'lower tuhc
                                conu(6, 1) = Y(i, 6) 'upper tuhc
                                conl(7, 1) = Y(i, 14) 'lower hcn
                                conu(7, 1) = Y(i, 18) 'upper hcn

                                If j = NumberRooms + 1 Then 'outside
                                    yflor(2) = yflor(1)
                                    yceil(2) = yceil(1)
                                    ylay(2) = yceil(1)
                                    pflor(2) = RoomPressure(i, 1)
                                    Call Outside_Conditions(conl, conu, TU, TL, cpu, cpl, pflor, denu, denl, 2)
                                Else 'j is an inside room
                                    uvol(2) = Y(j, 1)
                                    yflor(2) = FloorElevation(j)
                                    yceil(2) = RoomHeight(j) + FloorElevation(j)
                                    ylay(2) = layerz(j) + FloorElevation(j)
                                    denu(2) = (Atm_Pressure + Y(j, 4) / 1000) / (Gas_Constant / mw_upper(j)) / Y(j, 2)
                                    denl(2) = (Atm_Pressure + Y(j, 4) / 1000) / (Gas_Constant / mw_lower(j)) / Y(j, 3)
                                    pflor(2) = Y(j, 4)
                                    TU(2) = Y(j, 2)
                                    TL(2) = Y(j, 3)
                                    cpu(2) = cp_u(j)
                                    cpl(2) = cp_l(j)
                                    conl(1, 2) = Y(j, 15) 'lower oxygen
                                    conu(1, 2) = Y(j, 5) 'upper oxygen
                                    conl(2, 2) = Y(j, 9) 'lower co
                                    conu(2, 2) = Y(j, 8) 'upper co
                                    conl(3, 2) = Y(j, 11) 'lower co2
                                    conu(3, 2) = Y(j, 10) 'upper co2
                                    conl(4, 2) = Y(j, 13) 'lower soot
                                    conu(4, 2) = Y(j, 12) 'upper soot
                                    conl(5, 2) = Y(j, 17) 'lower water
                                    conu(5, 2) = Y(j, 16) 'upper water
                                    conl(6, 2) = Y(j, 7) 'lower tuhc
                                    conu(6, 2) = Y(j, 6) 'upper tuhc
                                    conl(7, 2) = Y(j, 14) 'lower hcn
                                    conu(7, 2) = Y(j, 18) 'upper hcn
                                End If

                                Call CommonWall(yflor, yceil, ylay, denu, denl, pflor, nelev, yelev, dp1m2, pab1, pab2)

                                'which vent
                                r1u = 0
                                r1l = 0
                                r2u = 0
                                r2l = 0
                                e1u = 0
                                e1l = 0
                                e2u = 0
                                e2l = 0
                                For h = 3 To NProd + 2
                                    prod1U(h) = 0
                                    prod1L(h) = 0
                                    prod2U(h) = 0
                                    prod2L(h) = 0
                                Next h
                                avent = 0

                                'considering vertical vents
                                For m = 1 To NumberVents(i, j)

                                    avent = ventarea(tim, i, j, m) 'determines vent area and if vent closed sets area to zero
                                    If avent > gcd_Machine_Error Then 'vent is open between the rooms
                                        yvtop = VentHeight(i, j, m) + VentSillHeight(i, j, m) + FloorElevation(i)
                                        yvbot = VentSillHeight(i, j, m) + FloorElevation(i)
                                        'If m = 2 Then Stop
                                        'sprinkler cooling coefficient
                                        If SprinklerFlag = 1 And i = fireroom Then
                                            'SprinkCoolCoeff = gcs_SprinkCoolCoeff 'sprinkler has operated
                                            SprinkCoolCoeff = SprCooling 'sprinkler has operated
                                        Else
                                            SprinkCoolCoeff = 1.0 'no effect modelled
                                        End If

                                        DischargeCoeff = VentCD(i, j, m) * SprinkCoolCoeff

                                        If DischargeCoeff > 0 Then

                                            Call VENTHP(DischargeCoeff, yflor, ylay, TU, TL, cpu, cpl, denl, denu, pflor, yvtop, yvbot, avent, yelev, dp1m2, nelev, pab1, pab2, conl, conu, NProd, epsp, cslab, pslab, tslab, qslab, yslab, dirs12, dpv1m2, yvelev, xmslab, nvelev, nslab)

                                            If frmOptions1.optCCFM.Checked = True Then
                                                'use vent flow depostion rules from CCFM
                                                Call FLogo2(uvol, WAfloor_flag(i, j, m), dirs12, NProd, yslab, yslab, xmslab, tslab, nslab, TU, TL, yflor, yceil, ylay, qslab, rlam, pslab, mxprd, mxslab, dteps, QSLVNT, QSUVNT, UFLW2)
                                            Else
                                                'use vent flow deposition rules from CFAST 3.17
                                                Call FLogo1(i, j, uvol, WAfloor_flag(i, j, m), dirs12, NProd, yslab, yslab, xmslab, tslab, nslab, TU, TL, yflor, yceil, ylay, qslab, rlam, pslab, mxprd, mxslab, dteps, QSLVNT, QSUVNT, UFLW2)
                                                'Call FLogo2(uvol, WAfloor_flag(i, j, m), dirs12, NProd, yslab, yslab, xmslab, tslab, nslab, TU, TL, yflor, yceil, ylay, qslab, rlam, pslab, mxprd, mxslab, dteps, QSLVNT, QSUVNT, UFLW2)
                                            End If

                                            r1u = r1u + UFLW2(1, 1, 2)
                                            r1l = r1l + UFLW2(1, 1, 1)
                                            r2u = r2u + UFLW2(2, 1, 2)
                                            r2l = r2l + UFLW2(2, 1, 1)
                                            e1u = e1u + UFLW2(1, 2, 2)
                                            e1l = e1l + UFLW2(1, 2, 1)
                                            e2u = e2u + UFLW2(2, 2, 2)
                                            e2l = e2l + UFLW2(2, 2, 1)
                                            For h = 3 To NProd + 2
                                                prod1U(h) = prod1U(h) + UFLW2(1, h, 2) 'net flow of species kg/s
                                                prod1L(h) = prod1L(h) + UFLW2(1, h, 1)
                                                prod2U(h) = prod2U(h) + UFLW2(2, h, 2)
                                                prod2L(h) = prod2L(h) + UFLW2(2, h, 1)
                                            Next h

                                            'keep track of the volumetric vent flow from each room to the outside, excluding vent mixing terms
                                            If r1u < 0 Then FlowOut(i) = FlowOut(i) - r1u / denu(1)

                                            ventfire = 0
                                            'Call door_mixing_flow(NProd, nelev, m, i, j, conu, prod2U, prod2L, prod1U, prod1L, TU, TL, cpu, cpl, e1u, e1l, e2u, e2l, r1u, r1l, r2l, r2u, ylay, UFLW2, dp1m2, yelev)

                                            'new - Utiskul and Quintiere
                                            Call near_vent_mixing(NProd, nelev, m, i, j, conu, prod2U, prod2L, prod1U, prod1L, TU, TL, cpu, cpl, e1u, e1l, e2u, e2l, r1u, r1l, r2l, r2u, ylay, UFLW2, dp1m2, yelev)

                                            'i=first room j=second room, m=vent id
                                            Call door_entrain4(dteps, Mass_Upper, vententrain, m, i, NProd, prod1L, prod2L, prod2U, prod1U, conu, conl, TU, TL, cpu, cpl, j, e1l, e1u, e2l, e2u, r2u, r2l, r1u, r1l, ylay, UFLW2, ventfire, mw_upper, mw_lower, weighted_hc, nelev, dp1m2, yelev)

                                            'write to a vent flow log
                                            If ventlog = True And (tim > tlast Or FirstTime = True) Then
                                                If tim - Int(tim) = 0 Then
                                                    Call Ventflow_log2(FirstTime, tim, i, j, m, nslab, xmslab, dirs12, yvelev, vententrain)
                                                End If
                                            End If
                                            If ventlog = True And (tim > tlast Or FirstTime3 = True) Then
                                                If tim - Int(tim) = 0 Then
                                                    If i = 1 And j = 2 And m = 1 Then
                                                        Call Ventflow_log_csv(FirstTime3, tim, i, j, m, nslab, xmslab, dirs12, yvelev, vententrain)
                                                    End If
                                                End If
                                            End If

                                            'keep track of flow to the fire room, for use in GER calculation
                                            If i = fireroom Then
                                                For D = 1 To nslab
                                                    If dirs12(D) < 0 Then
                                                        Flowtoroom(i) = xmslab(D)
                                                    End If
                                                Next
                                            End If

                                            If r1u < 0 Then 'ventfire burning in room 2
                                                ventfire_sum(j) = ventfire_sum(j) + ventfire
                                            Else 'ventfire burning in room 1
                                                ventfire_sum(i) = ventfire_sum(i) + ventfire
                                            End If
                                        End If
                                    End If
                                Next m ' moved this down 24/10/2003

                                ventflow(i, 1) = ventflow(i, 1) + r1l '1st room lower layer
                                ventflow(i, 2) = ventflow(i, 2) + r1u '1st room upper layer
                                ventflow(j, 1) = ventflow(j, 1) + r2l '2nd room lower layer
                                ventflow(j, 2) = ventflow(j, 2) + r2u '2nd room upper layer
                                Enthalpy(i, 1) = Enthalpy(i, 1) + e1l '1st room lower layer
                                Enthalpy(i, 2) = Enthalpy(i, 2) + e1u '1st room upper layer
                                Enthalpy(j, 1) = Enthalpy(j, 1) + e2l '2nd room lower layer
                                Enthalpy(j, 2) = Enthalpy(j, 2) + e2u '2nd room upper layer
                                For h = 3 To NProd + 2
                                    products(h, i, 1) = products(h, i, 1) + prod1L(h) '1st room lower layer
                                    products(h, i, 2) = products(h, i, 2) + prod1U(h) '1st room upper layer
                                    products(h, j, 1) = products(h, j, 1) + prod2L(h) '2nd room lower layer
                                    products(h, j, 2) = products(h, j, 2) + prod2U(h) '2nd room upper layer
                                Next h
                            End If
                        End If
                    End If

                    'now do the ceiling/floor vents
                    ' i is room above vent, j is room below vent
                    If NumberCVents(i, j) > 0 Then
                        avent = 0
                        r1u = 0
                        r1l = 0
                        r2u = 0
                        r2l = 0
                        e1u = 0
                        e1l = 0
                        e2u = 0
                        e2l = 0
                        For h = 3 To NProd + 2
                            prod1U(h) = 0
                            prod1L(h) = 0
                            prod2U(h) = 0
                            prod2L(h) = 0
                        Next h

                        'upper space conditions
                        If i <> NumberRooms + 1 Then 'inside room
                            conl(1, 1) = Y(i, 15) 'lower oxygen mass fraction
                            conu(1, 1) = Y(i, 5) 'upper oxygen
                            conl(2, 1) = Y(i, 9) 'lower co
                            conu(2, 1) = Y(i, 8) 'upper co
                            conl(3, 1) = Y(i, 11) 'lower co2
                            conu(3, 1) = Y(i, 10) 'upper co2
                            conl(4, 1) = Y(i, 13) 'lower soot
                            conu(4, 1) = Y(i, 12) 'upper soot
                            conl(5, 1) = Y(i, 17) 'lower water
                            conu(5, 1) = Y(i, 16) 'upper water
                            conl(6, 1) = Y(i, 7) 'lower tuhc
                            conu(6, 1) = Y(i, 6) 'upper tuhc
                            conl(7, 1) = Y(i, 14) 'lower hcn
                            conu(7, 1) = Y(i, 18) 'upper hcn
                            TU(1) = Y(i, 2) 'upper layer temp
                            TL(1) = Y(i, 3) 'lower layer temp
                            cpu(1) = cp_u(i)
                            cpl(1) = cp_l(i)
                            uvol(1) = Y(i, 1)
                            yflor(1) = FloorElevation(i)
                            yceil(1) = RoomHeight(i) + FloorElevation(i)
                            ylay(1) = layerz(i) + FloorElevation(i)
                            pflor(1) = Y(i, 4)
                            denu(1) = (Atm_Pressure + Y(i, 4) / 1000) / (Gas_Constant / mw_upper(i)) / Y(i, 2)
                            denl(1) = (Atm_Pressure + Y(i, 4) / 1000) / (Gas_Constant / mw_lower(i)) / Y(i, 3)
                        Else 'outside space
                            Call Outside_Conditions(conl, conu, TU, TL, cpu, cpl, pflor, denu, denl, 1)
                            yflor(1) = FloorElevation(j)
                            yceil(1) = yflor(1)
                            ylay(1) = yceil(1)
                            pflor(1) = -G * yflor(1) * (Atm_Pressure) / (Gas_Constant / MW_air) / ExteriorTemp
                        End If

                        'lower space conditions
                        If j <> NumberRooms + 1 Then 'inside room
                            conl(1, 2) = Y(j, 15) 'lower oxygen mass fraction
                            conu(1, 2) = Y(j, 5) 'upper oxygen
                            conl(2, 2) = Y(j, 9) 'lower co
                            conu(2, 2) = Y(j, 8) 'upper co
                            conl(3, 2) = Y(j, 11) 'lower co2
                            conu(3, 2) = Y(j, 10) 'upper co2
                            conl(4, 2) = Y(j, 13) 'lower soot
                            conu(4, 2) = Y(j, 12) 'upper soot
                            conl(5, 2) = Y(j, 17) 'lower water
                            conu(5, 2) = Y(j, 16) 'upper water
                            conl(6, 2) = Y(j, 7) 'lower tuhc
                            conu(6, 2) = Y(j, 6) 'upper tuhc
                            conl(7, 2) = Y(j, 14) 'lower hcn
                            conu(7, 2) = Y(j, 18) 'upper hcn
                            TU(2) = Y(j, 2) 'upper layer temp
                            TL(2) = Y(j, 3) 'lower layer temp
                            cpu(2) = cp_u(j)
                            cpl(2) = cp_l(j)
                            uvol(2) = Y(j, 1)
                            yflor(2) = FloorElevation(j)
                            yceil(2) = RoomHeight(j) + FloorElevation(j)
                            ylay(2) = layerz(j) + FloorElevation(j)
                            pflor(2) = Y(j, 4)
                            denu(2) = (Atm_Pressure + Y(j, 4) / 1000) / (Gas_Constant / mw_upper(j)) / Y(j, 2)
                            denl(2) = (Atm_Pressure + Y(j, 4) / 1000) / (Gas_Constant / mw_lower(j)) / Y(j, 3)
                        Else 'outside space
                            Call Outside_Conditions(conl, conu, TU, TL, cpu, cpl, pflor, denu, denl, 2)
                            yflor(2) = FloorElevation(i)
                            ylay(2) = yflor(2)
                            yceil(2) = yflor(2)
                            pflor(2) = -G * yflor(2) * (Atm_Pressure) / (Gas_Constant / MW_air) / ExteriorTemp
                        End If

                        If gb_first_time_vent = True Then 'only perform this the first time through ventflows
                            'Call check_vent_areas_C(i, j) 'check for vent area compared to room floor area once per combination of rooms
                            Call check_cvent_areas(i, j) 'colleen 2/6/09 fixed prob in 2009.1 for multpile ceiling vents
                        End If

                        For m = 1 To NumberCVents(i, j)
                            r1ltemp = 0
                            r1utemp = 0
                            r2utemp = 0
                            r2ltemp = 0

                            Call ventareasC(tim, i, j, m, avent) 'calculates vent area, 'avent', for horizontal/ceiling/floor vents

                            If avent > gcd_Machine_Error Then 'vent is open
                                'determine the elevation of the horizontal vent
                                If j = NumberRooms + 1 Then 'lower room is an outside space
                                    YVent = FloorElevation(i) 'vent is at the floor height of the upper room, i
                                Else 'lower room is an inside space
                                    YVent = FloorElevation(j) + RoomHeight(j) 'vent is at the ceiling height of the lower room, j
                                End If
                                If j <> NumberRooms + 1 And i <> NumberRooms + 1 Then
                                    If FloorElevation(i) + gcd_Machine_Error < FloorElevation(j) + RoomHeight(j) Then 'floor of upper room, i, is lower than ceiling of lower floor, j
                                        MsgBox("The horizontal vent configuration may be invalid. Please check.")
                                        flagstop = 1
                                    End If
                                End If

                                DischargeCoeff = CVentDC(i, j, m)

                                Call VENTCF2A(avent, conl, conu, NProd, TU, TL, cpu, cpl, yflor, yceil, ylay, YVent, pflor, epsp, denl, denu, C, cvent_Renamed, XML, XMU, XMVent, PL, PU, PVent, P, QL, QU, QVent, T, Tvent, DELP1, Den, DischargeCoeff)

                                'flow taken from lower room (2) and put in upper room (1)
                                Call CV_Flogo(dteps, ylay(1), yceil(1), yflor(1), Tvent(1), TU(1), TL(1), FU)
                                r1ltemp = XMVent(1) * (1 - FU)
                                r1utemp = XMVent(1) * FU
                                r2utemp = XMU(2)
                                r2ltemp = XML(2)

                                r1l = r1l + XMVent(1) * (1 - FU)
                                r1u = r1u + XMVent(1) * FU
                                r2u = r2u + XMU(2)
                                r2l = r2l + XML(2)
                                e1l = e1l + QVent(1) * (1 - FU)
                                e1u = e1u + QVent(1) * FU
                                e2u = e2u + QU(2)
                                e2l = e2l + QL(2)
                                For h = 3 To NProd + 2
                                    prod1L(h) = prod1L(h) + PVent(h - 2, 1) * (1 - FU)
                                    prod1U(h) = prod1U(h) + PVent(h - 2, 1) * FU
                                    prod2U(h) = prod2U(h) + PU(h - 2, 2)
                                    prod2L(h) = prod2L(h) + PL(h - 2, 2)
                                Next h

                                'flow taken from upper room (1) and put in lower room (2)
                                Call CV_Flogo(0, ylay(2), yceil(2), yflor(2), Tvent(2), TU(2), TL(2), FU)
                                r2ltemp = r2ltemp + XMVent(2) * (1 - FU)
                                r2utemp = r2utemp + XMVent(2) * FU
                                r1utemp = r1utemp + XMU(1)
                                r1ltemp = r1ltemp + XML(1)

                                r2l = r2l + XMVent(2) * (1 - FU)
                                r2u = r2u + XMVent(2) * FU
                                r1u = r1u + XMU(1)
                                r1l = r1l + XML(1)
                                e2l = e2l + QVent(2) * (1 - FU)
                                e2u = e2u + QVent(2) * FU
                                e1u = e1u + QU(1)
                                e1l = e1l + QL(1)

                                For h = 3 To NProd + 2
                                    prod2L(h) = prod2L(h) + PVent(h - 2, 2) * (1 - FU)
                                    prod2U(h) = prod2U(h) + PVent(h - 2, 2) * FU
                                    prod1U(h) = prod1U(h) + PU(h - 2, 1)
                                    prod1L(h) = prod1L(h) + PL(h - 2, 1)
                                Next h

                                'If frmOptions1.chkSaveVentFlow.Value = vbChecked And (tim > tlast Or tim = 0) Then
                                If frmInputs.chkSaveVentFlow.CheckState = System.Windows.Forms.CheckState.Checked And (tim > tlast Or FirstTime2 = True) Then
                                    If tim - Int(tim) = 0 Then
                                        Call Ventflow_log3(FirstTime2, tim, i, j, m, r1utemp, r1ltemp, r2utemp, r2ltemp, XMVent)
                                    End If
                                End If

                                'keep track of flow to the fire room, for use in GER calculation
                                If i = fireroom Then
                                    If r1u > 0 Then Flowtoroom(i) = Flowtoroom(i) + r1u
                                    If r1l > 0 Then Flowtoroom(i) = Flowtoroom(i) + r1l
                                ElseIf j = fireroom Then
                                    If r2u > 0 Then Flowtoroom(j) = Flowtoroom(j) + r2u
                                    If r2l > 0 Then Flowtoroom(j) = Flowtoroom(j) + r2l
                                End If

                            End If
                            'If m = NumberVents(i, j) Then gb_first_time_vent = False 'cw 18/3/2010
                        Next m
                        ventflow(i, 1) = ventflow(i, 1) + r1l 'upper room lower layer
                        ventflow(i, 2) = ventflow(i, 2) + r1u 'upper room upper layer
                        ventflow(j, 1) = ventflow(j, 1) + r2l 'lower room lower layer
                        ventflow(j, 2) = ventflow(j, 2) + r2u 'lower room upper layer
                        Enthalpy(i, 1) = Enthalpy(i, 1) + e1l 'upper room lower layer
                        Enthalpy(i, 2) = Enthalpy(i, 2) + e1u 'upper room upper layer
                        Enthalpy(j, 1) = Enthalpy(j, 1) + e2l 'lower room lower layer
                        Enthalpy(j, 2) = Enthalpy(j, 2) + e2u 'lower room upper layer
                        For h = 3 To NProd + 2
                            products(h, i, 1) = products(h, i, 1) + prod1L(h) '1st room lower layer
                            products(h, i, 2) = products(h, i, 2) + prod1U(h) '1st room upper layer
                            products(h, j, 1) = products(h, j, 1) + prod2L(h) '2nd room lower layer
                            products(h, j, 2) = products(h, j, 2) + prod2U(h) '2nd room upper layer
                        Next h
                    End If
                    If i = NumberRooms + 1 And r2u < 0 Then 'has to be a horizontal vent
                        FlowOut(j) = FlowOut(j) + -r2u / denu(2)
                    End If

                End If
            Next j 'next room
        Next i 'next room
        tlast = tim

        gb_first_time_vent = False

        Exit Sub

errorhandler:
        MsgBox("error in subroutine VENTFLOWS - terminating simulation")
        flagstop = 1

    End Sub

    Private Sub check_cvent_areas(ByVal i As Long, ByVal j As Long)
        'simple check that the total possible ceiling vent area does not exceed the room floor area
        ' i = the upper room
        ' j = the lower room
        If j > NumberRooms Then Exit Sub

        Dim total_area As Double
        Dim k As Integer

        total_area = 0 'initialise

        For k = 1 To NumberCVents(i, j)
            total_area = total_area + CVentArea(i, j, k)
        Next k

        If total_area > RoomLength(j) * RoomWidth(j) Then
            MsgBox("Horizontal vent area larger than room floor area. Please check.")
            flagstop = 1 'back out of run
        End If

    End Sub
    Private Sub check_vent_areas_C(ByRef i As Integer, ByRef j As Integer)
        '=================
        ' Check cumulative vent areas are less than room floor areas
        '=================
        Dim vent_open() As Integer
        Dim counter2, counter1, counter3 As Integer
        Dim vent_area() As Double

        ReDim vent_open(NumberCVents(i, j))
        ReDim vent_area(NumberCVents(i, j))

        'Call check_vent_dimensions_C(i, j)

        For counter1 = 1 To NumberCVents(i, j) 'initialise cumulative vent areas
            vent_area(counter1) = 0
        Next counter1

        If NumberCVents(i, j) > 1 Then
            'sort the vent numbers in order of opening times
            Call HSort_Vents(i, j, vent_open)
            Call Remove_Duplicate_Vents(i, j, vent_open, counter1)
        Else
            vent_open(1) = 1
            counter1 = 1
        End If

        'calculate maximum cumulative vent area for all time considered
        For counter2 = 1 To NumberCVents(i, j)
            'If CVentCloseTime(i, j, counter2) = 0 And CVentCloseTime(i, j, counter2) = 0 Then 'vent is always open 'v2008.3
            If CVentOpenTime(i, j, counter2) = 0 And CVentCloseTime(i, j, counter2) = 0 Then 'vent is always open 'v2008.4
                For counter3 = 1 To counter1
                    vent_area(counter3) = vent_area(counter3) + CVentArea(i, j, counter2)
                Next counter3
            Else 'vent is not always open
                For counter3 = 1 To counter1
                    If CVentOpenTime(i, j, counter2) <= CVentOpenTime(i, j, vent_open(counter3)) + gcd_Machine_Error And CVentCloseTime(i, j, counter2) + gcd_Machine_Error >= CVentOpenTime(i, j, vent_open(counter3)) Then
                        vent_area(counter3) = vent_area(counter3) + CVentArea(i, j, counter2)
                    End If
                Next counter3
            End If
        Next counter2

        For counter2 = 1 To counter1
            If j <> NumberRooms + 1 Then 'lower space is an inside room
                If i <> NumberRooms + 1 Then 'upper space is an inside room
                    If vent_area(counter2) > RoomWidth(j) * RoomLength(j) Or vent_area(counter2) > RoomWidth(j) * RoomLength(j) Then
                        MsgBox("Horizontal vent area larger than room floor area. Please check.")
                        flagstop = 1
                    End If
                Else 'upper space is outside
                    If vent_area(counter2) > RoomWidth(j) * RoomLength(j) Then
                        MsgBox("Horizontal vent area larger than room floor area. Please check.")
                        flagstop = 1
                    End If
                End If
            Else 'lower space is space is outside
                If i <> NumberRooms + 1 Then 'upper space is an inside room
                    If vent_area(counter2) > RoomWidth(j) * RoomLength(j) Then
                        MsgBox("Horizontal vent area larger than room floor area. Please check.")
                        flagstop = 1
                    End If
                Else 'both upper and lower spaces are outside
                    MsgBox("A vent is present between two 'outside' spaces. Please check.")
                    flagstop = 1
                End If
            End If
        Next counter2

    End Sub
	Private Sub Remove_Duplicate_Vents(ByRef i As Integer, ByRef j As Integer, ByRef vent_open() As Integer, ByRef counter1 As Integer)
		'==============
		'   removes vent numbers with same opening times
		'==============
		Dim Ii, jj, flag As Integer
		Dim diff As Double
		jj = 1
		flag = 0
		counter1 = NumberCVents(i, j)
		For Ii = 2 To counter1
			diff = CVentOpenTime(i, j, vent_open(Ii)) - CVentOpenTime(i, j, vent_open(jj))
			If diff > gcd_Machine_Error Then
				jj = jj + 1
				If flag = 1 Then
					CVentOpenTime(i, j, vent_open(jj)) = CVentOpenTime(i, j, vent_open(Ii))
				End If
			Else : flag = 1
			End If
		Next 
		counter1 = jj
	End Sub
	Private Sub HSort_Vents(ByRef i As Integer, ByRef j As Integer, ByRef vent_open() As Integer)
		'=======================================================================
		' This routine sorts vents opening in ascending order using a heap sort method.
		' vent_open() is an array of the number of each vent as it opens
		'======================================================================
        Dim Ii, RA, L, IR, jj As Integer
		
		L = NumberCVents(i, j) '- 1) \ 2
		IR = NumberCVents(i, j)
		Do While L > 0
			Ii = L
			jj = 1
			Do While jj < L
				If CVentOpenTime(i, j, vent_open(1)) > CVentOpenTime(i, j, vent_open(jj)) Then
					RA = vent_open(Ii)
					vent_open(Ii) = vent_open(jj)
					vent_open(jj) = RA
				End If
				If CVentOpenTime(i, j, vent_open(Ii)) < CVentOpenTime(i, j, vent_open(jj)) Then
					RA = vent_open(Ii)
					vent_open(Ii) = vent_open(jj)
					vent_open(jj) = RA
				End If
				If jj < IR Then If CVentOpenTime(i, j, vent_open(jj)) < CVentOpenTime(i, j, vent_open(jj + 1)) Then jj = jj + 1
			Loop 
			L = L - 1
		Loop 
		IR = NumberCVents(i, j)
		Do While IR > 1
			IR = IR - 1
			Ii = 1
			Do While Ii * 2 < IR
				jj = Ii * 2
				If jj < IR Then
					If CVentOpenTime(i, j, vent_open(jj)) < CVentOpenTime(i, j, vent_open(jj + 1)) Then jj = jj + 1
				End If
				If CVentOpenTime(i, j, vent_open(Ii)) > CVentOpenTime(i, j, vent_open(jj)) Then
					RA = vent_open(Ii)
					vent_open(Ii) = vent_open(jj)
					vent_open(jj) = vent_open(Ii)
					Ii = jj
				Else
					Exit Do
				End If
			Loop 
		Loop 
	End Sub
	
    Private Sub Outside_Conditions(ByRef conl(,) As Double, ByRef conu(,) As Double, ByRef TU() As Double, ByRef TL() As Double, ByRef cpu() As Double, ByRef cpl() As Double, ByRef pflor() As Double, ByRef denu() As Double, ByRef denl() As Double, ByRef room_number As Integer)
        '=================
        'Sets outside conditions for VENTCF2A
        '================
        conl(1, room_number) = O2Ambient 'lower oxygen
        conu(1, room_number) = O2Ambient 'upper oxygen
        conl(2, room_number) = COMassFraction(1, 1, 2) 'lower co
        conu(2, room_number) = COMassFraction(1, 1, 1) 'upper co
        conl(3, room_number) = CO2Ambient 'lower co2
        conu(3, room_number) = CO2Ambient 'upper co2
        conl(4, room_number) = SootMassFraction(1, 1, 2) 'lower soot
        conu(4, room_number) = SootMassFraction(1, 1, 1) 'upper soot
        conl(5, room_number) = H2OMassFraction(1, 1, 2) 'lower water
        conu(5, room_number) = H2OMassFraction(1, 1, 1) 'upper water
        conl(6, room_number) = TUHC(1, 1, 2) 'lower tuhc
        conu(6, room_number) = TUHC(1, 1, 1) 'upper tuhc
        conl(7, room_number) = HCNMassFraction(1, 1, 2) 'lower hcn
        conu(7, room_number) = HCNMassFraction(1, 1, 1) 'upper hcn
        TU(room_number) = ExteriorTemp
        TL(room_number) = ExteriorTemp
        cpu(room_number) = SpecificHeat_air
        cpl(room_number) = SpecificHeat_air
        denu(room_number) = (Atm_Pressure + pflor(room_number) / 1000) / (Gas_Constant / MW_air) / ExteriorTemp
        denl(room_number) = denu(room_number)
    End Sub
    Public Sub ventareasC(ByRef tim As Double, ByRef i As Integer, ByRef j As Integer, ByRef m As Integer, ByRef avent As Double)
        '==========================
        'calculates vent area for ceiling/floor vents
        'CVentAuto(i, j, m) = True, means ceiling vent opens automatically
        '==========================

        'If CVentAuto(i, j, m) = True And AutoFlag(j) = False Then 'we want to automatically open the ceiling vent

        '    'if we want to auto open the vent, but the CVentOpenTime is initially zero, this will make sure that vent is treated as closed
        '    'at the start of the simulation
        '    'If AutoFlag(j) = False And SDFlag(j) = 0 Then CVentOpenTime(i, j, m) = SimTime
        '    ' If SDFlag(j) = 0 Then CVentOpenTime(i, j, m) = SimTime
        '    If tim = 0 Then CVentOpenTime(i, j, m) = SimTime

        '    'open ceiling vent if the sprinkler/thermal detector activates, and the vent is in fireroom
        '    If (j = fireroom And AutoFlag(j) = False) Then
        '        If trigger_device2(1, i, j, m) = True And breakflag2(i, j, m) = False Then
        '            If SprinklerFlag = 1 Then
        '                CVentOpenTime(i, j, m) = SprinklerTime
        '                'also open any other ceiling vents connecting the same two rooms
        '                Dim Message As String = Format(SprinklerTime, "0") & " sec. Ceiling vent(s) in room " & j.ToString & " opens."

        '                If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

        '                For mm = 2 To NumberCVents(i, j)
        '                    CVentOpenTime(i, j, mm) = SprinklerTime
        '                    Message = CStr(SprinklerTime) & " sec. Ceiling vent " & i.ToString & j.ToString & mm.ToString & " opens."
        '                    If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

        '                Next
        '                AutoFlag(j) = True
        '            End If

        '            If HDFlag = 1 Then
        '                CVentOpenTime(i, j, m) = HDTime
        '                'also open any other ceiling vents connecting the same two rooms
        '                Dim Message As String = Format(HDTime, "0") & " sec. Ceiling vent(s) in room " & j.ToString & " opens."

        '                If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

        '                For mm = 2 To NumberCVents(i, j)
        '                    CVentOpenTime(i, j, mm) = HDTime
        '                    Message = CStr(HDTime) & " sec. Ceiling vent " & i.ToString & j.ToString & mm.ToString & " opens."
        '                    If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

        '                Next
        '                AutoFlag(j) = True
        '            End If
        '        End If
        '    End If

        '    'open ceiling vent if the smoke detector in same room activates
        '    If trigger_device2(0, i, j, m) = True And SDFlag(SDtriggerroom2(i, j, m)) = 1 And breakflag2(i, j, m) = False Then
        '        If (AutoFlag(j) = False) Then
        '            CVentOpenTime(i, j, m) = SDTime(j)
        '            Dim Message As String = Format(SDTime(j), "0") & " sec. Ceiling vent(s) in room " & j.ToString & " opens."
        '            If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

        '            'also open any other ceiling vents connecting the same two rooms
        '            For mm = 2 To NumberCVents(i, j)
        '                CVentOpenTime(i, j, mm) = SDTime(j)
        '                'Message = CStr(SDTime(j)) & " sec. Ceiling vent " & i.ToString & j.ToString & mm.ToString & " opens."
        '                'If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
        '            Next
        '            AutoFlag(j) = True

        '        End If

        '    End If
        'End If

        'If trigger_device2(6, i, j, m) = True Then 'if we are opening vent based on integrity failure
        '    avent = VentHeight(i, j, m) * VentWidth(i, j, m) * FRMaxOpening(i, j, m) / 100
        '    'open the vent over specified period
        '    If FRMaxOpeningTime(i, j, m) > 1 Then
        '        If (ttim - VentOpenTime(i, j, m)) / (FRMaxOpeningTime(i, j, m)) < 1 Then
        '            avent = avent * (ttim - VentOpenTime(i, j, m)) / (FRMaxOpeningTime(i, j, m))
        '        End If
        '    Else

        '        avent = avent * (ttim - VentOpenTime(i, j, m)) / 2

        '    End If
        '    'If ttim = 1920 Then Stop
        '    ventarea = avent
        '    Exit Sub
        '    'can't use both fire resistance and other autoventopen options for now 

        'Else

        'End If

        'Check if vent is open
        If tim >= CVentOpenTime(i, j, m) Then
            If (tim - CVentOpenTime(i, j, m) < 2) And CVentOpenTime(i, j, m) > 0 Then 'vent is partially open
                avent = CVentArea(i, j, m) * (tim - CVentOpenTime(i, j, m)) / 2
            Else 'vent is fully open
                avent = CVentArea(i, j, m)
            End If
        Else 'vent has not opened yet
            avent = 0
        End If

        'Check if vent is closed
        If tim > CVentCloseTime(i, j, m) And (CVentCloseTime(i, j, m) > CVentOpenTime(i, j, m)) Then
            If tim - CVentCloseTime(i, j, m) < 2 And CVentCloseTime(i, j, m) > CVentOpenTime(i, j, m) + 2 Then 'vent is closing
                avent = CVentArea(i, j, m) * (2 - (tim - CVentCloseTime(i, j, m))) / 2
            Else 'vent is fully closed
                avent = 0
            End If
        End If

    End Sub
	
	Public Function ventarea(ByRef tim As Double, ByRef i As Integer, ByRef j As Integer, ByRef m As Integer) As Double
		'=====================================================
		'   get the open area of a vent at any time
		' 28 Jan 2008
		''tim' and vent opening and closing times are a mixture of double and single precision. Seems ok so far. Keep an eye on it.
		'=====================================================
        Try


            Dim avent As Double
            Dim ttim As Single = Convert.ToSingle(tim) 'doesn't seem to make any difference

            avent = 0

            If TalkToEVACNZ = True Then
                If enzbreakflag(i, j, m) = True Then
                    avent = VentHeight(i, j, m) * VentWidth(i, j, m)
                End If
                ventarea = avent
                Exit Function
            End If

            If (ttim >= VentCloseTime(i, j, m)) And (VentCloseTime(i, j, m) = VentOpenTime(i, j, m)) And VentOpenTime(i, j, m) > 0 Then
                avent = 0
                ventarea = avent
                Exit Function
            End If

            If ttim >= VentOpenTime(i, j, m) Then

                'if we are opening vent based on integrity failure
                If trigger_device(6, i, j, m) = True Then
                    avent = VentHeight(i, j, m) * VentWidth(i, j, m) * FRMaxOpening(i, j, m) / 100

                    'open the vent over specified period
                    If FRMaxOpeningTime(i, j, m) > 1 Then
                        If (ttim - VentOpenTime(i, j, m)) / (FRMaxOpeningTime(i, j, m)) < 1 Then
                            avent = avent * (ttim - VentOpenTime(i, j, m)) / (FRMaxOpeningTime(i, j, m))
                        End If
                    Else

                        avent = avent * (ttim - VentOpenTime(i, j, m)) / 2

                    End If

                    ventarea = avent
                    Exit Function
                    'can't use both fire resistance and other autoventopen options for now 

                Else
                    avent = VentHeight(i, j, m) * VentWidth(i, j, m)
                End If

                If AutoBreakGlass(i, j, m) = True Then
                    If (ttim - VentOpenTime(i, j, m) < 2 + GLASSFalloutTime(i, j, m)) And (VentOpenTime(i, j, m) > 0) Then
                        'open the vent over a 2 second period
                        avent = avent * (ttim - VentOpenTime(i, j, m)) / (2 + GLASSFalloutTime(i, j, m))
                    End If
                Else
                    If (ttim - VentOpenTime(i, j, m) < 2) And VentOpenTime(i, j, m) > 0 Then
                        'open the vent over a 2 second period
                        avent = avent * (ttim - VentOpenTime(i, j, m)) / 2
                    End If
                    If (ttim > VentCloseTime(i, j, m)) And (VentCloseTime(i, j, m) > VentOpenTime(i, j, m)) Then
                        If (ttim - VentCloseTime(i, j, m) < 2) And (VentCloseTime(i, j, m) > VentOpenTime(i, j, m) + 2) Then
                            avent = avent * (2 - (ttim - VentCloseTime(i, j, m))) / 2
                        Else
                            avent = 0
                        End If
                    End If

                End If
            End If

            ventarea = avent

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in stiffsolv.vb ventarea")
        End Try
	End Function
	
    Public Function species_mass_rate_postflashover(ByVal GER As Single, ByVal room As Integer, ByRef Yield As Single, ByRef WallYield() As Double, ByRef CeilingYield() As Double, ByRef FloorYield() As Double, ByRef T As Double, ByRef mrate() As Double, ByRef mrate_wall As Double, ByRef mrate_ceiling As Double, ByRef mrate_floor As Double) As Double
        '*  ====================================================================
        '*  This function returns the value of the sum of the fuel mass loss rates
        '*  multiplied by the species generation rate for all burning objects.
        '*  ====================================================================

        'only use for CO2 and H2O species, these yields reduce in underventilated fires

        Dim dummy As Double
        Dim total As Single

        If room = fireroom Then
            If GER > 1 Then 'ventilation restricted
                'For i = 1 To NumberObjects
                dummy = mrate(1) * Yield / GER
                If useCLTmodel = True Then dummy = (mrate(1) - mrate_ceiling - mrate_wall - mrate_floor) * Yield / GER 'mrate includes linings
                'If i = 1 Then
                dummy = dummy + mrate_wall * WallYield(room) / GER + mrate_ceiling * CeilingYield(room) / GER + mrate_floor * FloorYield(room) / GER
                'End If
                total = total + dummy
                'Next i
            Else
                'For i = 1 To NumberObjects
                dummy = mrate(1) * Yield
                If useCLTmodel = True Then dummy = (mrate(1) - mrate_ceiling - mrate_wall - mrate_floor) * Yield  'mrate includes linings
                'If i = 1 Then
                dummy = dummy + mrate_wall * WallYield(room) + mrate_ceiling * CeilingYield(room) + mrate_floor * FloorYield(room)
                    'End If
                    total = total + dummy
                    'Next i
                End If
                Else
            total = mrate_ceiling * CeilingYield(room)
        End If

        species_mass_rate_postflashover = total 'kg-species per sec

    End Function
	
    Public Function species_mass_rate2(ByVal GER As Single, ByVal room As Integer, ByRef Yield() As Single, ByRef WallYield() As Double, ByRef CeilingYield() As Double, ByRef FloorYield() As Double, ByRef T As Double, ByRef mrate() As Double, ByRef mrate_wall As Double, ByRef mrate_ceiling As Double, ByRef mrate_floor As Double) As Double
        '*  ====================================================================
        '*  This function returns the value of the sum of the fuel mass loss rates
        '*  multiplied by the species generation rate for all burning objects.
        '*  ====================================================================

        Dim dummy As Double
        Dim total As Double
        Dim i As Integer

        If room = fireroom Then
            For i = 1 To NumberObjects
                dummy = mrate(i) * Yield(i) / GER

                If i = 1 Then
                    If useCLTmodel = True And KineticModel = False Then dummy = (mrate(i) - mrate_wall - mrate_ceiling - mrate_floor) * Yield(i) / GER
                    dummy = dummy + mrate_wall * WallYield(room) / GER + mrate_ceiling * CeilingYield(room) / GER + mrate_floor * FloorYield(room) / GER
                End If
                total = total + dummy
            Next i
        Else
            total = mrate_ceiling * CeilingYield(room)
        End If
        species_mass_rate2 = total 'kg-species per sec

    End Function
	
    Public Function specific_heat_upper(ByRef j_dummy As Integer, ByRef Y(,) As Double, ByRef temp As Double) As Double
        '==============================================
        ' look up the specific heat capacity for a given combustion product
        ' Based on Public Function specific_heat() modified specific_heat() to calculate specific heat of the upper layer of a room in one hit (to reduce the number of calls to the function)
        ' j_dummy = room number
        ' Y() = integration vector(room number, dx/dy variable)
        ' temp = temperature
        ' 1 feb 2008 Amanda
        '================================================
        If temp < 300 Then
            specific_heat_upper = 1.0403 * Y(j_dummy, 8) + 0.85262 * Y(j_dummy, 10) + 1.86 * Y(j_dummy, 16) + 0.9199 * Y(j_dummy, 5) + 1.0414 * (1 - Y(j_dummy, 5) - Y(j_dummy, 16) - Y(j_dummy, 8) - Y(j_dummy, 10)) : Exit Function
            'specific_heat_upper = 1.0403 * Y(j_dummy, 8) + 0.85262 * Y(j_dummy, 10) + 4.1806 * Y(j_dummy, 16) + 0.9199 * Y(j_dummy, 5) + 1.0414 * (1 - Y(j_dummy, 5) - Y(j_dummy, 16) - Y(j_dummy, 8) - Y(j_dummy, 10)) : Exit Function
        ElseIf temp < 350 Then
            specific_heat_upper = 1.0403 * Y(j_dummy, 8) + 0.85262 * Y(j_dummy, 10) + 1.88 * Y(j_dummy, 16) + 0.9199 * Y(j_dummy, 5) + 1.0414 * (1 - Y(j_dummy, 5) - Y(j_dummy, 16) - Y(j_dummy, 8) - Y(j_dummy, 10)) : Exit Function
            'specific_heat_upper = 1.0403 * Y(j_dummy, 8) + 0.85262 * Y(j_dummy, 10) + 4.2156 * Y(j_dummy, 16) + 0.9199 * Y(j_dummy, 5) + 1.0414 * (1 - Y(j_dummy, 5) - Y(j_dummy, 16) - Y(j_dummy, 8) - Y(j_dummy, 10)) : Exit Function
        ElseIf temp < 373 Then
            specific_heat_upper = 1.0393 * Y(j_dummy, 8) + 0.94177 * Y(j_dummy, 10) + 1.89 * Y(j_dummy, 16) + 0.9417 * Y(j_dummy, 5) + 1.045 * (1 - Y(j_dummy, 5) - Y(j_dummy, 16) - Y(j_dummy, 8) - Y(j_dummy, 10)) : Exit Function
            'specific_heat_upper = 1.0393 * Y(j_dummy, 8) + 0.94177 * Y(j_dummy, 10) + 2.0093 * Y(j_dummy, 16) + 0.9417 * Y(j_dummy, 5) + 1.045 * (1 - Y(j_dummy, 5) - Y(j_dummy, 16) - Y(j_dummy, 8) - Y(j_dummy, 10)) : Exit Function
        ElseIf temp < 450 Then
            specific_heat_upper = 1.0396 * Y(j_dummy, 8) + 1.0155 * Y(j_dummy, 10) + 2.0093 * Y(j_dummy, 16) + 0.9722 * Y(j_dummy, 5) + 1.0564 * (1 - Y(j_dummy, 5) - Y(j_dummy, 16) - Y(j_dummy, 8) - Y(j_dummy, 10)) : Exit Function
        ElseIf temp < 550 Then
            specific_heat_upper = 1.0396 * Y(j_dummy, 8) + 1.0155 * Y(j_dummy, 10) + 1.9816 * Y(j_dummy, 16) + 0.9722 * Y(j_dummy, 5) + 1.0564 * (1 - Y(j_dummy, 5) - Y(j_dummy, 16) - Y(j_dummy, 8) - Y(j_dummy, 10)) : Exit Function
        ElseIf temp < 650 Then
            specific_heat_upper = 1.0409 * Y(j_dummy, 8) + 1.0762 * Y(j_dummy, 10) + 2.0269 * Y(j_dummy, 16) + 1.0033 * Y(j_dummy, 5) + 1.0751 * (1 - Y(j_dummy, 5) - Y(j_dummy, 16) - Y(j_dummy, 8) - Y(j_dummy, 10)) : Exit Function
        ElseIf temp < 750 Then
            specific_heat_upper = 1.043 * Y(j_dummy, 8) + 1.1269 * Y(j_dummy, 10) + 2.0868 * Y(j_dummy, 16) + 1.031 * Y(j_dummy, 5) + 1.0981 * (1 - Y(j_dummy, 5) - Y(j_dummy, 16) - Y(j_dummy, 8) - Y(j_dummy, 10)) : Exit Function
        ElseIf temp < 850 Then
            specific_heat_upper = 1.0457 * Y(j_dummy, 8) + 1.1692 * Y(j_dummy, 10) + 2.1526 * Y(j_dummy, 16) + 1.0545 * Y(j_dummy, 5) + 1.1223 * (1 - Y(j_dummy, 5) - Y(j_dummy, 16) - Y(j_dummy, 8) - Y(j_dummy, 10)) : Exit Function
        ElseIf temp < 950 Then
            specific_heat_upper = 1.0488 * Y(j_dummy, 8) + 1.2046 * Y(j_dummy, 10) + 2.2217 * Y(j_dummy, 16) + 1.0739 * Y(j_dummy, 5) + 1.1457 * (1 - Y(j_dummy, 5) - Y(j_dummy, 16) - Y(j_dummy, 8) - Y(j_dummy, 10)) : Exit Function
        ElseIf temp < 1050 Then
            specific_heat_upper = 1.052 * Y(j_dummy, 8) + 1.2343 * Y(j_dummy, 10) + 2.2921 * Y(j_dummy, 16) + 1.0898 * Y(j_dummy, 5) + 1.1674 * (1 - Y(j_dummy, 5) - Y(j_dummy, 16) - Y(j_dummy, 8) - Y(j_dummy, 10)) : Exit Function
        ElseIf temp < 1150 Then
            specific_heat_upper = 1.052 * Y(j_dummy, 8) + 1.2593 * Y(j_dummy, 10) + 2.3621 * Y(j_dummy, 16) + 1.0898 * Y(j_dummy, 5) + 1.1868 * (1 - Y(j_dummy, 5) - Y(j_dummy, 16) - Y(j_dummy, 8) - Y(j_dummy, 10)) : Exit Function
        ElseIf temp < 1250 Then
            specific_heat_upper = 1.052 * Y(j_dummy, 8) + 1.2593 * Y(j_dummy, 10) + 2.4303 * Y(j_dummy, 16) + 1.0898 * Y(j_dummy, 5) + 1.204 * (1 - Y(j_dummy, 5) - Y(j_dummy, 16) - Y(j_dummy, 8) - Y(j_dummy, 10)) : Exit Function
        ElseIf temp < 1350 Then
            specific_heat_upper = 1.052 * Y(j_dummy, 8) + 1.2593 * Y(j_dummy, 10) + 2.4303 * Y(j_dummy, 16) + 1.0898 * Y(j_dummy, 5) + 1.2191 * (1 - Y(j_dummy, 5) - Y(j_dummy, 16) - Y(j_dummy, 8) - Y(j_dummy, 10)) : Exit Function
        ElseIf temp < 1450 Then
            specific_heat_upper = 1.052 * Y(j_dummy, 8) + 1.2593 * Y(j_dummy, 10) + 2.4303 * Y(j_dummy, 16) + 1.0898 * Y(j_dummy, 5) + 1.2323 * (1 - Y(j_dummy, 5) - Y(j_dummy, 16) - Y(j_dummy, 8) - Y(j_dummy, 10)) : Exit Function
        Else
            specific_heat_upper = 1.052 * Y(j_dummy, 8) + 1.2593 * Y(j_dummy, 10) + 2.4303 * Y(j_dummy, 16) + 1.0898 * Y(j_dummy, 5) + 1.2439 * (1 - Y(j_dummy, 5) - Y(j_dummy, 16) - Y(j_dummy, 8) - Y(j_dummy, 10)) : Exit Function
        End If
        MsgBox("Specific Heat Capacity not found - error in STIFFSOLV")
    End Function
    Public Function specific_heat_lower(ByRef j_dummy As Integer, ByRef Y(,) As Double, ByRef temp As Double) As Double
        '==============================================
        ' look up the specific heat capacity for a given combustion product
        ' Based on Public Function specific_heat() modified specific_heat() to calculate specific heat of the lower layer of a room in one hit (to reduce the number of calls to the function)
        ' j_dummy = room number
        ' Y() = integration vector(room number, dx/dy variable)
        ' temp = temperature
        ' 1 Feb 08 Amanda
        '================================================
        If temp < 300 Then
            specific_heat_lower = 1.0403 * Y(j_dummy, 9) + 0.85262 * Y(j_dummy, 11) + 1.86 * Y(j_dummy, 17) + 0.9199 * Y(j_dummy, 15) + 1.0414 * (1 - Y(j_dummy, 15) - Y(j_dummy, 17) - Y(j_dummy, 9) - Y(j_dummy, 11)) : Exit Function
            'specific_heat_lower = 1.0403 * Y(j_dummy, 9) + 0.85262 * Y(j_dummy, 11) + 4.1806 * Y(j_dummy, 17) + 0.9199 * Y(j_dummy, 15) + 1.0414 * (1 - Y(j_dummy, 15) - Y(j_dummy, 17) - Y(j_dummy, 9) - Y(j_dummy, 11)) : Exit Function
        ElseIf temp < 350 Then
            specific_heat_lower = 1.0403 * Y(j_dummy, 9) + 0.85262 * Y(j_dummy, 11) + 1.88 * Y(j_dummy, 17) + 0.9199 * Y(j_dummy, 15) + 1.0414 * (1 - Y(j_dummy, 15) - Y(j_dummy, 17) - Y(j_dummy, 9) - Y(j_dummy, 11)) : Exit Function
            'specific_heat_lower = 1.0403 * Y(j_dummy, 9) + 0.85262 * Y(j_dummy, 11) + 4.2156 * Y(j_dummy, 17) + 0.9199 * Y(j_dummy, 15) + 1.0414 * (1 - Y(j_dummy, 15) - Y(j_dummy, 17) - Y(j_dummy, 9) - Y(j_dummy, 11)) : Exit Function
        ElseIf temp < 373 Then
            specific_heat_lower = 1.0393 * Y(j_dummy, 9) + 0.94177 * Y(j_dummy, 11) + 1.89 * Y(j_dummy, 17) + 0.9417 * Y(j_dummy, 15) + 1.045 * (1 - Y(j_dummy, 15) - Y(j_dummy, 17) - Y(j_dummy, 9) - Y(j_dummy, 11)) : Exit Function
            'specific_heat_lower = 1.0393 * Y(j_dummy, 9) + 0.94177 * Y(j_dummy, 11) + 2.0093 * Y(j_dummy, 17) + 0.9417 * Y(j_dummy, 15) + 1.045 * (1 - Y(j_dummy, 15) - Y(j_dummy, 17) - Y(j_dummy, 9) - Y(j_dummy, 11)) : Exit Function
        ElseIf temp < 450 Then
            specific_heat_lower = 1.0396 * Y(j_dummy, 9) + 1.0155 * Y(j_dummy, 11) + 2.0093 * Y(j_dummy, 17) + 0.9722 * Y(j_dummy, 15) + 1.0564 * (1 - Y(j_dummy, 15) - Y(j_dummy, 17) - Y(j_dummy, 9) - Y(j_dummy, 11)) : Exit Function
        ElseIf temp < 550 Then
            specific_heat_lower = 1.0396 * Y(j_dummy, 9) + 1.0155 * Y(j_dummy, 11) + 1.9816 * Y(j_dummy, 17) + 0.9722 * Y(j_dummy, 15) + 1.0564 * (1 - Y(j_dummy, 15) - Y(j_dummy, 17) - Y(j_dummy, 9) - Y(j_dummy, 11)) : Exit Function
        ElseIf temp < 650 Then
            specific_heat_lower = 1.0409 * Y(j_dummy, 9) + 1.0762 * Y(j_dummy, 11) + 2.0269 * Y(j_dummy, 17) + 1.0033 * Y(j_dummy, 15) + 1.0751 * (1 - Y(j_dummy, 15) - Y(j_dummy, 17) - Y(j_dummy, 9) - Y(j_dummy, 11)) : Exit Function
        ElseIf temp < 750 Then
            specific_heat_lower = 1.043 * Y(j_dummy, 9) + 1.1269 * Y(j_dummy, 11) + 2.0868 * Y(j_dummy, 17) + 1.031 * Y(j_dummy, 15) + 1.0981 * (1 - Y(j_dummy, 15) - Y(j_dummy, 17) - Y(j_dummy, 9) - Y(j_dummy, 11)) : Exit Function
        ElseIf temp < 850 Then
            specific_heat_lower = 1.0457 * Y(j_dummy, 9) + 1.1692 * Y(j_dummy, 11) + 2.1526 * Y(j_dummy, 17) + 1.0545 * Y(j_dummy, 15) + 1.1223 * (1 - Y(j_dummy, 15) - Y(j_dummy, 17) - Y(j_dummy, 9) - Y(j_dummy, 11)) : Exit Function
        ElseIf temp < 950 Then
            specific_heat_lower = 1.0488 * Y(j_dummy, 9) + 1.2046 * Y(j_dummy, 11) + 2.2217 * Y(j_dummy, 17) + 1.0739 * Y(j_dummy, 15) + 1.1457 * (1 - Y(j_dummy, 15) - Y(j_dummy, 17) - Y(j_dummy, 9) - Y(j_dummy, 11)) : Exit Function
        ElseIf temp < 1050 Then
            specific_heat_lower = 1.052 * Y(j_dummy, 9) + 1.2343 * Y(j_dummy, 11) + 2.2921 * Y(j_dummy, 17) + 1.0898 * Y(j_dummy, 15) + 1.1674 * (1 - Y(j_dummy, 15) - Y(j_dummy, 17) - Y(j_dummy, 9) - Y(j_dummy, 11)) : Exit Function
        ElseIf temp < 1150 Then
            specific_heat_lower = 1.052 * Y(j_dummy, 9) + 1.2593 * Y(j_dummy, 11) + 2.3621 * Y(j_dummy, 17) + 1.0898 * Y(j_dummy, 15) + 1.1868 * (1 - Y(j_dummy, 15) - Y(j_dummy, 17) - Y(j_dummy, 9) - Y(j_dummy, 11)) : Exit Function
        ElseIf temp < 1250 Then
            specific_heat_lower = 1.052 * Y(j_dummy, 9) + 1.2593 * Y(j_dummy, 11) + 2.4303 * Y(j_dummy, 17) + 1.0898 * Y(j_dummy, 15) + 1.204 * (1 - Y(j_dummy, 15) - Y(j_dummy, 17) - Y(j_dummy, 9) - Y(j_dummy, 11)) : Exit Function
        ElseIf temp < 1350 Then
            specific_heat_lower = 1.052 * Y(j_dummy, 9) + 1.2593 * Y(j_dummy, 11) + 2.4303 * Y(j_dummy, 17) + 1.0898 * Y(j_dummy, 15) + 1.2191 * (1 - Y(j_dummy, 15) - Y(j_dummy, 17) - Y(j_dummy, 9) - Y(j_dummy, 11)) : Exit Function
        ElseIf temp < 1450 Then
            specific_heat_lower = 1.052 * Y(j_dummy, 9) + 1.2593 * Y(j_dummy, 11) + 2.4303 * Y(j_dummy, 17) + 1.0898 * Y(j_dummy, 15) + 1.2323 * (1 - Y(j_dummy, 15) - Y(j_dummy, 17) - Y(j_dummy, 9) - Y(j_dummy, 11)) : Exit Function
        Else
            specific_heat_lower = 1.052 * Y(j_dummy, 9) + 1.2593 * Y(j_dummy, 11) + 2.4303 * Y(j_dummy, 17) + 1.0898 * Y(j_dummy, 15) + 1.2439 * (1 - Y(j_dummy, 15) - Y(j_dummy, 17) - Y(j_dummy, 9) - Y(j_dummy, 11)) : Exit Function
        End If
        MsgBox("Specific Heat Capacity not found - error in STIFFSOLV")
    End Function
	

    Public Sub mech_vent(ByRef X As Double, ByRef j As Integer, ByRef layerz As Double, ByRef Y(,) As Double, ByRef denl As Double, ByRef denu As Double, ByRef species_to_upper() As Double, ByRef species_to_lower() As Double, ByRef Enthalpy_U As Double, ByRef Enthalpy_L As Double, ByRef cp_u() As Double, ByRef cp_l() As Double, ByRef massplume() As Double, ByRef M_wall_u As Double, ByRef M_wall_l As Double, ByRef fuelmassloss() As Double, ByRef Flow_to_Upper As Double, ByRef Flow_to_Lower As Double, ByRef dMudt As Double, ByRef dMldt As Double, ByRef M_wall As Double, ByRef frac_to_lower As Double)

        Dim Mcrit, NewExtractRate, CrossFanPreDiff, fanflowout, FU As Double

        If X > ExtractStartTime(j) Then
            If Extract(j) = True Then
                If UseFanCurve(j) = True Then
                    'work out the cross-fan pressure difference for a fan at ceiling level
                    If FanElevation(j) > layerz Then
                        'fan in upper layer
                        'pressure just outside less pressure just inside fan location
                        CrossFanPreDiff = (0 - (Atm_Pressure) / (Gas_Constant / MW_air) / ExteriorTemp * G * (FanElevation(j) + FloorElevation(j))) - (Y(j, 4) - denl * G * layerz - denu * G * (FanElevation(j) - layerz))
                    Else
                        'fan in lower layer
                        CrossFanPreDiff = (0 - (Atm_Pressure) / (Gas_Constant / MW_air) / ExteriorTemp * G * (FanElevation(j) + FloorElevation(j))) - (Y(j, 4) - denl * G * (FanElevation(j)))
                    End If
                    If CrossFanPreDiff > MaxPressure(j) Then
                        NewExtractRate = -ExtractRate(j) * ((CrossFanPreDiff - MaxPressure(j)) / MaxPressure(j)) ^ (1 / 6)
                        fanflowout = -NewExtractRate * Atm_Pressure / Gas_Constant_Air / ExteriorTemp
                        GoTo 200
                    Else
                        NewExtractRate = ExtractRate(j) * ((MaxPressure(j) - CrossFanPreDiff) / MaxPressure(j)) ^ (1 / 6)
                    End If
                Else
                    NewExtractRate = ExtractRate(j)
                End If

                If CSng(Y(j, 1)) > CSng(0.001 * RoomVolume(j)) And FanElevation(j) >= layerz Then
                    'extracting from the upper layer
                    fanflowout = NewExtractRate * denu 'kg/s
                    If (X - ExtractStartTime(j)) < 30 Then
                        fanflowout = (X - ExtractStartTime(j)) / 30 * fanflowout 'kg/sec
                    End If

                    'calculate max flow to avoid plugholing
                    If (1 - (denu / denl)) > 0 Then
                        Mcrit = denu * 1.6 * (RoomHeight(j) - layerz) ^ 2.5 * Sqrt(G * (1 - (denu / denl)))
                    Else
                        Mcrit = 0
                    End If
                    If fanflowout > Mcrit Then
                        'plugholing
                        FU = Mcrit / fanflowout
                        'Debug.Print X, FU, Mcrit, fanflowout
                        'FU = 1
                    Else
                        'no plugholing
                        FU = 1

                    End If

                    species_to_upper(3) = species_to_upper(3) - NumberFans(j) * FU * fanflowout * Y(j, 5) 'oxygen
                    species_to_upper(4) = species_to_upper(4) - NumberFans(j) * FU * fanflowout * Y(j, 8) 'co
                    species_to_upper(5) = species_to_upper(5) - NumberFans(j) * FU * fanflowout * Y(j, 10) 'co2
                    species_to_upper(6) = species_to_upper(6) - NumberFans(j) * FU * fanflowout * Y(j, 12) 'soot
                    species_to_upper(7) = species_to_upper(7) - NumberFans(j) * FU * fanflowout * Y(j, 16) 'h20
                    species_to_upper(8) = species_to_upper(8) - NumberFans(j) * FU * fanflowout * Y(j, 6) 'tuhc
                    species_to_upper(9) = species_to_upper(9) - NumberFans(j) * FU * fanflowout * Y(j, 18) 'hcn
                    Enthalpy_U = Enthalpy_U - NumberFans(j) * FU * fanflowout * cp_u(j) * Y(j, 2)

                    species_to_lower(3) = species_to_lower(3) - NumberFans(j) * (1 - FU) * fanflowout * Y(j, 15) 'oxygen
                    species_to_lower(4) = species_to_lower(4) - NumberFans(j) * (1 - FU) * fanflowout * Y(j, 9) 'co
                    species_to_lower(5) = species_to_lower(5) - NumberFans(j) * (1 - FU) * fanflowout * Y(j, 11) 'co2
                    species_to_lower(6) = species_to_lower(6) - NumberFans(j) * (1 - FU) * fanflowout * Y(j, 13) 'soot
                    species_to_lower(7) = species_to_lower(7) - NumberFans(j) * (1 - FU) * fanflowout * Y(j, 17) 'h20
                    species_to_lower(8) = species_to_lower(8) - NumberFans(j) * (1 - FU) * fanflowout * Y(j, 7) 'tuhc
                    species_to_lower(9) = species_to_lower(9) - NumberFans(j) * (1 - FU) * fanflowout * Y(j, 14) 'hcn
                    Enthalpy_L = Enthalpy_L - NumberFans(j) * (1 - FU) * fanflowout * cp_l(j) * Y(j, 3)

                    dMudt = massplume(j) - M_wall_u + (1 - frac_to_lower) * M_wall + fuelmassloss(j) + Flow_to_Upper - NumberFans(j) * FU * fanflowout
                    dMldt = Flow_to_Lower - M_wall_l + frac_to_lower * M_wall - massplume(j) - NumberFans(j) * (1 - FU) * fanflowout
                Else
                    'extracting from the lower layer
                    fanflowout = NewExtractRate * denl 'kg/s
                    If (X - ExtractStartTime(j)) < 30 Then
                        fanflowout = (X - ExtractStartTime(j)) / 30 * fanflowout
                    End If
                    species_to_lower(3) = species_to_lower(3) - NumberFans(j) * fanflowout * Y(j, 15) 'oxygen
                    species_to_lower(4) = species_to_lower(4) - NumberFans(j) * fanflowout * Y(j, 9) 'co
                    species_to_lower(5) = species_to_lower(5) - NumberFans(j) * fanflowout * Y(j, 11) 'co2
                    species_to_lower(6) = species_to_lower(6) - NumberFans(j) * fanflowout * Y(j, 13) 'soot
                    species_to_lower(7) = species_to_lower(7) - NumberFans(j) * fanflowout * Y(j, 17) 'h20
                    species_to_lower(8) = species_to_lower(8) - NumberFans(j) * fanflowout * Y(j, 7) 'tuhc
                    Enthalpy_L = Enthalpy_L - NumberFans(j) * fanflowout * cp_l(j) * Y(j, 3)

                    dMudt = massplume(j) - M_wall_u + (1 - frac_to_lower) * M_wall + fuelmassloss(j) + Flow_to_Upper
                    dMldt = Flow_to_Lower - M_wall_l + frac_to_lower * M_wall - massplume(j) - NumberFans(j) * fanflowout
                End If
            Else
                'pressurising the lower layer
                If UseFanCurve(j) = True Then
                    'work out the cross-fan pressure difference for a fan at floor level
                    'CrossFanPreDiff = Y(j, 4) - (-ReferenceDensity * G * FloorElevation(j))
                    If FanElevation(j) > layerz Then
                        'fan in upper layer
                        CrossFanPreDiff = (Y(j, 4) - denl * G * layerz - denu * G * (FanElevation(j) - layerz)) - (-(Atm_Pressure) / (Gas_Constant / MW_air) / ExteriorTemp * G * (FanElevation(j) + FloorElevation(j)))
                    Else
                        'fan in lower layer
                        CrossFanPreDiff = (Y(j, 4) - denl * G * FanElevation(j)) - (-(Atm_Pressure) / (Gas_Constant / MW_air) / ExteriorTemp * G * (FanElevation(j) + FloorElevation(j)))
                    End If

                    If CrossFanPreDiff < 0 Then CrossFanPreDiff = -CrossFanPreDiff

                    If CrossFanPreDiff > MaxPressure(j) Then
                        NewExtractRate = -ExtractRate(j) * ((CrossFanPreDiff - MaxPressure(j)) / MaxPressure(j)) ^ (1 / 6)
                        'NewExtractRate = 0
                    Else
                        NewExtractRate = ExtractRate(j) * ((MaxPressure(j) - CrossFanPreDiff) / MaxPressure(j)) ^ (1 / 6)
                    End If
                Else
                    NewExtractRate = ExtractRate(j)
                End If

                fanflowout = NewExtractRate * Atm_Pressure / Gas_Constant_Air / ExteriorTemp
200:
                If (X - ExtractStartTime(j)) < 30 Then
                    fanflowout = (X - ExtractStartTime(j)) / 30 * fanflowout
                End If
                'If FanElevation(j) < layerz Then
                If Y(j, 2) > Y(j, 3) Then
                    If ExteriorTemp >= Y(j, 2) Then
                        'deposit all in upper layer
                        'pressurising the upper layer
                        FU = 1
                    ElseIf ExteriorTemp <= Y(j, 3) Then
                        'deposit all in lower layer
                        FU = 0
                    Else
                        'exterior temp lies between upper and lower layer temps
                        FU = (ExteriorTemp - Y(j, 3)) / (Y(j, 2) - Y(j, 3))
                    End If
                Else
                    'upper layer temp<=lower layer temp
                    If ExteriorTemp > Y(j, 2) Then
                        FU = 1
                    ElseIf ExteriorTemp < Y(j, 2) Then
                        FU = 0
                    ElseIf FanElevation(j) > layerz Then
                        FU = 1
                    ElseIf FanElevation(j) < layerz Then
                        FU = 0
                    Else
                        FU = 0.5
                    End If
                End If

                species_to_upper(3) = species_to_upper(3) + NumberFans(j) * FU * fanflowout * O2MassFraction(j, 1, 2) 'oxygen
                species_to_upper(4) = species_to_upper(4) + NumberFans(j) * FU * fanflowout * COMassFraction(j, 1, 2) 'co
                species_to_upper(5) = species_to_upper(5) + NumberFans(j) * FU * fanflowout * CO2MassFraction(j, 1, 2) 'co2
                species_to_upper(6) = species_to_upper(6) + NumberFans(j) * FU * fanflowout * SootMassFraction(j, 1, 2) 'soot
                species_to_upper(7) = species_to_upper(7) + NumberFans(j) * FU * fanflowout * H2OMassFraction(j, 1, 2) 'h20
                species_to_upper(8) = species_to_upper(8) + NumberFans(j) * FU * fanflowout * TUHC(j, 1, 2) 'tuhc
                species_to_upper(9) = species_to_upper(9) + NumberFans(j) * FU * fanflowout * HCNMassFraction(j, 1, 2) 'hcn
                Enthalpy_U = Enthalpy_U + NumberFans(j) * FU * fanflowout * SpecificHeat_air * ExteriorTemp
                species_to_lower(3) = species_to_lower(3) + NumberFans(j) * (1 - FU) * fanflowout * O2MassFraction(j, 1, 2) 'oxygen
                species_to_lower(4) = species_to_lower(4) + NumberFans(j) * (1 - FU) * fanflowout * COMassFraction(j, 1, 2) 'co
                species_to_lower(5) = species_to_lower(5) + NumberFans(j) * (1 - FU) * fanflowout * CO2MassFraction(j, 1, 2) 'co2
                species_to_lower(6) = species_to_lower(6) + NumberFans(j) * (1 - FU) * fanflowout * SootMassFraction(j, 1, 2) 'soot
                species_to_lower(7) = species_to_lower(7) + NumberFans(j) * (1 - FU) * fanflowout * H2OMassFraction(j, 1, 2) 'h20
                species_to_lower(8) = species_to_lower(8) + NumberFans(j) * (1 - FU) * fanflowout * TUHC(j, 1, 2) 'tuhc
                species_to_lower(9) = species_to_lower(9) + NumberFans(j) * (1 - FU) * fanflowout * HCNMassFraction(j, 1, 2) 'hcn
                Enthalpy_L = Enthalpy_L + NumberFans(j) * (1 - FU) * fanflowout * SpecificHeat_air * ExteriorTemp
                dMudt = massplume(j) - M_wall_u + (1 - frac_to_lower) * M_wall + fuelmassloss(j) + Flow_to_Upper + NumberFans(j) * FU * fanflowout
                dMldt = Flow_to_Lower - M_wall_l + (frac_to_lower) * M_wall - massplume(j) + NumberFans(j) * (1 - FU) * fanflowout
            End If
        Else
            'fan not yet started
            dMudt = massplume(j) - M_wall_u + (1 - frac_to_lower) * M_wall + fuelmassloss(j) + Flow_to_Upper
            dMldt = Flow_to_Lower - M_wall_l + (frac_to_lower) * M_wall - massplume(j)
        End If

    End Sub

    Public Sub mech_vent2(ByRef upperfans As Single, ByRef lowerfans As Single, ByRef fanid As Integer, ByRef X As Double, ByRef j As Integer, ByRef layerz As Double, ByRef Y(,) As Double, ByRef denl As Double, ByRef denu As Double, ByRef species_to_upper() As Double, ByRef species_to_lower() As Double, ByRef Enthalpy_U As Double, ByRef Enthalpy_L As Double, ByRef cp_u() As Double, ByRef cp_l() As Double, ByRef massplume() As Double, ByRef M_wall_u As Double, ByRef M_wall_l As Double, ByRef fuelmassloss() As Double, ByRef Flow_to_Upper As Double, ByRef Flow_to_Lower As Double, ByRef dMudt As Double, ByRef dMldt As Double, ByRef M_wall As Double, ByRef frac_to_lower As Double)
        'to work with multiple fan objects
        'fandata(fanid, 0) = fanid
        'fandata(fanid, 1) = flowrate m3/s
        'fandata(fanid, 2) = max pressure limit (pa)
        'fandata(fanid, 3) = start time (sec)
        'fandata(fanid, 4) = fan manual true/false  1/0
        'fandata(fanid, 5) = fan curve true/false 1/0
        'fandata(fanid, 6) = elevation (m)
        'fandata(fanid, 7) = fan extract/pressurise 1/0
        'fandata(fanid, 8) = fan reliability

        Dim Mcrit, NewExtractRate, CrossFanPreDiff, fanflowout, FU As Double

        If X > fandata(fanid, 3) And FANactive(fanid) = True Then 'start time
            If fandata(fanid, 7) = -1 Then 'extract 
                If fandata(fanid, 5) = -1 Then 'fancurve
                    'work out the cross-fan pressure difference for a fan at ceiling level
                    If fandata(fanid, 6) > layerz Then
                        'fan in upper layer
                        'pressure just outside less pressure just inside fan location
                        CrossFanPreDiff = (0 - (Atm_Pressure) / (Gas_Constant / MW_air) / ExteriorTemp * G * (fandata(fanid, 6) + FloorElevation(j))) - (Y(j, 4) - denl * G * layerz - denu * G * (fandata(fanid, 6) - layerz))
                    Else
                        'fan in lower layer
                        CrossFanPreDiff = (0 - (Atm_Pressure) / (Gas_Constant / MW_air) / ExteriorTemp * G * (fandata(fanid, 6) + FloorElevation(j))) - (Y(j, 4) - denl * G * (fandata(fanid, 6)))
                    End If
                    If CrossFanPreDiff > fandata(fanid, 2) Then
                        NewExtractRate = -fandata(fanid, 1) * ((CrossFanPreDiff - fandata(fanid, 2)) / fandata(fanid, 2)) ^ (1 / 6)
                        fanflowout = -NewExtractRate * Atm_Pressure / Gas_Constant_Air / ExteriorTemp
                        GoTo 200
                    Else
                        NewExtractRate = fandata(fanid, 1) * ((fandata(fanid, 2) - CrossFanPreDiff) / fandata(fanid, 2)) ^ (1 / 6)
                    End If
                Else
                    NewExtractRate = fandata(fanid, 1)
                End If

                If CSng(Y(j, 1)) > CSng(0.001 * RoomVolume(j)) And fandata(fanid, 6) >= layerz Then
                    'extracting from the upper layer
                    fanflowout = NewExtractRate * denu 'kg/s
                    If (X - fandata(fanid, 3)) < 30 Then
                        fanflowout = (X - fandata(fanid, 3)) / 30 * fanflowout 'kg/sec
                    End If

                    'calculate max flow to avoid plugholing
                    If (1 - (denu / denl)) > 0 Then
                        Mcrit = denu * 1.6 * (RoomHeight(j) - layerz) ^ 2.5 * Sqrt(G * (1 - (denu / denl)))
                    Else
                        Mcrit = 0
                    End If
                    'If X > 30 Then Stop
                    If fanflowout > Mcrit Then
                        'plugholing
                        FU = Mcrit / fanflowout

                    Else
                        'no plugholing
                        FU = 1

                    End If

                    species_to_upper(3) = species_to_upper(3) - FU * fanflowout * Y(j, 5) 'oxygen
                    species_to_upper(4) = species_to_upper(4) - FU * fanflowout * Y(j, 8) 'co
                    species_to_upper(5) = species_to_upper(5) - FU * fanflowout * Y(j, 10) 'co2
                    species_to_upper(6) = species_to_upper(6) - FU * fanflowout * Y(j, 12) 'soot
                    species_to_upper(7) = species_to_upper(7) - FU * fanflowout * Y(j, 16) 'h20
                    species_to_upper(8) = species_to_upper(8) - FU * fanflowout * Y(j, 6) 'tuhc
                    species_to_upper(9) = species_to_upper(9) - FU * fanflowout * Y(j, 18) 'hcn
                    Enthalpy_U = Enthalpy_U - FU * fanflowout * cp_u(j) * Y(j, 2)

                    species_to_lower(3) = species_to_lower(3) - (1 - FU) * fanflowout * Y(j, 15) 'oxygen
                    species_to_lower(4) = species_to_lower(4) - (1 - FU) * fanflowout * Y(j, 9) 'co
                    species_to_lower(5) = species_to_lower(5) - (1 - FU) * fanflowout * Y(j, 11) 'co2
                    species_to_lower(6) = species_to_lower(6) - (1 - FU) * fanflowout * Y(j, 13) 'soot
                    species_to_lower(7) = species_to_lower(7) - (1 - FU) * fanflowout * Y(j, 17) 'h20
                    species_to_lower(8) = species_to_lower(8) - (1 - FU) * fanflowout * Y(j, 7) 'tuhc
                    species_to_lower(9) = species_to_lower(9) - (1 - FU) * fanflowout * Y(j, 14) 'hcn
                    Enthalpy_L = Enthalpy_L - (1 - FU) * fanflowout * cp_l(j) * Y(j, 3)

                    'dMudt = massplume(j) - M_wall_u + (1 - frac_to_lower) * M_wall + fuelmassloss(j) + Flow_to_Upper - FU * fanflowout
                    'dMldt = Flow_to_Lower - M_wall_l + frac_to_lower * M_wall - massplume(j) - (1 - FU) * fanflowout
                Else
                    'extracting from the lower layer
                    fanflowout = NewExtractRate * denl 'kg/s
                    If (X - fandata(fanid, 3)) < 30 Then
                        fanflowout = (X - fandata(fanid, 3)) / 30 * fanflowout
                    End If
                    species_to_lower(3) = species_to_lower(3) - fanflowout * Y(j, 15) 'oxygen
                    species_to_lower(4) = species_to_lower(4) - fanflowout * Y(j, 9) 'co
                    species_to_lower(5) = species_to_lower(5) - fanflowout * Y(j, 11) 'co2
                    species_to_lower(6) = species_to_lower(6) - fanflowout * Y(j, 13) 'soot
                    species_to_lower(7) = species_to_lower(7) - fanflowout * Y(j, 17) 'h20
                    species_to_lower(8) = species_to_lower(8) - fanflowout * Y(j, 7) 'tuhc
                    species_to_lower(9) = species_to_lower(9) - fanflowout * Y(j, 14) 'hcn
                    Enthalpy_L = Enthalpy_L - fanflowout * cp_l(j) * Y(j, 3)

                    FU = 0
                    'dMudt = massplume(j) - M_wall_u + (1 - frac_to_lower) * M_wall + fuelmassloss(j) + Flow_to_Upper
                    'dMldt = Flow_to_Lower - M_wall_l + frac_to_lower * M_wall - massplume(j) - fanflowout
                End If
            Else
                'pressurising the lower layer
                If fandata(fanid, 5) = -1 Then
                    'work out the cross-fan pressure difference for a fan at floor level
                    'CrossFanPreDiff = Y(j, 4) - (-ReferenceDensity * G * FloorElevation(j))
                    If fandata(fanid, 6) > layerz Then
                        'fan in upper layer
                        CrossFanPreDiff = (Y(j, 4) - denl * G * layerz - denu * G * (fandata(fanid, 6) - layerz)) - (-(Atm_Pressure) / (Gas_Constant / MW_air) / ExteriorTemp * G * (fandata(fanid, 6) + FloorElevation(j)))
                    Else
                        'fan in lower layer
                        CrossFanPreDiff = (Y(j, 4) - denl * G * fandata(fanid, 6)) - (-(Atm_Pressure) / (Gas_Constant / MW_air) / ExteriorTemp * G * (fandata(fanid, 6) + FloorElevation(j)))
                    End If

                    If CrossFanPreDiff < 0 Then CrossFanPreDiff = -CrossFanPreDiff

                    If CrossFanPreDiff > fandata(fanid, 2) Then
                        NewExtractRate = -fandata(fanid, 1) * ((CrossFanPreDiff - fandata(fanid, 2)) / fandata(fanid, 2)) ^ (1 / 6)
                    Else
                        NewExtractRate = fandata(fanid, 1) * ((fandata(fanid, 2) - CrossFanPreDiff) / fandata(fanid, 2)) ^ (1 / 6)
                    End If
                Else
                    'not using fan curve
                    NewExtractRate = fandata(fanid, 1)
                End If

                fanflowout = NewExtractRate * Atm_Pressure / Gas_Constant_Air / ExteriorTemp
200:
                If (X - fandata(fanid, 3)) < 30 Then
                    fanflowout = (X - fandata(fanid, 3)) / 30 * fanflowout
                End If
                'If FanElevation(j) < layerz Then
                If Y(j, 2) > Y(j, 3) Then
                    If ExteriorTemp >= Y(j, 2) Then
                        'deposit all in upper layer
                        'pressurising the upper layer
                        FU = 1
                    ElseIf ExteriorTemp <= Y(j, 3) Then
                        'deposit all in lower layer
                        FU = 0
                    Else
                        'exterior temp lies between upper and lower layer temps
                        FU = (ExteriorTemp - Y(j, 3)) / (Y(j, 2) - Y(j, 3))
                    End If
                Else
                    'upper layer temp<=lower layer temp
                    If ExteriorTemp > Y(j, 2) Then
                        FU = 1
                    ElseIf ExteriorTemp < Y(j, 2) Then
                        FU = 0
                    ElseIf fandata(fanid, 6) > layerz Then
                        FU = 1
                    ElseIf fandata(fanid, 6) < layerz Then
                        FU = 0
                    Else
                        FU = 0.5
                    End If
                End If

                'expt, lets only add to upper layer
                'FU = 0

                species_to_upper(3) = species_to_upper(3) + FU * fanflowout * O2MassFraction(j, 1, 2) 'oxygen
                'species_to_upper(4) = species_to_upper(4) + FU * fanflowout * COMassFraction(j, 1, 2) 'co
                species_to_upper(5) = species_to_upper(5) + FU * fanflowout * CO2MassFraction(j, 1, 2) 'co2
                'species_to_upper(6) = species_to_upper(6) + FU * fanflowout * SootMassFraction(j, 1, 2) 'soot
                species_to_upper(7) = species_to_upper(7) + FU * fanflowout * H2OMassFraction(j, 1, 2) 'h20
                'species_to_upper(8) = species_to_upper(8) + FU * fanflowout * TUHC(j, 1, 2) 'tuhc
                'species_to_upper(9) = species_to_upper(9) + FU * fanflowout * HCNMassFraction(j, 1, 2) 'hcn
                Enthalpy_U = Enthalpy_U + FU * fanflowout * SpecificHeat_air * ExteriorTemp

                species_to_lower(3) = species_to_lower(3) + (1 - FU) * fanflowout * O2MassFraction(j, 1, 2) 'oxygen
                'species_to_lower(4) = species_to_lower(4) + (1 - FU) * fanflowout * COMassFraction(j, 1, 2) 'co
                species_to_lower(5) = species_to_lower(5) + (1 - FU) * fanflowout * CO2MassFraction(j, 1, 2) 'co2
                'species_to_lower(6) = species_to_lower(6) + (1 - FU) * fanflowout * SootMassFraction(j, 1, 2) 'soot
                species_to_lower(7) = species_to_lower(7) + (1 - FU) * fanflowout * H2OMassFraction(j, 1, 2) 'h20
                'species_to_lower(8) = species_to_lower(8) + (1 - FU) * fanflowout * TUHC(j, 1, 2) 'tuhc
                'species_to_lower(9) = species_to_lower(9) + (1 - FU) * fanflowout * HCNMassFraction(j, 1, 2) 'hcn
                Enthalpy_L = Enthalpy_L + (1 - FU) * fanflowout * SpecificHeat_air * ExteriorTemp

                fanflowout = -fanflowout
                'dMudt = massplume(j) - M_wall_u + (1 - frac_to_lower) * M_wall + fuelmassloss(j) + Flow_to_Upper + FU * fanflowout
                'dMldt = Flow_to_Lower - M_wall_l + (frac_to_lower) * M_wall - massplume(j) + (1 - FU) * fanflowout
            End If
        Else
            'fan not yet started
            fanflowout = 0
            'dMudt = massplume(j) - M_wall_u + (1 - frac_to_lower) * M_wall + fuelmassloss(j) + Flow_to_Upper
            'dMldt = Flow_to_Lower - M_wall_l + (frac_to_lower) * M_wall - massplume(j)
        End If

        upperfans = upperfans + FU * fanflowout
        lowerfans = lowerfans + (1 - FU) * fanflowout

        'dMudt = massplume(j) - M_wall_u + (1 - frac_to_lower) * M_wall + fuelmassloss(j) + Flow_to_Upper - FU * fanflowout
        'dMldt = Flow_to_Lower - M_wall_l + frac_to_lower * M_wall - massplume(j) - (1 - FU) * fanflowout

    End Sub

End Module