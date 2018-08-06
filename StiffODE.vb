Imports CenterSpace.NMath.Core
Imports CenterSpace.NMath.Analysis
Imports System.Math
Namespace CenterSpace.NMath.Analysis.Examples.VisualBasic

    ' <summary>
    ' A .NET Example in Visual Basic showing how to use the RungeKutta45OdeSolver 
    ' to solve a nonstiff set of equations describing the motion of a rigid body 
    ' without external forces:
    ' 
    ' y1' = y2*y3,       y1(0) = 0
    ' y2' = -y1*y3,      y2(0) = 1
    ' y3' = -0.51*y1*y2, y3(0) = 1
    ' </summary>
    Module StiffSolverExample

        Sub Main()

            ' This section shows to use various configuration settings. All of these settings can
            ' also be set via configuration file or by environment variable. More on this at: 
            ' http://www.centerspace.net/blog/nmath/nmath-configuration/

            ' For NMath to continue to work after your evaluation period, you must set your license 
            ' key. You will receive a license key after purchase. 
            ' More here:
            ' http://www.centerspace.net/blog/nmath/setting-the-nmath-license-key/
            NMathConfiguration.LicenseKey = "XTUU-3UM1-YXFF-89WC-DAUS-VUUF-1QS1-WPUF-19SU-WPSS-8EJE-680T-FRPP-X8NS-WQSF-E842-68RT-WPTF-683S-X1D0-YR8V-19P2-9D3Y-LRTU-A10U"

            ' This will start a log file that you can use to ensure that your configuration is 
            ' correct. This can be especially useful
            ' for deployment. Please turn this off when you are convinced everything is 
            ' working. If you aren't sure, please send the resulting log file to 
            ' support@centerspace.net. Please note that this directory can be relative or absolute.
            ' NMathConfiguration.LogLocation = "<dir>"

            ' NMath loads native, optimized libraries at runtime. There is a one-time cost to 
            ' doing so. To take control of when this happens, use Init(). If your program calls 
            ' Init() successfully, then your configuration is definitely correct.
            ' NMathConfiguration.Init()


            Console.WriteLine()

            ' Simple example solving the system of differential equations of
            ' motion for a rigid body with no external forces. The equations
            ' are defined above by the function Rigid().
            Console.WriteLine("Simple Solve -----------------------------------")
            Console.WriteLine()
            StiffSolve()
            Console.WriteLine(Environment.NewLine)

            ' A mass matrix and an output function are options that may
            ' be specified for the solver. The output function is a callback
            ' that is invoked at initialization, each integration step and 
            ' after the conclusion of the last integration step. Thus the
            ' solution process can be monitored, and even terminated Imports
            ' the output function.
            Console.WriteLine("Output Function and Mass Matrix Solve ---------")
            Console.WriteLine()
            WithOutputFunctionAndMassMatrix()
            Console.WriteLine(Environment.NewLine)

            Console.WriteLine()
            Console.WriteLine("Press Enter Key")
            Console.Read()

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

        ' <summary>
        ' Function describing the system of differential equations.
        ' </summary>
        ' <param name="t">Time parameter.</param>
        ' <param name="y">State vector.</param>
        ' <returns>The vector of values of the derivatives of the functions
        ' at a specified time t and function values y:
        ' y1' = y2*y3,
        ' y2' = -y1*y3,
        ' y3' = -0.51*y1*y2
        ' </returns>
        Function Rigid(ByVal T As Double, ByVal Y As DoubleVector) As DoubleVector

            Dim dy As New DoubleVector(3)
            dy(0) = Y(1) * Y(2)
            dy(1) = -Y(0) * Y(2)
            dy(2) = -0.51 * Y(0) * Y(1)
            Return dy

        End Function

        ' <summary>
        ' The output function to be called by the ODE solver at each integration step. 
        ' If this function returns true the solver continues with the next step, if 
        ' it returns false the solver is halted.
        ' This output function just outputs to the console upon each invocation.
        ' </summary>
        ' <param name="t">The time value at the current integration step. If flag is equal to Initialize,
        ' this will be the initial time value t0.</param>
        ' <param name="y">The calculated function value at the current integration step. If flag is equal to Initialize,
        ' this will be the initial function value y0.</param>
        ' <param name="flag">Flag indicating what stage the solver is at:
        ' Initialize - before solving begins. y and t values are the problems initial values.
        ' IntegrationStep - just finished an integration step. y and t values are the values
        ' calculated at that step.
        ' Done - just finished last step. y and t values are the last values in the returned 
        ' solution.</param>
        ' <returns>true if the solver is to proceed with the calculation, false forces the solver
        ' to stop.</returns>
        Function OutputFunction(ByVal T As DoubleVector, ByVal Y As DoubleMatrix, ByVal Flag As VariableOrderOdeSolver.OutputFunctionFlag) As Boolean

            If Flag = VariableOrderOdeSolver.OutputFunctionFlag.Initialize Then
                Console.WriteLine("Output function : Initialize")
                Return True
            End If

            If Flag = VariableOrderOdeSolver.OutputFunctionFlag.IntegrationStep Then
                Console.WriteLine("Output function Integration step:")
                Console.WriteLine("t = " + T.ToString("G5"))
                Console.WriteLine("y = ")
                Console.WriteLine(Y.ToTabDelimited("G5"))
                Return True
            End If

            If Flag = VariableOrderOdeSolver.OutputFunctionFlag.Done Then
                Console.WriteLine("Output function: Done")
                Return True
            End If

            Console.WriteLine("Output function: Unknown flag")
            Return False

        End Function

        Sub WithOutputFunctionAndMassMatrix()

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
            ' e(i) <= max(RelativeTolerance*Math.Abs(y(i)), AbsoluteTolerance(i))
            Dim SolverOptions As New VariableOrderOdeSolver.Options()
            SolverOptions.AbsoluteTolerance = New DoubleVector(0.0001, 0.0001, 0.00001)

            SolverOptions.OutputFunction = New Func(Of DoubleVector, DoubleMatrix, VariableOrderOdeSolver.OutputFunctionFlag, Boolean)(AddressOf OutputFunction)
            SolverOptions.RelativeTolerance = 0.0001

            ' Construct the delegate representing our system of differential equations...
            Dim odeFunction As New Func(Of Double, DoubleVector, DoubleVector)(AddressOf Rigid)

            ' ...and solve. The solution is returned as a key/value pair. The first 'Key' element of the pair is
            ' the time span vector, the second 'Value' element of the pair is the corresponding solution values.
            ' That is, if the computed solution function is y then
            ' y(soln.Key(i)) = soln.Value(i)
            Dim Soln As VariableOrderOdeSolver.Solution(Of DoubleMatrix) = Solver.Solve(odeFunction, TimeSpan, y0, SolverOptions)

            ' Print out a few values
            Console.WriteLine("Number of time values = " & Soln.T.Length)
            For I As Integer = 0 To Soln.T.Length - 1
                If (I Mod 7 = 0) Then
                    Console.WriteLine("y({0}) = {1}", Soln.T(I).ToString("G5"), Soln.Y.Row(I).ToString("G5"))
                End If
            Next

            ' We can also specify a mass matrix in the solver options...
            Dim M As New DoubleMatrix("3x3[1     2     1 2     1     3     9     4     1]")
            SolverOptions.MassMatrix = M
            Soln = Solver.Solve(odeFunction, TimeSpan, y0, SolverOptions)

            ' Print out a few values for the mass matrix solution.
            Console.WriteLine(Environment.NewLine & "With constant mass matrix:")
            Console.WriteLine("Number of time values = " & Soln.T.Length)
            For I = 0 To Soln.T.Length - 1
                If (I Mod 7 = 0) Then
                    Console.WriteLine("y({0}) = {1}", Soln.T(I).ToString("G5"), Soln.Y.Row(I).ToString("G5"))
                End If
            Next

            ' You can also solve with default values for all the options
            Soln = Solver.Solve(odeFunction, TimeSpan, y0, SolverOptions)

        End Sub

        '        Private Sub diff_eqns_oneroom(ByVal room As Integer, ByVal X As Double, ByRef Y() As Double, ByRef DYDX() As Double)

        '            Dim r2u, deltaEU, sdot, deltaEL, beta, r2l As Double
        '            Dim response As Short
        '            Dim k, i, h, j As Integer
        '            Dim dMudt, dMldt As Double
        '            Dim doormix As Double
        '            Dim Enthalpy_L, denl, denu, Enthalpy_U, sum As Double
        '            Dim layerz As Double
        '            Dim massplume As Double
        '            Dim rate_Renamed As Double
        '            Dim mrate() As Double
        '            Dim Enthalpy(1, 1) As Double
        '            Dim products(1, 1, 1) As Double
        '            Dim fuelmassloss As Double
        '            Dim ventflow(1, 1) As Double
        '            Dim Flow_to_Lower, Flow_to_Upper, F As Double
        '            Dim FlowOut As Double
        '            Dim Flowtoroom As Double
        '            Dim id As Integer
        '            Dim weighted_hc, absorb_l, absorb_u, hrr, mrate_floor As Double
        '            Dim hrrlimit As Double
        '            Dim massplumeX As Double
        '            Dim ventfire_sum() As Double
        '            Dim fanflowout As Double
        '            Dim mrate_wall As Double
        '            Dim mrate_ceiling As Double
        '            Dim vent As Integer
        '            Dim matc(,) As Double
        '            Dim Qplume, hrrnextroom, M_wall, burningrate As Double
        '            Dim mw_upper As Double
        '            Dim mw_lower As Double
        '            Dim species_to_upper(9) As Double
        '            Dim species_to_lower(9) As Double
        '            Dim genrate As Double
        '            Dim GER As Double
        '            Dim CJTemp, CJMax As Double
        '            Dim vententrain(1, 1, 1) As Double
        '            Dim stemp, ftemp As Double
        '            Dim Mass_Upper As Double
        '            Dim Mass_Lower As Double
        '            Dim M_wall_u, frac_to_lower, Mwu, Mwl, E_wall, M_wall_l As Double
        '            Dim cp_u As Double
        '            Dim cp_l As Double
        '            Dim gamma_l, gamma_u, species_mass As Double
        '            Dim wall_spec(7) As Double

        '            ReDim mrate(NumberObjects)
        '            ReDim ventfire_sum(NumberRooms + 1)
        '            ReDim matc(4, 1)

        '            If MaxNumberVents > 0 Then
        '                ReDim vententrain(NumberRooms, NumberRooms, MaxNumberVents)
        '            End If
        '            Static xflag As Short
        '            Static xlast As Double

        '            If X = xlast Then
        '                xflag = 1
        '            Else
        '                xflag = 0
        '            End If

        '            On Error GoTo errorhandler

        '            i = stepcount


        '            If Y(2) < 50 Or Y(2) > 3000 Then
        '                Y(2) = uppertemp(room, i)
        '            End If
        '            If Y(3) < 50 Or Y(3) > 3000 Then
        '                Y(3) = lowertemp(room, i)
        '            End If

        '            'check upper volume within valid range
        '            If TwoZones(room) = False Then

        '                Y(1) = RoomVolume(room) - layerheight(room, i) * RoomLength(room) * RoomWidth(room)
        '                layerz = layerheight(room, i)
        '            End If

        '            'put some limits on variable values
        '            If Y(1) < 0.0001 * RoomVolume(room) Then
        '                Y(1) = 0.0001 * RoomVolume(room)
        '            End If

        '            If Y(1) > RoomVolume(room) - Smallvalue Then
        '                Y(1) = RoomVolume(room) - Smallvalue
        '            End If

        '            If Y(12) < Smallvalue Then Y(12) = Smallvalue
        '            If Y(13) < Smallvalue Then Y(13) = Smallvalue
        '            If Y(14) < Smallvalue Then Y(14) = Smallvalue
        '            If Y(8) < Smallvalue Then Y(8) = Smallvalue
        '            If Y(9) < Smallvalue Then Y(9) = Smallvalue
        '            If Y(10) < CO2Ambient Then Y(10) = CO2Ambient
        '            If Y(11) < CO2Ambient Then Y(11) = CO2Ambient
        '            If Y(16) < H2OMassFraction(room, 1, 1) Then Y(16) = H2OMassFraction(room, 1, 1)
        '            If Y(17) < H2OMassFraction(room, 1, 2) Then Y(17) = H2OMassFraction(room, 1, 2)
        '            If Y(5) > O2Ambient Then Y(5) = O2Ambient 'upper oxygen not higher than ambient
        '            If Y(15) > O2Ambient Then
        '                Y(15) = O2Ambient 'lower oxygen not higher than ambient
        '            End If

        '            If Y(5) < 0.001 Then Y(5) = 0.001 'upper oxygen not less than 0.1%
        '            If Y(15) < 0.001 Then Y(15) = 0.001 'lower oxygen not less than 0.1%
        '            If Y(6) < TUHC(room, 1, 1) Then Y(6) = TUHC(room, 1, 1) 'upper tuhc not less than 0.01%
        '            If Y(7) < TUHC(room, 1, 2) Then Y(7) = TUHC(room, 1, 2) 'lower tuhc not less than 0.01%

        '            If TwoZones(room) = True Then
        '                Call Layer_Height(Y(1), layerz, room)
        '            End If

        '            'effective molecular weight of the layers for fireroom
        '            mw_upper = MolecularWeightCO * Y(8) + MolecularWeightCO2 * Y(10) + MolecularWeightH2O * Y(16) + MolecularWeightO2 * Y(5) + MolecularWeightN2 * (1 - Y(5) - Y(16) - Y(8) - Y(10))
        '            If mw_upper < 0 Then
        '                Exit Sub
        '            End If
        '            mw_lower = MolecularWeightCO * Y(9) + MolecularWeightCO2 * Y(11) + MolecularWeightH2O * Y(17) + MolecularWeightO2 * Y(15) + MolecularWeightN2 * (1 - Y(15) - Y(17) - Y(9) - Y(11))
        '            If mw_lower < 0 Then
        '                Exit Sub
        '            End If
        '            'specific heats
        '            cp_u = specific_heat_upper(room, Y, Y(2)) 'specific_heat("CO", Y(j, 2)) * Y(j, 8) + specific_heat("CO2", Y(j, 2)) * Y(j, 10) + specific_heat("H2O", Y(j, 2)) * Y(j, 16) + specific_heat("O2", Y(j, 2)) * Y(j, 5) + specific_heat("N2", Y(j, 2)) * (1 - Y(j, 5) - Y(j, 16) - Y(j, 8) - Y(j, 10))
        '            cp_l = specific_heat_lower(room, Y, Y(3)) 'specific_heat("CO", Y(j, 3)) * Y(j, 9) + specific_heat("CO2", Y(j, 3)) * Y(j, 11) + specific_heat("H2O", Y(j, 3)) * Y(j, 17) + specific_heat("O2", Y(j, 3)) * Y(j, 15) + specific_heat("N2", Y(j, 3)) * (1 - Y(j, 15) - Y(j, 17) - Y(j, 9) - Y(j, 11))

        '            gamma_u = gamma
        '            gamma_l = gamma

        '            If room = fireroom Then

        '                If FuelResponseEffects = True Then
        '                    massplumeX = Mass_Plume_2012(X, layerz, HeatRelease(room, stepcount, 1), Y(2), Y(3))
        '                    Call mass_rate_withfuelresponse(X, mrate, mrate_wall, mrate_ceiling, mrate_floor, Y(15), Y(3), massplumeX, Target(fireroom, i), burningrate)
        '                    hrrlimit = 1000 * EnergyYield(1) * burningrate
        '                    Call hrr_estimate_fuelresponse(fireroom, Mass_Upper, massplumeX, hrrlimit, X, layerz, Y(2), Y(3), Y(5), Y(15), mw_upper, mw_lower, Y(6), weighted_hc, Target(fireroom, i))
        '                Else
        '                    Call mass_rate(X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
        '                End If

        '                fuelmassloss = 0
        '                sum = 0
        '                ftemp = 0 'kg/sec
        '                stemp = 0
        '                If Flashover = True And g_post = True Then
        '                    fuelmassloss = mrate(1)
        '                    sum = mrate(1) * NewHoC_fuel 'kg/s.kJ/g
        '                Else
        '                    For id = 1 To NumberObjects
        '                        fuelmassloss = fuelmassloss + mrate(id)
        '                        sum = sum + mrate(id) * NewEnergyYield(id) 'kg/s.kJ/g
        '                        ftemp = ftemp + fuelmassloss 'kg/sec
        '                        stemp = stemp + sum 'kg/s.kJ/g
        '                        If id = 1 Then
        '                            fuelmassloss = fuelmassloss + mrate_wall + mrate_ceiling + mrate_floor 'kg/sec
        '                            sum = sum + mrate_wall * WallEffectiveHeatofCombustion(room) + mrate_ceiling * CeilingEffectiveHeatofCombustion(room) + mrate_floor * FloorEffectiveHeatofCombustion(room)
        '                        End If
        '                    Next id
        '                End If
        '                If fuelmassloss > 0 Then weighted_hc = sum / fuelmassloss 'MJ/kg or kJ/g
        '                If ftemp > 0 Then stemp = stemp / ftemp 'kJ/g
        '            End If

        '            'mass of the upper layer (kg)
        '            Mass_Lower = (RoomVolume(room) - Y(1)) * (Atm_Pressure + Y(4) / 1000) / (Gas_Constant / mw_lower) / Y(j, 3)
        '            Mass_Upper = Y(1) * (Atm_Pressure + Y(4) / 1000) / (Gas_Constant / mw_upper) / Y(2)


        '            Call hrr_estimate(fireroom, Mass_Upper, massplumeX, hrrlimit, X, layerz, Y(2), Y(3), Y(5), Y(15), mw_upper, mw_lower, Y(6), weighted_hc)

        '            'calculate wall vent flows
        '            'Call ventflows(Mass_Upper, vententrain, X, Y, ventflow, layerz, Enthalpy, products, ventfire_sum, mw_upper, mw_lower, weighted_hc, FlowOut, Flowtoroom, cp_u, cp_l)
        '            If flagstop = 1 Then Exit Sub


        '            denu = Mass_Upper / Y(1)
        '            denl = Mass_Lower / (RoomVolume(room) - Y(1))
        '            If room <> fireroom Then
        '                massplumeX = 0

        '                If fireroom < room Then
        '                    For vent = 1 To NumberVents(fireroom, room)
        '                        massplumeX = massplumeX + vententrain(fireroom, room, vent)
        '                    Next vent
        '                Else
        '                    For vent = 1 To NumberVents(room, fireroom)
        '                        massplumeX = massplumeX + vententrain(fireroom, room, vent)
        '                    Next vent
        '                End If
        '                If HeatRelease(room, stepcount, 1) + ventfire_sum(room) > 0 Then
        '                    hrrnextroom = O2_limit_cfast(room, Mass_Upper, HeatRelease(room, stepcount, 1) + ventfire_sum(room), massplumeX, Y(2), Y(5), Y(15), mw_upper, mw_lower, Qplume)
        '                End If
        '            End If

        '            Flow_to_Upper = 0
        '            Flow_to_Lower = 0
        '            Enthalpy_U = 0
        '            Enthalpy_L = 0
        '            For h = 3 To 9
        '                species_to_upper(h) = 0
        '                species_to_lower(h) = 0
        '            Next h


        '            Flow_to_Upper = ventflow(room, 2)

        '            Flow_to_Lower = ventflow(room, 1)

        '            If room = fireroom Then 'Mass flow in the plume
        '                massplume = massplumeX
        '                If massplume > 0.0001 Then

        '                    GER = weighted_hc * fuelmassloss / (13.1 * massplume * Y(15)) 'plume GER

        '                Else
        '                    GER = 0
        '                End If
        '            Else
        '                massplume = 0
        '                If IgniteNextRoom = True And X > CeilingIgniteTime(room) Then
        '                    fuelmassloss = (hrrnextroom - ventfire_sum(room)) / (CeilingEffectiveHeatofCombustion(room) * 1000)
        '                    If fuelmassloss < 0 Then fuelmassloss = 0
        '                Else
        '                    fuelmassloss = 0
        '                End If
        '                If massplumeX > 0 Then
        '                    GER = weighted_hc * fuelmassloss / (deltaHair * massplumeX)
        '                Else
        '                    GER = 0
        '                End If
        '            End If

        '            If GER > 0.5 And room = fireroom Then
        '                If frmOptions1.chkModGER.CheckState = System.Windows.Forms.CheckState.Checked Then
        '                    For k = 1 To NumberObjects
        '                        NewEnergyYield(k) = EnergyYield(k) * (1 - (0.97 / (Exp((GER / 2.15) ^ (-1.2))))) 'SFPE h/b 3rd ed page 3-118 eqn 40
        '                        NewEnergyYield(k) = EnergyYield(k)
        '                        NewRadiantLossFraction(k) = 1 - ((1 - ObjectRLF(k)) * (1 - 1 / Exp((GER / 1.38) ^ (-2.8))) / (1 - (0.97) / (Exp((GER / 2.15) ^ (-1.2)))))

        '                    Next k
        '                    'If Flashover = False Then NewHoC_fuel = HoC_fuel * (1 - (0.97 / (Exp((GER / 2.15) ^ (-1.2)))))
        '                    NewHoC_fuel = HoC_fuel * (1 - (0.97 / (Exp((GER / 2.15) ^ (-1.2)))))
        '                Else
        '                    For k = 1 To NumberObjects
        '                        NewEnergyYield(k) = EnergyYield(k)
        '                        NewRadiantLossFraction(k) = ObjectRLF(k)
        '                    Next k
        '                    NewHoC_fuel = HoC_fuel

        '                End If
        '            ElseIf room = fireroom Then
        '                For k = 1 To NumberObjects
        '                    NewEnergyYield(k) = EnergyYield(k)
        '                    NewRadiantLossFraction(k) = ObjectRLF(k)
        '                Next k
        '                NewHoC_fuel = HoC_fuel

        '            End If

        '            If i < NumberTimeSteps + 1 Then
        '                If room = fireroom Then
        '                    Call heat_losses_stiff2(room, X, layerz, Y(2), Y(3), Upperwalltemp(room, i + 1), CeilingTemp(room, i + 1), LowerWallTemp(room, i + 1), FloorTemp(room, i + 1), hrrlimit, absorb_u, absorb_l, matc, Y(10), Y(16), Y(11), Y(17), Y(1))
        '                Else
        '                    Call heat_losses_stiff2(room, X, layerz, Y(2), Y(3), Upperwalltemp(room, i + 1), CeilingTemp(room, i + 1), LowerWallTemp(room, i + 1), FloorTemp(room, i + 1), hrrnextroom, absorb_u, absorb_l, matc, Y(10), Y(16), Y(11), Y(17), Y(1))
        '                End If
        '            End If
        '            Enthalpy_U = Enthalpy(room, 2) / 1000 'kw
        '            Enthalpy_L = Enthalpy(room, 1) / 1000 'kw

        '            For h = 3 To 9
        '                species_to_upper(h) = products(h, room, 2) 'kg species/sec in ventflows
        '                species_to_lower(h) = products(h, room, 1) 'kg species/sec in ventflows
        '            Next h

        '            If room = fireroom Then
        '                Enthalpy_U = Enthalpy_U + massplume * cp_l * Y(3) + absorb_u + (1 - NewRadiantLossFraction(1)) * hrrlimit + fuelmassloss * SpecificHeat_fuel * Flametemp
        '                Enthalpy_L = Enthalpy_L - massplume * cp_l * Y(3) + absorb_l
        '            Else
        '                If IgniteNextRoom = True And X > CeilingIgniteTime(room) Then
        '                    Enthalpy_U = Enthalpy_U + absorb_u + (1 - NewRadiantLossFraction(1)) * hrrnextroom
        '                Else
        '                    Enthalpy_U = Enthalpy_U + absorb_u
        '                End If
        '                Enthalpy_L = Enthalpy_L + absorb_l
        '            End If

        '            frac_to_lower = 0
        '            M_wall = 0
        '            M_wall_u = 0
        '            M_wall_l = 0

        '            If room = fireroom And nowallflow = 0 Then
        '                'wall flows
        '                If Y(2) > Upperwalltemp(room, i) Then
        '                    Mwu = wall_flow_momentum(Mass_Upper / Y(1), X, room, Y(2), Upperwalltemp(room, i), layerz, RoomHeight(room) - layerz)
        '                    'flow from upper layer to lower layer
        '                    M_wall_u = wall_flow(Mass_Upper / Y(1), X, room, Y(2), Upperwalltemp(room, i), layerz, RoomHeight(room) - layerz) 'kg/s
        '                Else
        '                    Mwu = 0
        '                    M_wall_u = 0
        '                End If

        '                If Y(3) < LowerWallTemp(room, i) Then
        '                    Mwl = wall_flow_momentum(Mass_Lower / (RoomVolume(room) - Y(1)), X, room, Y(3), LowerWallTemp(room, i), layerz, layerz)
        '                    'flow from lower layer to upper layer
        '                    M_wall_l = wall_flow(Mass_Lower / (RoomVolume(room) - Y(1)), X, room, Y(3), LowerWallTemp(room, i), layerz, layerz)
        '                Else
        '                    Mwl = 0
        '                    M_wall_l = 0
        '                End If

        '                'remove flow from source
        '                species_to_upper(3) = species_to_upper(3) - M_wall_u * Y(5) 'oxygen
        '                species_to_upper(4) = species_to_upper(4) - M_wall_u * Y(8) 'co
        '                species_to_upper(5) = species_to_upper(5) - M_wall_u * Y(10) 'co2
        '                species_to_upper(6) = species_to_upper(6) - M_wall_u * Y(12) 'soot
        '                species_to_upper(7) = species_to_upper(7) - M_wall_u * Y(16) 'h20
        '                species_to_upper(8) = species_to_upper(8) - M_wall_u * Y(6) 'tuhc
        '                species_to_upper(9) = species_to_upper(9) - M_wall_u * Y(18) 'hcn

        '                species_to_lower(3) = species_to_lower(3) - M_wall_l * Y(15) 'oxygen
        '                species_to_lower(4) = species_to_lower(4) - M_wall_l * Y(9) 'co
        '                species_to_lower(5) = species_to_lower(5) - M_wall_l * Y(11) 'co2
        '                species_to_lower(6) = species_to_lower(6) - M_wall_l * Y(13) 'soot
        '                species_to_lower(7) = species_to_lower(7) - M_wall_l * Y(17) 'h20
        '                species_to_lower(8) = species_to_lower(8) - M_wall_l * Y(7) 'tuhc
        '                species_to_lower(9) = species_to_lower(9) - M_wall_l * Y(14) 'hcn
        '                Enthalpy_U = Enthalpy_U - M_wall_u * Y(2) * cp_u
        '                Enthalpy_L = Enthalpy_L - M_wall_l * Y(3) * cp_l

        '                'distribute flow to destinations
        '                M_wall = M_wall_u + M_wall_l
        '                If M_wall > 0 Then
        '                    frac_to_lower = Mwu / (Mwu + Mwl)

        '                    E_wall = M_wall_u * Y(2) * cp_u + M_wall_l * Y(3) * cp_l
        '                    wall_spec(1) = (M_wall_u * Y(5) + M_wall_l * Y(15)) / M_wall 'oxy
        '                    wall_spec(2) = (M_wall_u * Y(8) + M_wall_l * Y(9)) / M_wall
        '                    wall_spec(3) = (M_wall_u * Y(10) + M_wall_l * Y(11)) / M_wall
        '                    wall_spec(4) = (M_wall_u * Y(12) + M_wall_l * Y(13)) / M_wall
        '                    wall_spec(5) = (M_wall_u * Y(16) + M_wall_l * Y(17)) / M_wall
        '                    wall_spec(6) = (M_wall_u * Y(6) + M_wall_l * Y(7)) / M_wall
        '                    wall_spec(7) = (M_wall_u * Y(18) + M_wall_l * Y(14)) / M_wall

        '                    species_to_upper(3) = species_to_upper(3) + M_wall * (1 - frac_to_lower) * wall_spec(1) 'oxygen
        '                    species_to_upper(4) = species_to_upper(4) + M_wall * (1 - frac_to_lower) * wall_spec(2) 'co
        '                    species_to_upper(5) = species_to_upper(5) + M_wall * (1 - frac_to_lower) * wall_spec(3) 'co2
        '                    species_to_upper(6) = species_to_upper(6) + M_wall * (1 - frac_to_lower) * wall_spec(4) 'soot
        '                    species_to_upper(7) = species_to_upper(7) + M_wall * (1 - frac_to_lower) * wall_spec(5) 'h20
        '                    species_to_upper(8) = species_to_upper(8) + M_wall * (1 - frac_to_lower) * wall_spec(6) 'tuhc
        '                    species_to_upper(9) = species_to_upper(9) + M_wall * (1 - frac_to_lower) * wall_spec(7) 'hcn

        '                    species_to_lower(3) = species_to_lower(3) + M_wall * frac_to_lower * wall_spec(1) 'oxygen
        '                    species_to_lower(4) = species_to_lower(4) + M_wall * frac_to_lower * wall_spec(2) 'co
        '                    species_to_lower(5) = species_to_lower(5) + M_wall * frac_to_lower * wall_spec(3) 'co2
        '                    species_to_lower(6) = species_to_lower(6) + M_wall * frac_to_lower * wall_spec(4) 'soot
        '                    species_to_lower(7) = species_to_lower(7) + M_wall * frac_to_lower * wall_spec(5) 'h20
        '                    species_to_lower(8) = species_to_lower(8) + M_wall * frac_to_lower * wall_spec(6) 'tuhc
        '                    species_to_lower(9) = species_to_lower(9) + M_wall * frac_to_lower * wall_spec(7) 'hcn
        '                    Enthalpy_U = Enthalpy_U + E_wall * (1 - frac_to_lower)
        '                    Enthalpy_L = Enthalpy_L + E_wall * frac_to_lower
        '                End If

        '            End If

        '            If X = tim(i, 1) Then

        '                If room = fireroom Then GlobalER(i) = GER
        '                FuelMassLossRate(i, room) = fuelmassloss
        '                If room = fireroom Then WoodBurningRate(i) = mrate_ceiling + mrate_wall + mrate_floor 'kg/s
        '                WallFlowtoUpper(room, i) = -M_wall_u + (1 - frac_to_lower) * M_wall
        '                WallFlowtoLower(room, i) = -M_wall_l + frac_to_lower * M_wall
        '                FlowToLower(room, i) = Flow_to_Lower
        '                FlowToUpper(room, i) = Flow_to_Upper
        '                UFlowToOutside(room, i) = FlowOut
        '                FlowGER(room, i) = Flowtoroom
        '                ventfire(NumberRooms + 1, i) = ventfire_sum(NumberRooms + 1)
        '                ventfire(room, i) = ventfire_sum(room)
        '                HeatRelease(fireroom, i, 2) = CDbl(hrrlimit)
        '                If room = fireroom Then FuelHoC(i) = stemp 'excludes walls or ceiling
        '                If room <> fireroom Then HeatRelease(room, i, 2) = hrrnextroom
        '                xlast = X
        '            End If

        '            'mechanical ventilation
        '            Dim fanflag As Boolean = False
        '            If NumFans > 0 And mv_mode = True Then 'system is operational
        '                'can run this routine for each fan in room j
        '                Dim upperfans As Single = 0
        '                Dim lowerfans As Single = 0

        '                'fanflag = True
        '                For fx = 1 To NumFans
        '                    If fandata(fx, 0) = room Then 'room
        '                        fanflag = True
        '                        If fandata(fx, 4) = 0 Then 'manual op
        '                            Call mech_vent2(upperfans, lowerfans, fx, X, room, layerz, Y, denl, denu, species_to_upper, species_to_lower, Enthalpy_U, Enthalpy_L, cp_u, cp_l, massplume, M_wall_u, M_wall_l, fuelmassloss, Flow_to_Upper, Flow_to_Lower, dMudt, dMldt, M_wall, frac_to_lower)
        '                        ElseIf fandata(fx, 4) = 1 Then 'local sd
        '                            If SDFlag(room) = 1 Then
        '                                fandata(fx, 3) = SDTime(room) 'smoke det in room j has operated
        '                                Call mech_vent2(upperfans, lowerfans, fx, X, room, layerz, Y, denl, denu, species_to_upper, species_to_lower, Enthalpy_U, Enthalpy_L, cp_u, cp_l, massplume, M_wall_u, M_wall_l, fuelmassloss, Flow_to_Upper, Flow_to_Lower, dMudt, dMldt, M_wall, frac_to_lower)
        '                            End If
        '                        Else
        '                            If SDFlag(fireroom) = 1 Then
        '                                fandata(fx, 3) = SDTime(fireroom) 'smoke det in fireroom has operated
        '                                Call mech_vent2(upperfans, lowerfans, fx, X, room, layerz, Y, denl, denu, species_to_upper, species_to_lower, Enthalpy_U, Enthalpy_L, cp_u, cp_l, massplume, M_wall_u, M_wall_l, fuelmassloss, Flow_to_Upper, Flow_to_Lower, dMudt, dMldt, M_wall, frac_to_lower)
        '                            End If

        '                        End If
        '                    End If
        '                Next
        '                If fanflag = True Then
        '                    dMudt = massplume - M_wall_u + (1 - frac_to_lower) * M_wall + fuelmassloss + Flow_to_Upper - upperfans
        '                    dMldt = Flow_to_Lower - M_wall_l + frac_to_lower * M_wall - massplume - lowerfans
        '                End If
        '            End If
        '            If fanflag = False Then 'no mechanical ventilation
        '                dMudt = massplume + fuelmassloss + Flow_to_Upper - M_wall_u + (1 - frac_to_lower) * M_wall
        '                dMldt = Flow_to_Lower - massplume - M_wall_l + (frac_to_lower) * M_wall
        '            End If

        '            'pressure
        '            DYDX(4) = derivs_Pressure(room, Enthalpy_U, Enthalpy_L, gamma_u)

        '            'volume upper layer
        '            If TwoZones(room) = True Then
        '                DYDX(1) = derivs_UVol(DYDX(4), Y(4), Y(1), Enthalpy_U, gamma_u)
        '            Else
        '                DYDX(1) = 0
        '            End If

        '            'upper layer temperature
        '            DYDX(2) = derivs_Temp(cp_u, mw_upper, DYDX(4), Y(1), Y(2), Enthalpy_U, dMudt, Y(4))


        '            'mass used in derivs_species_U
        '            species_mass = species_mass_U(mw_upper, Y(4), Y(1), Y(2))

        '            'O2 concentration
        '            rate_Renamed = 0
        '            If room = fireroom Then
        '                rate_Renamed = O2Mass_Rate(hrrlimit)
        '            Else
        '                If ventfire_sum(room) + HeatRelease(room, stepcount, 2) > 0 Then 'unburned fuel flowing to upper layer of room j
        '                    rate_Renamed = O2Mass_Rate(hrrnextroom) 'assumes oxygen taken from upper layer
        '                End If
        '            End If
        '            DYDX(5) = derivs_species_U(species_mass, dMudt, rate_Renamed, species_to_upper(3), Y(5), Y(15), massplume)

        '            'H2O concentration
        '            rate_Renamed = 0
        '            If room = fireroom Then
        '                If Flashover = True And frmOptions1.optPostFlashover.Checked = True Then
        '                    rate_Renamed = species_mass_rate_postflashover(GER, room, H2OYieldPF, WallH2OYield, CeilingH2OYield, FloorH2OYield, X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
        '                Else
        '                    If GER > 1 Then
        '                        rate_Renamed = species_mass_rate2(GER, room, WaterVaporYield, WallH2OYield, CeilingH2OYield, FloorH2OYield, X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
        '                    Else
        '                        rate_Renamed = species_mass_rate(room, WaterVaporYield, WallH2OYield, CeilingH2OYield, FloorH2OYield, X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
        '                    End If
        '                End If
        '                If fuelmassloss > 0 Then rate_Renamed = rate_Renamed * (fuelmassloss - genrate) / fuelmassloss
        '            Else
        '                rate_Renamed = species_mass_rate(room, WaterVaporYield, WallH2OYield, CeilingH2OYield, FloorH2OYield, X, mrate, 0, fuelmassloss, 0)
        '                If fuelmassloss > 0 Then rate_Renamed = rate_Renamed * (fuelmassloss - genrate) / fuelmassloss
        '            End If
        '            DYDX(16) = derivs_species_U(species_mass, dMudt, rate_Renamed, species_to_upper(7), Y(16), Y(17), massplume)

        '            'Soot Concentration
        '            rate_Renamed = 0
        '            If room = fireroom Then
        '                If Flashover = True And frmOptions1.optPostFlashover.Checked = True Then
        '                    'using postflashover model
        '                    'lets adjust the soot yield for underventilated conditions
        '                    rate_Renamed = soot_mass_rate2(room, GER, SootYieldPF, WallSootYield, CeilingSootYield, FloorSootYield, X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
        '                Else
        '                    'lets adjust the soot yield for underventilated conditions
        '                    rate_Renamed = soot_mass_rate(room, GER, SootYield, WallSootYield, CeilingSootYield, FloorSootYield, X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
        '                End If
        '                If X = tim(i, 1) And xflag = 0 Then
        '                    SootMassGen(i) = rate_Renamed 'kg-soot/s in fireroom
        '                End If
        '            Else
        '                rate_Renamed = species_mass_rate(room, SootYield, WallSootYield, CeilingSootYield, FloorSootYield, X, mrate, 0, fuelmassloss(j), 0)
        '                If fuelmassloss > 0 Then rate_Renamed = rate_Renamed * (fuelmassloss) / fuelmassloss
        '            End If
        '            DYDX(12) = derivs_species_U(species_mass, dMudt, rate_Renamed, species_to_upper(6), Y(12), Y(13), massplume)

        '            'HCN Concentration
        '            rate_Renamed = 0
        '            If room = fireroom Then
        '                If GER > 1 Then
        '                    rate_Renamed = species_mass_rate2(GER, room, HCNYield, WallHCNYield, CeilingHCNYield, FloorHCNYield, X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
        '                Else
        '                    rate_Renamed = species_mass_rate(room, HCNYield, WallHCNYield, CeilingHCNYield, FloorHCNYield, X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
        '                End If
        '                If fuelmassloss > 0 Then rate_Renamed = rate_Renamed * (fuelmassloss - genrate) / fuelmassloss
        '            Else
        '                rate_Renamed = species_mass_rate(room, HCNYield, WallHCNYield, CeilingHCNYield, FloorHCNYield, X, mrate, 0, fuelmassloss, 0)
        '                If fuelmassloss > 0 Then rate_Renamed = rate_Renamed * (fuelmassloss - genrate) / fuelmassloss
        '            End If
        '            DYDX(18) = derivs_species_U(species_mass, dMudt, rate_Renamed, species_to_upper(9), Y(18), Y(14), massplume)

        '            'TUHC concentration
        '            rate_Renamed = 0
        '            If room = fireroom Then
        '                'rate at which unburned fuel generated in fire room
        '                rate_Renamed = TUHC_Rate(hrrlimit, fuelmassloss, weighted_hc)
        '            Else
        '                'rate at which unburned fuel consumed in adjacent rooms
        '                If weighted_hc > 0 Then
        '                    rate_Renamed = -ventfire_sum(room) / (weighted_hc * 1000) + TUHC_Rate(hrrnextroom - ventfire_sum(room), fuelmassloss, weighted_hc)
        '                End If
        '            End If
        '            DYDX(6) = derivs_species_U(species_mass, dMudt, rate_Renamed, species_to_upper(8), Y(6), Y(7), massplume)
        '            genrate = rate_Renamed
        '            If genrate < 0 Then genrate = 0

        '            'CO concentration
        '            rate_Renamed = 0
        '            'uses combusted mass
        '            If room = fireroom Then rate_Renamed = COMass_Rate2(GER, X, fuelmassloss - genrate, massplume, Y(2))
        '            DYDX(8) = derivs_species_U(species_mass, dMudt, rate_Renamed, species_to_upper(4), Y(8), Y(9), massplume)

        '            'CO2 concentration
        '            rate_Renamed = 0
        '            If room = fireroom Then
        '                If Flashover = True And frmOptions1.optPostFlashover.Checked = True Then
        '                    rate_Renamed = species_mass_rate_postflashover(GER, room, CO2YieldPF, WallCO2Yield, CeilingCO2Yield, FloorCO2Yield, X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
        '                Else
        '                    If GER > 1 Then 'adjust co2 yields for underventilated fires
        '                        rate_Renamed = species_mass_rate2(GER, room, CO2Yield, WallCO2Yield, CeilingCO2Yield, FloorCO2Yield, X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
        '                    Else
        '                        rate_Renamed = species_mass_rate(room, CO2Yield, WallCO2Yield, CeilingCO2Yield, FloorCO2Yield, X, mrate, mrate_wall, mrate_ceiling, mrate_floor)
        '                    End If
        '                End If
        '                If fuelmassloss > 0 Then rate_Renamed = rate_Renamed * (fuelmassloss - genrate) / fuelmassloss
        '            Else
        '                rate_Renamed = species_mass_rate(room, CO2Yield, WallCO2Yield, CeilingCO2Yield, FloorCO2Yield, X, mrate, 0, fuelmassloss, 0)
        '                If fuelmassloss > 0 Then rate_Renamed = rate_Renamed * (fuelmassloss - genrate) / fuelmassloss
        '            End If
        '            DYDX(10) = derivs_species_U(species_mass, dMudt, rate_Renamed, species_to_upper(5), Y(10), Y(11), massplume)

        '            'lower layer temperature
        '            DYDX(3) = derivs_Temp(cp_l, mw_lower, DYDX(4), RoomVolume(room) - Y(1), Y(3), Enthalpy_L, dMldt, Y(4))

        '            'mass used in derivs_species_L
        '            species_mass = species_mass_L(mw_lower, Y(4), Y(3), RoomVolume(room) - Y(1))

        '            'TUHC
        '            DYDX(7) = derivs_species_L(species_mass, dMldt, species_to_lower(8), Y(7), Y(6), massplume)

        '            'lower layer
        '            'CO concentration
        '            DYDX(9) = derivs_species_L(species_mass, dMldt, species_to_lower(4), Y(9), Y(8), massplume)

        '            'CO2 concentration
        '            DYDX(11) = derivs_species_L(species_mass, dMldt, species_to_lower(5), Y(11), Y(10), massplume)

        '            'Soot Concentration
        '            DYDX(13) = derivs_species_L(species_mass, dMldt, species_to_lower(6), Y(13), Y(12), massplume)

        '            'O2 concentration
        '            DYDX(15) = derivs_species_L(species_mass, dMldt, species_to_lower(3), Y(15), Y(5), massplume)

        '            'H2O concentration
        '            DYDX(17) = derivs_species_L(species_mass, dMldt, species_to_lower(7), Y(17), Y(16), massplume)

        '            'HCN concentration
        '            DYDX(14) = derivs_species_L(species_mass, dMldt, species_to_lower(9), Y(14), Y(18), massplume)

        '            'do some other stuff if you want
        '            Erase mrate
        '            Erase Enthalpy
        '            Erase ventflow
        '            Erase products
        '            Erase ventfire_sum
        '            Erase matc
        '            Erase species_to_upper
        '            Erase species_to_lower
        '            Erase vententrain
        '            System.Windows.Forms.Application.DoEvents()
        '            If flagstop = 1 Then Exit Sub
        '            gTimeX = X

        '            Exit Sub

        'errorhandler:
        '            response = MsgBox("Error in DIFF_EQNS. Do you want to continue?", MsgBoxStyle.YesNo)

        '            If response = MsgBoxResult.No Then flagstop = 1
        '            MsgBox(ErrorToString())
        '            Err.Clear()
        '            Exit Sub
        '        End Sub

        '        Private Sub ventflows(ByRef Mass_Upper() As Double, ByRef vententrain(,,) As Double, ByVal tim As Double, ByRef Y(,) As Double, ByRef ventflow(,) As Double, ByRef layerz() As Double, ByRef Enthalpy(,) As Double, ByRef products(,,) As Double, ByRef ventfire_sum() As Double, ByRef mw_upper() As Double, ByRef mw_lower() As Double, ByVal weighted_hc As Double, ByRef FlowOut() As Double, ByRef Flowtoroom() As Double, ByRef cp_u() As Double, ByRef cp_l() As Double)
        '            '====== ===============================================================
        '            '   VENT FLOW ROUTINE BASED ON PRESSURE
        '            '   C A WADE 21/8/98
        '            ' 1 feb 2008 Amanda
        '            '=====================================================================
        '            Dim DischargeCoeff As Double
        '            Dim SprinkCoolCoeff As Double
        '            Dim NProd, D, k, i, h, j, L, nelev, mxprd As Integer
        '            Dim xmslab() As Double
        '            Dim dpv1m2(9) As Double
        '            Dim yvelev(9) As Double
        '            Dim nvelev As Integer
        '            Dim nslab As Integer
        '            Dim yflor(2) As Double
        '            Dim yceil(2) As Double
        '            Dim denu(2) As Double
        '            Dim denl(2) As Double
        '            Dim r1u As Double
        '            Dim r1l As Double
        '            Dim ylay(2) As Double
        '            Dim pflor(2) As Double
        '            Dim TU(2) As Double
        '            Dim TL(2) As Double
        '            Dim r2u As Double
        '            Dim r2l As Double
        '            Dim yelev(9) As Double
        '            Dim dp1m2(9) As Double
        '            Dim pab1(9) As Double
        '            Dim pab2(9) As Double
        '            Dim yvbot, epsp, yvtop, avent As Double
        '            Dim conl(,) As Double
        '            Dim conu(,) As Double
        '            Dim cslab(,) As Double
        '            Dim pslab(,) As Double
        '            Dim qslab() As Double
        '            Dim yslab() As Double
        '            Dim dirs12() As Double
        '            Dim m, mxslab As Integer
        '            Dim dteps, rlam As Double
        '            Dim tslab() As Double
        '            Dim UFLW2(,,) As Double
        '            Dim QSUVNT() As Double
        '            Dim QSLVNT() As Double
        '            Dim e2u, e1u, neutralplane, e1l, e2l As Double
        '            Dim r2utemp, r1utemp, r1ltemp, r2ltemp As Double
        '            Dim prod1U() As Double
        '            Dim prod2U() As Double
        '            Dim prod1L() As Double
        '            Dim prod2L() As Double
        '            Dim term, vp3, vp1, vp, vp2, vent_entrain, zp As Double
        '            Dim doormix(2) As Double
        '            Dim YVent, ventfire, FU As Double
        '            Dim cpu(2) As Double
        '            Dim cpl(2) As Double
        '            Dim uvol(2) As Double
        '            Dim cvent_Renamed(,) As Double
        '            Dim C(,) As Double
        '            Dim PL(,) As Double
        '            Dim PU(,) As Double
        '            Dim PVent(,) As Double
        '            Dim P() As Double
        '            Dim XML(2) As Double
        '            Dim XMU(2) As Double
        '            Dim XMVent(2) As Double
        '            Dim QL(2) As Double
        '            Dim QU(2) As Double
        '            Dim QVent(2) As Double
        '            Dim T(2) As Double
        '            Dim Tvent(2) As Double
        '            Dim DELP1 As Double
        '            Dim Den(2) As Double

        '            Static tlast As Double
        '            Static FirstTime, FirstTime2 As Boolean

        '            On Error GoTo errorhandler

        '            k = stepcount

        '            If gb_first_time_vent = True Then 'If tim = 0 Then
        '                FirstTime = True
        '                FirstTime2 = True
        '            End If
        '            NProd = 7 'number of species
        '            mxprd = 7 'max number of products
        '            epsp = Error_Control_ventflow
        '            'epsp = 0.001
        '            'error tolerance
        '            mxslab = 6
        '            dteps = 3 'K difference in relative temps before layer will form
        '            rlam = 0

        '            ReDim conl(mxprd, NumberRooms + 1)
        '            ReDim conu(mxprd, NumberRooms + 1)
        '            ReDim cslab(mxslab, mxprd)
        '            ReDim pslab(mxslab, mxprd)
        '            ReDim xmslab(mxslab)
        '            ReDim yslab(mxslab)
        '            ReDim qslab(mxslab)
        '            ReDim UFLW2(2, NProd + 2, 2)
        '            ReDim tslab(mxslab)
        '            ReDim dirs12(mxslab)
        '            ReDim QSUVNT(2)
        '            ReDim QSLVNT(2)
        '            ReDim ventflow(NumberRooms + 1, 2)
        '            ReDim Enthalpy(NumberRooms + 1, 2)
        '            ReDim products(NProd + 2, NumberRooms + 1, 2)
        '            ReDim prod1U(NProd + 2)
        '            ReDim prod2U(NProd + 2)
        '            ReDim prod1L(NProd + 2)
        '            ReDim prod2L(NProd + 2)
        '            ReDim cvent_Renamed(NProd, 2)
        '            ReDim C(NProd, 2)
        '            ReDim PL(NProd, 2)
        '            ReDim PU(NProd, 2)
        '            ReDim PVent(NProd, 2)
        '            ReDim P(2)

        '            For i = 1 To NumberRooms + 1 'first room
        '                For j = 1 To NumberRooms + 1 'second room
        '                    If i <> j Then
        '                        r1u = 0
        '                        r1l = 0
        '                        r2u = 0
        '                        r2l = 0
        '                        If i < NumberRooms + 1 And j > 1 Then 'i is an inside room
        '                            If i < j Then
        '                                If NumberVents(i, j) > 0 Then 'horizontal vents exist
        '                                    'setup routines
        '                                    uvol(1) = Y(i, 1)
        '                                    yflor(1) = FloorElevation(i)
        '                                    yceil(1) = RoomHeight(i) + FloorElevation(i)
        '                                    ylay(1) = layerz(i) + FloorElevation(i)
        '                                    denu(1) = (Atm_Pressure + Y(i, 4) / 1000) / (Gas_Constant / mw_upper(i)) / Y(i, 2)
        '                                    denl(1) = (Atm_Pressure + Y(i, 4) / 1000) / (Gas_Constant / mw_lower(i)) / Y(i, 3)
        '                                    pflor(1) = Y(i, 4)
        '                                    TU(1) = Y(i, 2)
        '                                    TL(1) = Y(i, 3)
        '                                    cpu(1) = cp_u(i)
        '                                    cpl(1) = cp_l(i)
        '                                    conl(1, 1) = Y(i, 15) 'lower oxygen mass fraction
        '                                    conu(1, 1) = Y(i, 5) 'upper oxygen
        '                                    conl(2, 1) = Y(i, 9) 'lower co
        '                                    conu(2, 1) = Y(i, 8) 'upper co
        '                                    conl(3, 1) = Y(i, 11) 'lower co2
        '                                    conu(3, 1) = Y(i, 10) 'upper co2
        '                                    conl(4, 1) = Y(i, 13) 'lower soot
        '                                    conu(4, 1) = Y(i, 12) 'upper soot
        '                                    conl(5, 1) = Y(i, 17) 'lower water
        '                                    conu(5, 1) = Y(i, 16) 'upper water
        '                                    conl(6, 1) = Y(i, 7) 'lower tuhc
        '                                    conu(6, 1) = Y(i, 6) 'upper tuhc
        '                                    conl(7, 1) = Y(i, 14) 'lower hcn
        '                                    conu(7, 1) = Y(i, 18) 'upper hcn

        '                                    If j = NumberRooms + 1 Then 'outside
        '                                        yflor(2) = yflor(1)
        '                                        yceil(2) = yceil(1)
        '                                        ylay(2) = yceil(1)
        '                                        pflor(2) = RoomPressure(i, 1)
        '                                        Call Outside_Conditions(conl, conu, TU, TL, cpu, cpl, pflor, denu, denl, 2)
        '                                    Else 'j is an inside room
        '                                        uvol(2) = Y(j, 1)
        '                                        yflor(2) = FloorElevation(j)
        '                                        yceil(2) = RoomHeight(j) + FloorElevation(j)
        '                                        ylay(2) = layerz(j) + FloorElevation(j)
        '                                        denu(2) = (Atm_Pressure + Y(j, 4) / 1000) / (Gas_Constant / mw_upper(j)) / Y(j, 2)
        '                                        denl(2) = (Atm_Pressure + Y(j, 4) / 1000) / (Gas_Constant / mw_lower(j)) / Y(j, 3)
        '                                        pflor(2) = Y(j, 4)
        '                                        TU(2) = Y(j, 2)
        '                                        TL(2) = Y(j, 3)
        '                                        cpu(2) = cp_u(j)
        '                                        cpl(2) = cp_l(j)
        '                                        conl(1, 2) = Y(j, 15) 'lower oxygen
        '                                        conu(1, 2) = Y(j, 5) 'upper oxygen
        '                                        conl(2, 2) = Y(j, 9) 'lower co
        '                                        conu(2, 2) = Y(j, 8) 'upper co
        '                                        conl(3, 2) = Y(j, 11) 'lower co2
        '                                        conu(3, 2) = Y(j, 10) 'upper co2
        '                                        conl(4, 2) = Y(j, 13) 'lower soot
        '                                        conu(4, 2) = Y(j, 12) 'upper soot
        '                                        conl(5, 2) = Y(j, 17) 'lower water
        '                                        conu(5, 2) = Y(j, 16) 'upper water
        '                                        conl(6, 2) = Y(j, 7) 'lower tuhc
        '                                        conu(6, 2) = Y(j, 6) 'upper tuhc
        '                                        conl(7, 2) = Y(j, 14) 'lower hcn
        '                                        conu(7, 2) = Y(j, 18) 'upper hcn
        '                                    End If

        '                                    Call CommonWall(yflor, yceil, ylay, denu, denl, pflor, nelev, yelev, dp1m2, pab1, pab2)

        '                                    'which vent
        '                                    r1u = 0
        '                                    r1l = 0
        '                                    r2u = 0
        '                                    r2l = 0
        '                                    e1u = 0
        '                                    e1l = 0
        '                                    e2u = 0
        '                                    e2l = 0
        '                                    For h = 3 To NProd + 2
        '                                        prod1U(h) = 0
        '                                        prod1L(h) = 0
        '                                        prod2U(h) = 0
        '                                        prod2L(h) = 0
        '                                    Next h
        '                                    avent = 0

        '                                    'considering vertical vents
        '                                    For m = 1 To NumberVents(i, j)

        '                                        avent = ventarea(tim, i, j, m) 'determines vent area and if vent closed sets area to zero
        '                                        If avent > gcd_Machine_Error Then 'vent is open between the rooms
        '                                            yvtop = VentHeight(i, j, m) + VentSillHeight(i, j, m) + FloorElevation(i)
        '                                            yvbot = VentSillHeight(i, j, m) + FloorElevation(i)
        '                                            'If m = 2 Then Stop
        '                                            'sprinkler cooling coefficient
        '                                            If SprinklerFlag = 1 And i = fireroom Then
        '                                                'SprinkCoolCoeff = gcs_SprinkCoolCoeff 'sprinkler has operated
        '                                                SprinkCoolCoeff = SprCooling 'sprinkler has operated
        '                                            Else
        '                                                SprinkCoolCoeff = 1.0 'no effect modelled
        '                                            End If

        '                                            DischargeCoeff = VentCD(i, j, m) * SprinkCoolCoeff

        '                                            If DischargeCoeff > 0 Then

        '                                                Call VENTHP(DischargeCoeff, yflor, ylay, TU, TL, cpu, cpl, denl, denu, pflor, yvtop, yvbot, avent, yelev, dp1m2, nelev, pab1, pab2, conl, conu, NProd, epsp, cslab, pslab, tslab, qslab, yslab, dirs12, dpv1m2, yvelev, xmslab, nvelev, nslab)

        '                                                If frmOptions1.optCCFM.Checked = True Then
        '                                                    'use vent flow depostion rules from CCFM
        '                                                    Call FLogo2(uvol, WAfloor_flag(i, j, m), dirs12, NProd, yslab, yslab, xmslab, tslab, nslab, TU, TL, yflor, yceil, ylay, qslab, rlam, pslab, mxprd, mxslab, dteps, QSLVNT, QSUVNT, UFLW2)
        '                                                Else
        '                                                    'use vent flow deposition rules from CFAST 3.17
        '                                                    Call FLogo1(i, j, uvol, WAfloor_flag(i, j, m), dirs12, NProd, yslab, yslab, xmslab, tslab, nslab, TU, TL, yflor, yceil, ylay, qslab, rlam, pslab, mxprd, mxslab, dteps, QSLVNT, QSUVNT, UFLW2)
        '                                                    'Call FLogo2(uvol, WAfloor_flag(i, j, m), dirs12, NProd, yslab, yslab, xmslab, tslab, nslab, TU, TL, yflor, yceil, ylay, qslab, rlam, pslab, mxprd, mxslab, dteps, QSLVNT, QSUVNT, UFLW2)
        '                                                End If

        '                                                r1u = r1u + UFLW2(1, 1, 2)
        '                                                r1l = r1l + UFLW2(1, 1, 1)
        '                                                r2u = r2u + UFLW2(2, 1, 2)
        '                                                r2l = r2l + UFLW2(2, 1, 1)
        '                                                e1u = e1u + UFLW2(1, 2, 2)
        '                                                e1l = e1l + UFLW2(1, 2, 1)
        '                                                e2u = e2u + UFLW2(2, 2, 2)
        '                                                e2l = e2l + UFLW2(2, 2, 1)
        '                                                For h = 3 To NProd + 2
        '                                                    prod1U(h) = prod1U(h) + UFLW2(1, h, 2) 'net flow of species kg/s
        '                                                    prod1L(h) = prod1L(h) + UFLW2(1, h, 1)
        '                                                    prod2U(h) = prod2U(h) + UFLW2(2, h, 2)
        '                                                    prod2L(h) = prod2L(h) + UFLW2(2, h, 1)
        '                                                Next h

        '                                                'keep track of the volumetric vent flow from each room to the outside, excluding vent mixing terms
        '                                                If r1u < 0 Then FlowOut(i) = FlowOut(i) - r1u / denu(1)

        '                                                ventfire = 0
        '                                                'Call door_mixing_flow(NProd, nelev, m, i, j, conu, prod2U, prod2L, prod1U, prod1L, TU, TL, cpu, cpl, e1u, e1l, e2u, e2l, r1u, r1l, r2l, r2u, ylay, UFLW2, dp1m2, yelev)

        '                                                'new - Utiskul and Quintiere
        '                                                Call near_vent_mixing2(NProd, nelev, m, i, j, conu, prod2U, prod2L, prod1U, prod1L, TU, TL, cpu, cpl, e1u, e1l, e2u, e2l, r1u, r1l, r2l, r2u, ylay, UFLW2, dp1m2, yelev)

        '                                                'i=first room j=second room, m=vent id
        '                                                Call door_entrain4(dteps, Mass_Upper, vententrain, m, i, NProd, prod1L, prod2L, prod2U, prod1U, conu, conl, TU, TL, cpu, cpl, j, e1l, e1u, e2l, e2u, r2u, r2l, r1u, r1l, ylay, UFLW2, ventfire, mw_upper, mw_lower, weighted_hc, nelev, dp1m2, yelev)

        '                                                'write to a vent flow log
        '                                                If ventlog = True And (tim > tlast Or FirstTime = True) Then
        '                                                    If tim - Int(tim) = 0 Then
        '                                                        Call Ventflow_log2(FirstTime, tim, i, j, m, nslab, xmslab, dirs12, yvelev, vententrain)
        '                                                    End If
        '                                                End If

        '                                                'keep track of flow to the fire room, for use in GER calculation
        '                                                If i = fireroom Then
        '                                                    For D = 1 To nslab
        '                                                        If dirs12(D) < 0 Then
        '                                                            Flowtoroom(i) = xmslab(D)
        '                                                        End If
        '                                                    Next
        '                                                End If

        '                                                If r1u < 0 Then 'ventfire burning in room 2
        '                                                    ventfire_sum(j) = ventfire_sum(j) + ventfire
        '                                                Else 'ventfire burning in room 1
        '                                                    ventfire_sum(i) = ventfire_sum(i) + ventfire
        '                                                End If
        '                                            End If
        '                                        End If
        '                                    Next m ' moved this down 24/10/2003

        '                                    ventflow(i, 1) = ventflow(i, 1) + r1l '1st room lower layer
        '                                    ventflow(i, 2) = ventflow(i, 2) + r1u '1st room upper layer
        '                                    ventflow(j, 1) = ventflow(j, 1) + r2l '2nd room lower layer
        '                                    ventflow(j, 2) = ventflow(j, 2) + r2u '2nd room upper layer
        '                                    Enthalpy(i, 1) = Enthalpy(i, 1) + e1l '1st room lower layer
        '                                    Enthalpy(i, 2) = Enthalpy(i, 2) + e1u '1st room upper layer
        '                                    Enthalpy(j, 1) = Enthalpy(j, 1) + e2l '2nd room lower layer
        '                                    Enthalpy(j, 2) = Enthalpy(j, 2) + e2u '2nd room upper layer
        '                                    For h = 3 To NProd + 2
        '                                        products(h, i, 1) = products(h, i, 1) + prod1L(h) '1st room lower layer
        '                                        products(h, i, 2) = products(h, i, 2) + prod1U(h) '1st room upper layer
        '                                        products(h, j, 1) = products(h, j, 1) + prod2L(h) '2nd room lower layer
        '                                        products(h, j, 2) = products(h, j, 2) + prod2U(h) '2nd room upper layer
        '                                    Next h
        '                                End If
        '                            End If
        '                        End If

        '                        'now do the ceiling/floor vents
        '                        ' i is room above vent, j is room below vent
        '                        If NumberCVents(i, j) > 0 Then
        '                            avent = 0
        '                            r1u = 0
        '                            r1l = 0
        '                            r2u = 0
        '                            r2l = 0
        '                            e1u = 0
        '                            e1l = 0
        '                            e2u = 0
        '                            e2l = 0
        '                            For h = 3 To NProd + 2
        '                                prod1U(h) = 0
        '                                prod1L(h) = 0
        '                                prod2U(h) = 0
        '                                prod2L(h) = 0
        '                            Next h

        '                            'upper space conditions
        '                            If i <> NumberRooms + 1 Then 'inside room
        '                                conl(1, 1) = Y(i, 15) 'lower oxygen mass fraction
        '                                conu(1, 1) = Y(i, 5) 'upper oxygen
        '                                conl(2, 1) = Y(i, 9) 'lower co
        '                                conu(2, 1) = Y(i, 8) 'upper co
        '                                conl(3, 1) = Y(i, 11) 'lower co2
        '                                conu(3, 1) = Y(i, 10) 'upper co2
        '                                conl(4, 1) = Y(i, 13) 'lower soot
        '                                conu(4, 1) = Y(i, 12) 'upper soot
        '                                conl(5, 1) = Y(i, 17) 'lower water
        '                                conu(5, 1) = Y(i, 16) 'upper water
        '                                conl(6, 1) = Y(i, 7) 'lower tuhc
        '                                conu(6, 1) = Y(i, 6) 'upper tuhc
        '                                conl(7, 1) = Y(i, 14) 'lower hcn
        '                                conu(7, 1) = Y(i, 18) 'upper hcn
        '                                TU(1) = Y(i, 2) 'upper layer temp
        '                                TL(1) = Y(i, 3) 'lower layer temp
        '                                cpu(1) = cp_u(i)
        '                                cpl(1) = cp_l(i)
        '                                uvol(1) = Y(i, 1)
        '                                yflor(1) = FloorElevation(i)
        '                                yceil(1) = RoomHeight(i) + FloorElevation(i)
        '                                ylay(1) = layerz(i) + FloorElevation(i)
        '                                pflor(1) = Y(i, 4)
        '                                denu(1) = (Atm_Pressure + Y(i, 4) / 1000) / (Gas_Constant / mw_upper(i)) / Y(i, 2)
        '                                denl(1) = (Atm_Pressure + Y(i, 4) / 1000) / (Gas_Constant / mw_lower(i)) / Y(i, 3)
        '                            Else 'outside space
        '                                Call Outside_Conditions(conl, conu, TU, TL, cpu, cpl, pflor, denu, denl, 1)
        '                                yflor(1) = FloorElevation(j)
        '                                yceil(1) = yflor(1)
        '                                ylay(1) = yceil(1)
        '                                pflor(1) = -G * yflor(1) * (Atm_Pressure) / (Gas_Constant / MW_air) / ExteriorTemp
        '                            End If

        '                            'lower space conditions
        '                            If j <> NumberRooms + 1 Then 'inside room
        '                                conl(1, 2) = Y(j, 15) 'lower oxygen mass fraction
        '                                conu(1, 2) = Y(j, 5) 'upper oxygen
        '                                conl(2, 2) = Y(j, 9) 'lower co
        '                                conu(2, 2) = Y(j, 8) 'upper co
        '                                conl(3, 2) = Y(j, 11) 'lower co2
        '                                conu(3, 2) = Y(j, 10) 'upper co2
        '                                conl(4, 2) = Y(j, 13) 'lower soot
        '                                conu(4, 2) = Y(j, 12) 'upper soot
        '                                conl(5, 2) = Y(j, 17) 'lower water
        '                                conu(5, 2) = Y(j, 16) 'upper water
        '                                conl(6, 2) = Y(j, 7) 'lower tuhc
        '                                conu(6, 2) = Y(j, 6) 'upper tuhc
        '                                conl(7, 2) = Y(j, 14) 'lower hcn
        '                                conu(7, 2) = Y(j, 18) 'upper hcn
        '                                TU(2) = Y(j, 2) 'upper layer temp
        '                                TL(2) = Y(j, 3) 'lower layer temp
        '                                cpu(2) = cp_u(j)
        '                                cpl(2) = cp_l(j)
        '                                uvol(2) = Y(j, 1)
        '                                yflor(2) = FloorElevation(j)
        '                                yceil(2) = RoomHeight(j) + FloorElevation(j)
        '                                ylay(2) = layerz(j) + FloorElevation(j)
        '                                pflor(2) = Y(j, 4)
        '                                denu(2) = (Atm_Pressure + Y(j, 4) / 1000) / (Gas_Constant / mw_upper(j)) / Y(j, 2)
        '                                denl(2) = (Atm_Pressure + Y(j, 4) / 1000) / (Gas_Constant / mw_lower(j)) / Y(j, 3)
        '                            Else 'outside space
        '                                Call Outside_Conditions(conl, conu, TU, TL, cpu, cpl, pflor, denu, denl, 2)
        '                                yflor(2) = FloorElevation(i)
        '                                ylay(2) = yflor(2)
        '                                yceil(2) = yflor(2)
        '                                pflor(2) = -G * yflor(2) * (Atm_Pressure) / (Gas_Constant / MW_air) / ExteriorTemp
        '                            End If

        '                            If gb_first_time_vent = True Then 'only perform this the first time through ventflows
        '                                'Call check_vent_areas_C(i, j) 'check for vent area compared to room floor area once per combination of rooms
        '                                Call check_cvent_areas(i, j) 'colleen 2/6/09 fixed prob in 2009.1 for multpile ceiling vents
        '                            End If

        '                            For m = 1 To NumberCVents(i, j)
        '                                r1ltemp = 0
        '                                r1utemp = 0
        '                                r2utemp = 0
        '                                r2ltemp = 0

        '                                Call ventareasC(tim, i, j, m, avent) 'calculates vent area, 'avent', for horizontal/ceiling/floor vents

        '                                If avent > gcd_Machine_Error Then 'vent is open
        '                                    'determine the elevation of the horizontal vent
        '                                    If j = NumberRooms + 1 Then 'lower room is an outside space
        '                                        YVent = FloorElevation(i) 'vent is at the floor height of the upper room, i
        '                                    Else 'lower room is an inside space
        '                                        YVent = FloorElevation(j) + RoomHeight(j) 'vent is at the ceiling height of the lower room, j
        '                                    End If
        '                                    If j <> NumberRooms + 1 And i <> NumberRooms + 1 Then
        '                                        If FloorElevation(i) + gcd_Machine_Error < FloorElevation(j) + RoomHeight(j) Then 'floor of upper room, i, is lower than ceiling of lower floor, j
        '                                            MsgBox("The horizontal vent configuration may be invalid. Please check.")
        '                                            flagstop = 1
        '                                        End If
        '                                    End If

        '                                    DischargeCoeff = CVentDC(i, j, m)

        '                                    Call VENTCF2A(avent, conl, conu, NProd, TU, TL, cpu, cpl, yflor, yceil, ylay, YVent, pflor, epsp, denl, denu, C, cvent_Renamed, XML, XMU, XMVent, PL, PU, PVent, P, QL, QU, QVent, T, Tvent, DELP1, Den, DischargeCoeff)

        '                                    'flow taken from lower room (2) and put in upper room (1)
        '                                    Call CV_Flogo(dteps, ylay(1), yceil(1), yflor(1), Tvent(1), TU(1), TL(1), FU)
        '                                    r1ltemp = XMVent(1) * (1 - FU)
        '                                    r1utemp = XMVent(1) * FU
        '                                    r2utemp = XMU(2)
        '                                    r2ltemp = XML(2)

        '                                    r1l = r1l + XMVent(1) * (1 - FU)
        '                                    r1u = r1u + XMVent(1) * FU
        '                                    r2u = r2u + XMU(2)
        '                                    r2l = r2l + XML(2)
        '                                    e1l = e1l + QVent(1) * (1 - FU)
        '                                    e1u = e1u + QVent(1) * FU
        '                                    e2u = e2u + QU(2)
        '                                    e2l = e2l + QL(2)
        '                                    For h = 3 To NProd + 2
        '                                        prod1L(h) = prod1L(h) + PVent(h - 2, 1) * (1 - FU)
        '                                        prod1U(h) = prod1U(h) + PVent(h - 2, 1) * FU
        '                                        prod2U(h) = prod2U(h) + PU(h - 2, 2)
        '                                        prod2L(h) = prod2L(h) + PL(h - 2, 2)
        '                                    Next h

        '                                    'flow taken from upper room (1) and put in lower room (2)
        '                                    Call CV_Flogo(0, ylay(2), yceil(2), yflor(2), Tvent(2), TU(2), TL(2), FU)
        '                                    r2ltemp = r2ltemp + XMVent(2) * (1 - FU)
        '                                    r2utemp = r2utemp + XMVent(2) * FU
        '                                    r1utemp = r1utemp + XMU(1)
        '                                    r1ltemp = r1ltemp + XML(1)

        '                                    r2l = r2l + XMVent(2) * (1 - FU)
        '                                    r2u = r2u + XMVent(2) * FU
        '                                    r1u = r1u + XMU(1)
        '                                    r1l = r1l + XML(1)
        '                                    e2l = e2l + QVent(2) * (1 - FU)
        '                                    e2u = e2u + QVent(2) * FU
        '                                    e1u = e1u + QU(1)
        '                                    e1l = e1l + QL(1)

        '                                    For h = 3 To NProd + 2
        '                                        prod2L(h) = prod2L(h) + PVent(h - 2, 2) * (1 - FU)
        '                                        prod2U(h) = prod2U(h) + PVent(h - 2, 2) * FU
        '                                        prod1U(h) = prod1U(h) + PU(h - 2, 1)
        '                                        prod1L(h) = prod1L(h) + PL(h - 2, 1)
        '                                    Next h

        '                                    'If frmOptions1.chkSaveVentFlow.Value = vbChecked And (tim > tlast Or tim = 0) Then
        '                                    If frmInputs.chkSaveVentFlow.CheckState = System.Windows.Forms.CheckState.Checked And (tim > tlast Or FirstTime2 = True) Then
        '                                        If tim - Int(tim) = 0 Then
        '                                            Call Ventflow_log3(FirstTime2, tim, i, j, m, r1utemp, r1ltemp, r2utemp, r2ltemp, XMVent)
        '                                        End If
        '                                    End If

        '                                    'keep track of flow to the fire room, for use in GER calculation
        '                                    If i = fireroom Then
        '                                        If r1u > 0 Then Flowtoroom(i) = Flowtoroom(i) + r1u
        '                                        If r1l > 0 Then Flowtoroom(i) = Flowtoroom(i) + r1l
        '                                    ElseIf j = fireroom Then
        '                                        If r2u > 0 Then Flowtoroom(j) = Flowtoroom(j) + r2u
        '                                        If r2l > 0 Then Flowtoroom(j) = Flowtoroom(j) + r2l
        '                                    End If

        '                                End If
        '                                'If m = NumberVents(i, j) Then gb_first_time_vent = False 'cw 18/3/2010
        '                            Next m
        '                            ventflow(i, 1) = ventflow(i, 1) + r1l 'upper room lower layer
        '                            ventflow(i, 2) = ventflow(i, 2) + r1u 'upper room upper layer
        '                            ventflow(j, 1) = ventflow(j, 1) + r2l 'lower room lower layer
        '                            ventflow(j, 2) = ventflow(j, 2) + r2u 'lower room upper layer
        '                            Enthalpy(i, 1) = Enthalpy(i, 1) + e1l 'upper room lower layer
        '                            Enthalpy(i, 2) = Enthalpy(i, 2) + e1u 'upper room upper layer
        '                            Enthalpy(j, 1) = Enthalpy(j, 1) + e2l 'lower room lower layer
        '                            Enthalpy(j, 2) = Enthalpy(j, 2) + e2u 'lower room upper layer
        '                            For h = 3 To NProd + 2
        '                                products(h, i, 1) = products(h, i, 1) + prod1L(h) '1st room lower layer
        '                                products(h, i, 2) = products(h, i, 2) + prod1U(h) '1st room upper layer
        '                                products(h, j, 1) = products(h, j, 1) + prod2L(h) '2nd room lower layer
        '                                products(h, j, 2) = products(h, j, 2) + prod2U(h) '2nd room upper layer
        '                            Next h
        '                        End If
        '                        If i = NumberRooms + 1 And r2u < 0 Then 'has to be a horizontal vent
        '                            FlowOut(j) = FlowOut(j) + -r2u / denu(2)
        '                        End If

        '                    End If
        '                Next j 'next room
        '            Next i 'next room
        '            tlast = tim

        '            gb_first_time_vent = False

        '            Exit Sub

        'errorhandler:
        '            MsgBox("error in subroutine VENTFLOWS - terminating simulation")
        '            flagstop = 1

        '        End Sub

    End Module



End Namespace