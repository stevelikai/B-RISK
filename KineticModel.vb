Option Strict Off
Option Explicit On
Imports System.Math
Imports CenterSpace.NMath.Core
Imports CenterSpace.NMath.Analysis


Module KineticModelCode
    Dim Y_pyrol() As Double
    Dim DYDX_pyrol() As Double

    Sub MLR_kinetic(ByVal i As Integer, maxceilingnodes As Integer, maxwallnodes As Integer)
        'called once per timestep
        'following the calculation of the nodal temperatures in the surface boundary

        Try

            'Initial mass fraction (of the wood solid)
            Dim mf_init(0 To 3) As Double '0 = H20; 1 = cellulose; 2 = hemicellulose; 3 = lignin
            mf_init(1) = 0.44
            mf_init(2) = 0.37
            mf_init(3) = 0.09
            mf_init(0) = 0.1
            Dim elements As Integer
            Dim Zstart() As Double
            Dim CharYield As Double = 0.13
            Dim DensityInitial As Double = 450 'kg/m3
            Dim chardensity As Double = 85 'kg/m3



            'the ceiling
            'CeilingNode(room, node, timestep) contains the temperature at each node at each timestep
            'CeilingElementMF (element,timsetep) contains the residual mass fraction of each component (relative to its initial value = 1) 

            ReDim Zstart(0 To 3)
            elements = maxceilingnodes - 1
            CeilingWoodMLR_tot(i + 1) = 0

            For count = 1 To elements 'loop through each finite difference element in the ceiling
                If i = 1 Then
                    For m = 1 To 3
                        CeilingResidualMass(count, i) = DensityInitial * mf_init(m) 'initialise
                    Next
                End If

                For m = 0 To 3
                    Zstart(m) = CeilingElementMF(count, m, i)
                Next
                elementcounter = count

                Call ODE_Solver_Pyrolysis(Zstart, i)

                For m = 0 To 3
                    CeilingElementMF(count, m, i + 1) = Max(Min(Zstart(m), 1), 0) 'residual mass fraction at the next time step
                Next

                'total mass fraction of char residue in this element
                CeilingCharResidue(count, i + 1) = (1 - CeilingElementMF(count, 1, i + 1)) * mf_init(1) * CharYield + (1 - CeilingElementMF(count, 2, i + 1)) * mf_init(2) * CharYield + (1 - CeilingElementMF(count, 3, i + 1)) * mf_init(3) * CharYield

                'total mass (per unit vol) of residual fuel (cellulose, hemicellulose, lignin) in this element
                CeilingResidualMass(count, i + 1) = DensityInitial * (CeilingElementMF(count, 1, i + 1) * mf_init(1) + CeilingElementMF(count, 2, i + 1) * mf_init(2) + CeilingElementMF(count, 3, i + 1) * mf_init(3)) 'kg/m3

                'mass loss rate of wood fuel over this timestep
                If i > 1 Then CeilingWoodMLR(count, i + 1) = -(CeilingResidualMass(count, i + 1) - CeilingResidualMass(count, i)) / Timestep 'kg/(s.m3)

                CeilingWoodMLR_tot(i + 1) = CeilingWoodMLR_tot(i + 1) + CeilingWoodMLR(count, i + 1) * CLTceilingpercent / 100 * RoomFloorArea(fireroom) * CeilingThickness(fireroom) / 1000 / elements 'kg/s
            Next


        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in " & Err.Source & " Line " & Err.Erl)
        End Try


    End Sub
    Sub ODE_Solver_Pyrolysis(ByRef Zstart() As Double, i As Integer)

        Try

            'Ystart holds the mass fractions of each of the components relative to their initial mass fraction =1.0
            'It tells us the residual fraction of that component.
            'this is not the mass fraction of the original solid wood that it represents 
            'for element = count

            Dim index As Integer = 0
            Dim x2, x1 As Double
            Dim Nvariables As Integer = 4 '4 components - cellulose, hemicellulose, lignin and water
            Dim VLength As Integer

            ReDim Y_pyrol(Nvariables)
            ReDim DYDX_pyrol(Nvariables)

            Y_pyrol = Zstart.Clone

            x1 = tim(i, 1) 'initial time
            x2 = x1 + Timestep 'final time

            Dim TimeSpan As New DoubleVector(x1, x2)
            Dim y0 As New DoubleVector()
            VLength = Nvariables

            y0.Resize(VLength)

            For var = 1 To Nvariables
                y0(index) = Y_pyrol(var - 1)
                index = index + 1
            Next

            ' Construct the solver.
            Dim Solver As New RungeKutta45OdeSolver()

            ' Construct the time span vector. If this vector contains exactly
            ' two points, the solver interprets these to be the initial and
            ' final time values. Step size and function output points are 
            ' provided automatically by the solver. Here the initial time
            ' value t0 is 0.0 and the final time value is 12.0.
            ' Dim TimeSpan As New DoubleVector(0.0, 12.0)

            ' Initial y vector.
            ' Dim y0 As New DoubleVector(0.0, 1.0, 1.0)

            ' Construct solver options. Here we set the absolute and relative tolerances to use.
            ' At the ith integration step the error, e(i) for the estimated solution
            ' y(i) satisfies
            ' e(i) <= max(RelativeTolerance * Math.Abs(y(i)), AbsoluteTolerance(i))
            ' The solver can increase the number of output points by a specified factor
            ' called Refine (useful for creating smooth plots). The default value is 
            ' 4. Here we set the Refine value to 1, meaning we do not wish any
            ' additional output points to be added by the solver.

            Dim SolverOptions As New RungeKutta45OdeSolver.Options()
            SolverOptions.AbsoluteTolerance = New DoubleVector(0.0001, 0.0001, 0.00001, 0.00001)
            SolverOptions.RelativeTolerance = 0.0001
            SolverOptions.Refine = 1

            ' Construct the delegate representing our system of differential equations...
            Dim odeFunctionPyrol As New Func(Of Double, DoubleVector, DoubleVector)(AddressOf RigidPyrol)

            ' ...and solve. The solution is returned as a key/value pair. The first 'Key' element of the pair is
            ' the time span vector, the second 'Value' element of the pair is the corresponding solution values.
            ' That is, if the computed solution function is y then
            ' y(soln.Key(i)) = soln.Value(i)
            Dim Soln As RungeKutta45OdeSolver.Solution(Of DoubleMatrix) = Solver.Solve(odeFunctionPyrol, TimeSpan, y0, SolverOptions)

            index = 0

            For var = 1 To Nvariables
                Y_pyrol(var - 1) = Soln.Y.Col(var - 1).Last
                index = index + 1
            Next

            Zstart = Y_pyrol.Clone

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in " & Err.Source & " Line " & Err.Erl)
            flagstop = 1
        End Try

    End Sub
    Function RigidPyrol(ByVal T As Double, ByVal Y As DoubleVector) As DoubleVector

        Dim Nvariables As Integer = 4
        Dim VectorLength As Integer
        Dim dy As New DoubleVector()

        VectorLength = Nvariables
        dy.Resize(VectorLength)

        'kinetic propeties for each component
        'Activation Energy
        Dim E_array(0 To 3) As Double '0 = H20; 1 = cellulose; 2 = hemicellulose; 3 = lignin
        E_array(1) = 198000.0 'J/mol cellulose
        E_array(2) = 164000.0 'J/mol hemicellulose
        E_array(3) = 152000.0 'J/mol lignin
        E_array(0) = 100000.0 'J/mol

        'Pre-exponential factor 
        Dim A_array(0 To 3) As Double '0 = H20; 1 = cellulose; 2 = hemicellulose; 3 = lignin
        A_array(1) = 351000000000000.0 '1/s
        A_array(2) = 32500000000000.0 '1/s
        A_array(3) = 84100000000000.0 '1/s
        A_array(0) = 10000000000000.0 '1/s

        'Reaction order
        Dim n_array(0 To 3) As Double '0 = H20; 1 = cellulose; 2 = hemicellulose; 3 = lignin
        n_array(1) = 1.1
        n_array(2) = 2.1
        n_array(3) = 5
        n_array(0) = 1

        Dim ElementTemp As Double = InteriorTemp
        'Gas_Constant As Double = 8.3145 'kJ/kmol K universal gas constant

        'put our ODE equations here
        'need to know the temperature of the current element at the current time - the same for all 4 components

        For k = 0 To 3

            'for the ceiling use CeilingNode(,,)
            ElementTemp = (CeilingNode(fireroom, elementcounter, stepcount) + CeilingNode(fireroom, elementcounter + 1, stepcount)) / 2 'use average temperature of the two adjacent nodes

            'test code element temperature rises 5K/min
            'ElementTemp = 293 + (stepcount - 1) * Timestep * 5 / 60

            DYDX_pyrol(k) = -A_array(k) * Exp(-E_array(k) / (Gas_Constant * ElementTemp)) * Y_pyrol(k) ^ n_array(k) 'arhennius equation

        Next

        Dim index As Integer = 0

        For var = 0 To Nvariables - 1
            dy(index) = DYDX_pyrol(var)
            index = index + 1
        Next

        Return dy

    End Function
    Function MassLoss_Total_Kinetic(ByVal T As Double, ByRef mwall As Double, ByRef mceiling As Double) As Double
        '*  ====================================================================
        '*  This function return the value of the total fuel mass loss rate for
        '*  a combination of wood cribs and burning wood surfaces at a given time T (sec)
        '*  Spearpoint, M.. & Quintiere, J.. 2000. Predicting the burning of wood using an integral model. 
        '*  Combustion and Flame. 123(3):308–325. DOI: 10.1016/S0010-2180(00)00162-0.
        '*  ====================================================================

        Static Tsave, mwallsave, mceilingsave, MLRsave As Double

        If T = Tsave Then
            'nO need to calculate function again for same T
            mceiling = mceilingsave
            mwall = mwallsave
            MassLoss_Total_Kinetic = MLRsave
            Exit Function
        End If

        Dim totalarea, total As Double
        Dim mass, wood_MLR As Double

        'this procedure only called if flashover is true and postflashover model is selected.
        'assume wood surface linings start burning at flashover

        totalarea = 0
        burnmode = False

        mass = InitialFuelMass 'kg wood cribs

        Dim thickness_ceil As Double = CeilingThickness(fireroom) / 1000 'm thick of wood
        Dim thickness_wall As Double = WallThickness(fireroom) / 1000  'm thick of wood

        'get HRR for all burning objects
        Dim total1, vp, total2, total3, Qtemp As Double
        Dim woodtotal As Double = 0

        If Flashover = True And g_post = True Then
            'in this case, we should only need this subroutine once per timestep

            'postflashover burning
            If TotalFuel(stepcount - 1) >= mass Then
                total = 0 'fuel contents is fully consumed
            Else

                'fuel surface control
                vp = 0.0000022 * Fuel_Thickness ^ (-0.6) 'wood crib fire regression rate m/s

                total1 = 4 / Fuel_Thickness * mass * vp * ((mass - TotalFuel(stepcount - 1)) / mass) ^ (0.5) 'kg/s

                'crib porosity control
                total2 = 0.00044 * (Stick_Spacing / Cribheight) * (mass / Fuel_Thickness)

                total = Min(total1, total2) 'use the lesser

                'this should use the max Q given the oxygen in the plume flow.
                Qtemp = massplumeflow(stepcount - 1, fireroom) * O2MassFraction(fireroom, stepcount - 1, 2) * 13100
                total3 = 1 / 1000 * Qtemp / NewHoC_fuel  'kJ/s / kJ/g = g/s

                If total3 < total Then
                    total = total3 'use the lesser, ventilation control
                    burnmode = True
                End If

                'better to apply excess fuel factor after determining mode of burning and then only if vent-limited
                If burnmode = True Then
                    total = total * ExcessFuelFactor

                End If

            End If

            'ceiling - interpolating
            mceiling = (CeilingWoodMLR_tot(stepcount + 1) - CeilingWoodMLR_tot(stepcount)) * (T - tim(stepcount, 1)) / Timestep + CeilingWoodMLR_tot(stepcount)

        End If

        If IEEERemainder(T, Timestep) = 0 Then
            If CLTwallpercent > 0 Then
                If Lamella2 > thickness_wall + gcd_Machine_Error Then
                    'no more fuel
                    mwall = 0
                End If
            End If
            If CLTceilingpercent > 0 Then
                If Lamella1 > thickness_ceil + gcd_Machine_Error Then
                    'no more fuel
                    mceiling = 0
                End If
            End If

        End If

        wall_char(stepcount, 1) = mwall 'kg/s
        ceil_char(stepcount, 1) = mceiling

        wood_MLR = mceiling + mwall 'kg/s  total MLR

        mwallsave = mwall
        mceilingsave = mceiling

        MassLoss_Total_Kinetic = total + wood_MLR 'includes contents+surfaces
        MLRsave = MassLoss_Total_Kinetic

        Tsave = T

    End Function
End Module
