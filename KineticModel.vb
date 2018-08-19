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

            For count = 1 To elements 'loop through each finite difference element in the ceiling
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
                ElementTemp = 273 + (stepcount - 1) * Timestep * 5 / 60

                DYDX_pyrol(k) = -A_array(k) * Exp(-E_array(k) / (Gas_Constant * ElementTemp)) * Y_pyrol(k) ^ n_array(k) 'arhennius equation

            Next

            Dim index As Integer = 0

            For var = 0 To Nvariables - 1
                dy(index) = DYDX_pyrol(var)
                index = index + 1
            Next

            Return dy

    End Function
End Module
