Imports System
Imports System.Collections.Generic
Imports System.Diagnostics

Imports CenterSpace.NMath.Core
Imports CenterSpace.NMath.Analysis

'Namespace CenterSpace.NMath.Analysis.Examples.VisualBasic

' <summary>
' A .NET Example in Visual Basic showing how to use the RungeKutta45OdeSolver 
' to solve a nonstiff set of equations describing the motion of a rigid body 
' without external forces:
' 
' y1' = y2*y3,       y1(0) = 0
' y2' = -y1*y3,      y2(0) = 1
' y3' = -0.51*y1*y2, y3(0) = 1
' </summary>
Module RungeKutta45OdeSolverExample

        Sub Main()

            ' This section shows to use various configuration settings. All of these settings can
            ' also be set via configuration file or by environment variable. More on this at: 
            ' http://www.centerspace.net/blog/nmath/nmath-configuration/

            ' For NMath to continue to work after your evaluation period, you must set your license 
            ' key. You will receive a license key after purchase. 
            ' More here:
            ' http://www.centerspace.net/blog/nmath/setting-the-nmath-license-key/
            ' NMathConfiguration.LicenseKey = "<key>"

            ' This will start a log file that you can use to ensure that your configuration is 
            ' correct. This can be especially useful
            ' for deployment. Please turn this off when you are convinced everything is 
            ' working. If you aren't sure, please send the resulting log file to 
            ' support@centerspace.net. Please note that this directory can be relative or absolute.
            ' NMathConfiguration.LogLocation = "<dir>"

            ' NMath loads native, optimized libraries at runtime. There is a one-time cost to 
            ' doing so. To take control of when this happens, use Init(). If your program calls 
            ' Init() successfully, then your configuration is definitely correct.
            NMathConfiguration.Init()


            Console.WriteLine()

            ' Simple example solving the system of differential equations of
            ' motion for a rigid body with no external forces. The equations
            ' are defined above by the function Rigid().
            Console.WriteLine("Simple Solve -----------------------------------")
            Console.WriteLine()
            SimpleSolve()
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

            ' The above examples use the NMath class RungeKutta45OdeSolver. This 
            ' class uses an adaptive algorithm to figure out the optimal time step
            ' size at each integration step. NMath also supplies a class, 
            ' RungeKutta5OdeSolver that does NOT use an adaptive step size 
            ' algorithm, allowing the user to specify the time values for 
            ' each integration step.
            Console.WriteLine("Non-adaptive Step Size -------------------------")
            Console.WriteLine()
        NonAdaptiveStepSizeExample()

        Console.WriteLine()
            Console.WriteLine("Press Enter Key")
            Console.Read()

        End Sub

        Sub SimpleSolve()

        ' Construct the solver.
        Dim Solver As New RungeKutta45OdeSolver()

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
            Dim SolverOptions As New RungeKutta45OdeSolver.Options()
            SolverOptions.AbsoluteTolerance = New DoubleVector(0.0001, 0.0001, 0.00001)
            SolverOptions.RelativeTolerance = 0.0001
            SolverOptions.Refine = 1

            ' Construct the delegate representing our system of differential equations...
            Dim odeFunction As New Func(Of Double, DoubleVector, DoubleVector)(AddressOf Rigid)

            ' ...and solve. The solution is returned as a key/value pair. The first 'Key' element of the pair is
            ' the time span vector, the second 'Value' element of the pair is the corresponding solution values.
            ' That is, if the computed solution function is y then
            ' y(soln.Key(i)) = soln.Value(i)
            Dim Soln As RungeKutta45OdeSolver.Solution(Of DoubleMatrix) = Solver.Solve(odeFunction, TimeSpan, y0, SolverOptions)
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
    Function OutputFunction(ByVal T As DoubleVector, ByVal Y As DoubleMatrix, ByVal Flag As RungeKutta45OdeSolver.OutputFunctionFlag) As Boolean

        If Flag = RungeKutta45OdeSolver.OutputFunctionFlag.Initialize Then
            Console.WriteLine("Output function : Initialize")
            Return True
        End If

        If Flag = RungeKutta45OdeSolver.OutputFunctionFlag.IntegrationStep Then
            Console.WriteLine("Output function Integration step:")
            Console.WriteLine("t = " + T.ToString("G5"))
            Console.WriteLine("y = ")
            Console.WriteLine(Y.ToTabDelimited("G5"))
            Return True
        End If

        If Flag = RungeKutta45OdeSolver.OutputFunctionFlag.Done Then
            Console.WriteLine("Output function: Done")
            Return True
        End If

        Console.WriteLine("Output function: Unknown flag")
        Return False

    End Function

        Sub WithOutputFunctionAndMassMatrix()

        ' Construct the solver.
        Dim Solver As New RungeKutta45OdeSolver()

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
        Dim SolverOptions As New RungeKutta45OdeSolver.Options()
        SolverOptions.AbsoluteTolerance = New DoubleVector(0.0001, 0.0001, 0.00001)

        SolverOptions.OutputFunction = New Func(Of DoubleVector, DoubleMatrix, RungeKutta45OdeSolver.OutputFunctionFlag, Boolean)(AddressOf OutputFunction)
        SolverOptions.RelativeTolerance = 0.0001

        ' Construct the delegate representing our system of differential equations...
        Dim odeFunction As New Func(Of Double, DoubleVector, DoubleVector)(AddressOf Rigid)

        ' ...and solve. The solution is returned as a key/value pair. The first 'Key' element of the pair is
        ' the time span vector, the second 'Value' element of the pair is the corresponding solution values.
        ' That is, if the computed solution function is y then
        ' y(soln.Key(i)) = soln.Value(i)
        Dim Soln As RungeKutta45OdeSolver.Solution(Of DoubleMatrix) = Solver.Solve(odeFunction, TimeSpan, y0, SolverOptions)

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
        Soln = Solver.Solve(odeFunction, TimeSpan, y0)

    End Sub

    Sub NonAdaptiveStepSizeExample()

        ' construct the solver.
        Dim solver As New RungeKutta5OdeSolver()

        ' construct the time span vector. imports the non-adaptive solve
        ' function, integration steps will be performed only at the points
        ' in the time span vector.
        Dim timespan As New DoubleVector(85, 0, 0.15)

        ' initial y vector.
        Dim y0 As New DoubleVector(0.0, 1.0, 1.0)

        ' construct the delegate representing our system of differential equations...
        Dim odefunction As New Func(Of Double, DoubleVector, DoubleVector)(AddressOf Rigid)

        ' ...and solve. each row of the returned matrix contains the calculated function values at the
        ' corresponding point in the input time span vector.
        Dim soln As DoubleMatrix = solver.Solve(odefunction, timespan, y0)

        ' print out a few values
        Console.WriteLine("number of time values = " & timespan.Length)
        For i As Integer = 0 To timespan.Length - 1
            If (i Mod 7 = 0) Then
                Console.WriteLine("y({0}) = {1}", timespan(i).ToString("g5"), soln.Row(i).ToString("g5"))
            End If
        Next

        ' solve with a constant mass matrix
        Dim m As New DoubleMatrix("3x3[1     2     1 2     1     3     9     4     1]")
        Dim solveroptions As New RungeKutta5OdeSolver.Options()
        solveroptions.MassMatrix = m
        soln = solver.Solve(odefunction, timespan, y0, solveroptions)

        ' print out a few values for the mass matrix solution.
        Console.WriteLine(Environment.NewLine & "with constant mass matrix:")
        Console.WriteLine("number of time values = " & timespan.Length)
        For i As Integer = 0 To timespan.Length - 1
            If (i Mod 7 = 0) Then
                Console.WriteLine("y({0}) = {1}", timespan(i).ToString("g5"), soln.Row(i).ToString("g5"))
            End If
        Next

        ' solve with an output function that is a class member function.
        Dim outputfunctionobject As New SimpleNonAdaptiveOutput()
        solveroptions.OutputFunction = New Func(Of Double, DoubleVector, RungeKutta45OdeSolver.OutputFunctionFlag, Boolean)(AddressOf outputfunctionobject.outputfunction)
        soln = solver.Solve(odefunction, timespan, y0, solveroptions)

    End Sub
    End Module

    ' Class containing an example output function for the Dormand Prince ODE non-adaptive solve
    ' function. The class does nothing but accumulate the solutions at each solver 
    ' integration stop.
    Class SimpleNonAdaptiveOutput

    ' Gets and sets a list of the time values.
    Private _T As List(Of Double) = Nothing
    Public Property T As List(Of Double)
        Get
            Return _T
        End Get
        Set(ByVal value As List(Of Double))
            _T = value
        End Set
    End Property

    '  Gets And sets a list of the calculated function values.
    Private _Y As List(Of DoubleVector) = Nothing
    Public Property Y As List(Of DoubleVector)
        Get
            Return _Y
        End Get
        Set(ByVal value As List(Of DoubleVector))
            _Y = value
        End Set
    End Property

    ' Gets and sets a flag indicating whether or not the output
    ' function/class has been initialized.
    Private _Init As Boolean = False
    Public Property IsInitialized As Boolean
        Get
            Return _Init
        End Get
        Set(ByVal value As Boolean)
            _Init = value
        End Set
    End Property

    ' Flag indicating the output function is done being invoked.
    Private _IsDone As Boolean = False
    Public Property IsDone As Boolean
        Get
            Return _IsDone
        End Get
        Set(ByVal value As Boolean)
            _IsDone = value
        End Set
    End Property

    ' Constructs and initializes a SimpleNonAdaptiveOutput object.
    Public Sub SimpleNonAdaptiveOutput()
        IsInitialized = False
        IsDone = False
        T = Nothing
        Y = Nothing
    End Sub

    ' <summary>
    ' The output function to be called by the ODE solver at each integration step. If this function returns true
    ' the solver continues with the next step, if it returns false the solver is halted.
    ' </summary>
    ' <param name="Time">The time value at the current integration step. If flag is equal to Initialize,
    ' this will be the initial time value t0.</param>
    ' <param name="YValue">The calculated function value at the current integration step. If flag is equal to Initialize,
    ' this will be the initial function value y0.</param>
    ' <param name="Flag">Flag indicating what stage the solver is at:
    ' Initialize - before solving begins. y and t values are the problems initial values.
    ' IntegrationStep - just finished an integration step. y and t values are the values
    ' calculated at that step.
    ' Done - just finished last step. y and t values are the last values in the returned 
    ' solution.</param>
    ' <returns>true if the solver is to proceed with the calculation, false forces the solver
    ' to stop.</returns>
    Public Function OutputFunction(ByVal Time As Double, ByVal YValue As DoubleVector, ByVal Flag As RungeKutta45OdeSolver.OutputFunctionFlag) As Boolean

        If (Flag = RungeKutta45OdeSolver.OutputFunctionFlag.Initialize) Then
            T = New List(Of Double)()
            T.Add(Time)
            Y = New List(Of DoubleVector)()
            Y.Add(YValue)
            IsInitialized = True
            IsDone = False
            Return True
        End If

        If (Flag = RungeKutta45OdeSolver.OutputFunctionFlag.Done) Then
            IsDone = True
            T.Add(Time)
            Y.Add(YValue)
            Return True
        End If

        If (Flag = RungeKutta45OdeSolver.OutputFunctionFlag.IntegrationStep) Then
            T.Add(Time)
            Y.Add(YValue)
            Return True
        End If

        Throw New InvalidArgumentException("Unknown output function flag: " + Flag)

    End Function

End Class

'End Namespace
