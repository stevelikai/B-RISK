Option Strict Off
Option Explicit On
Imports System.Data
Imports System.Math
Imports System.Xml
Imports System.Xml.Linq
Imports System.Xml.Serialization
Imports System.IO
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Collections.Generic
Imports Microsoft.Win32
Imports CenterSpace.NMath.Core
Imports System.Reflection.MethodBase

Module MAIN
    Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Integer, ByVal lpBuffer As String) As Integer
    Public Const MAX_BUFFER_LENGTH As Short = 256
    Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer

    Dim imax() As Short
    Dim ig_step() As Short
    Dim data_keep(,,) As Single
    Dim oSprinklers As New List(Of oSprinkler)


    Dim result1 As Double
    Dim result2 As Double


    Public Function GetTempPathName() As String
        Dim strBufferString As String
        Dim lngResult As Integer

        strBufferString = New String("X", MAX_BUFFER_LENGTH)

        lngResult = GetTempPath(MAX_BUFFER_LENGTH, strBufferString)

        GetTempPathName = Mid(strBufferString, 1, lngResult)

    End Function
    Public Sub FRating(ByRef tlimit As Double, ByRef energy As Double, ByRef room As Integer)
        '==============================================
        ' calculate the cumulative radiant energy input to compartment barrier
        ' 10/10/2002
        '================================================
        Dim i As Integer
        Dim E As Double

        Try
            energy = 0
            E = 0

            For i = 1 To NumberTimeSteps
                E = Timestep * StefanBoltzmann * uppertemp(room, i) ^ 4 'Ws/m2
                'E = Timestep * StefanBoltzmann * QCeilingAST(room, 2, i) ^ 4 'Ws/m2
                'E = Timestep * StefanBoltzmann * QUpperWallAST(room, 2, i) ^ 4 'Ws/m2
                energy = energy + E / 1000 'kJ/m2
                If tim(i, 1) >= tlimit Then Exit For
            Next i

        Catch ex As Exception
            MsgBox(Err.Description & " Line " & Err.Erl, MsgBoxStyle.Exclamation, "Exception in FRating ")
        End Try

    End Sub

    Function TANH(ByRef X As Double) As Double

        ' Hyperbolic tangent function
        TANH = (1.0# - Exp(-2.0# * X)) / (1.0# + Exp(-2.0# * X))

    End Function

    Function CeilingJet_Temp(ByVal UT As Double, ByVal Q As Double, ByVal radialdistance As Single) As Double
        '*  =============================================================
        '*      This function returns the value of the maximum temperature
        '*      in the ceiling jet at the location of the sprinkler or
        '*      detector using the Alpert quasi-steady correlations.
        '*
        '*      Arguments passed to the function:
        '*      Q = heat release rate for object 1
        '*  =============================================================

        Dim h As Single

        'ceiling height above base of fire
        h = RoomHeight(fireroom) - FireHeight(1)

        If radialdistance / h <= 0.18 Then
            'alpert
            CeilingJet_Temp = 16.9 * Q ^ (2 / 3) * h ^ (-5 / 3)
            Exit Function
        Else
            'alpert
            CeilingJet_Temp = 5.38 * (Q / radialdistance) ^ (2 / 3) / h
            Exit Function
        End If

    End Function
    Function CeilingJet_Temp2(ByVal UT As Double, ByVal Q As Double, ByVal radialdistance As Single) As Double
        '*  =============================================================
        '*      This function returns the value of the maximum temperature rise
        '*      in the ceiling jet at the location of the sprinkler or
        '*      detector using the Alpert quasi-steady correlations.
        '*
        '*      Arguments passed to the function:
        '*      Q = heat release rate for object 1
        '*  =============================================================

        Dim h, rh As Single

        'ceiling height above base of fire
        h = RoomHeight(fireroom) - FireHeight(1)
        rh = radialdistance / h

        If rh <= 0.18 Then
            'alpert
            CeilingJet_Temp2 = 16.9 * Q ^ (2 / 3) * h ^ (-5 / 3)
            Exit Function
        Else
            'alpert
            CeilingJet_Temp2 = 5.38 * Q ^ (2 / 3) * h ^ (-5 / 3) / rh ^ (2 / 3)
            Exit Function
        End If

    End Function

    Function CeilingJet_Velocity(ByVal Q As Double, ByVal radialdistance As Single) As Double
        '*  =============================================================
        '*      This function returns the value of the maximum velocity
        '*      in the ceiling jet at the location of the sprinkler or
        '*      detector using the Alpert quasi-steady correlations.
        '*
        '*      Arguments passed to the function:
        '*      Q = heat release rate for object 1
        '*  =============================================================

        Dim h As Single

        If Q < 0 Then Q = 0

        'ceiling height above base of fire
        h = RoomHeight(fireroom) - FireHeight(1)

        If radialdistance / h <= 0.15 Then
            CeilingJet_Velocity = 0.96 * (Q / h) ^ (1 / 3)
            Exit Function
        Else
            CeilingJet_Velocity = 0.195 * Q ^ (1 / 3) * Sqrt(h) * radialdistance ^ (-5 / 6)
            Exit Function
        End If

    End Function
    Function CeilingJet_Velocity2(ByVal Q As Double, ByVal radialdistance As Single) As Double
        '*  =============================================================
        '*      This function returns the value of the maximum velocity
        '*      in the ceiling jet at the location of the sprinkler or
        '*      detector using the Alpert quasi-steady correlations.
        '*      5th ed SFPE HB
        '*      Arguments passed to the function:
        '*      Q = heat release rate for object 1
        '*  =============================================================

        Dim rh, h As Single

        If Q < 0 Then Q = 0

        'ceiling height above base of fire
        h = RoomHeight(fireroom) - FireHeight(1)

        rh = radialdistance / h

        If rh <= 0.15 Then
            CeilingJet_Velocity2 = 0.947 * (Q / h) ^ (1 / 3)
            Exit Function
        Else
            CeilingJet_Velocity2 = 0.197 * (Q / h) ^ (1 / 3) / rh ^ (-5 / 6)
            Exit Function
        End If

    End Function

    Sub Centre_Form(ByRef FormName As System.Windows.Forms.Form)
        '*  ====================================================================
        '*  Centre the current form on the screen
        '*  ====================================================================

        FormName.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(FormName.Width)) / 2)
        FormName.Top = VB6.TwipsToPixelsY(1.1 * (VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(FormName.Height)) / 2)

    End Sub

    Sub Check_Vent()
        '*  ===================================================================
        '*  Check that the size of the vent specified is valid.
        '*  ===================================================================

        Dim j, i, k As Short
        Dim totalarea, ventarea1 As Double
        Dim highest_floor, lowest_ceiling As Double

        For j = 1 To NumberRooms + 1
            For k = 1 To NumberRooms + 1
                If j = k Then NumberVents(j, k) = 0
            Next k
        Next j

        For j = 1 To NumberRooms
            For k = 2 To NumberRooms + 1
                If NumberVents(j, k) > 0 And j < k Then
                    'find highest floor
                    If k <> NumberRooms + 1 Then
                        If FloorElevation(j) > FloorElevation(k) Then
                            highest_floor = FloorElevation(j)
                        Else
                            highest_floor = FloorElevation(k)
                        End If
                        lowest_ceiling = DMin1(RoomHeight(j) + FloorElevation(j), RoomHeight(k) + FloorElevation(k))
                    Else
                        highest_floor = FloorElevation(j)
                        lowest_ceiling = RoomHeight(j) + FloorElevation(j)
                    End If

                    For i = 1 To NumberVents(j, k)
                        If VentSillHeight(j, k, i) + FloorElevation(j) < highest_floor Then
                            MsgBox("Vent sill height not valid for Vent " & CStr(i) & " between rooms " & CStr(j) & " and " & CStr(k), MsgBoxStyle.OkOnly + MsgBoxStyle.Information)

                        End If

                        If VentSillHeight(j, k, i) + VentHeight(j, k, i) + FloorElevation(j) > lowest_ceiling Then
                            'The vent size specified is too high for the room height.
                            'The vent height has been reduced to fit."
                            MsgBox("Vent height not valid for Vent " & CStr(i) & " between rooms " & CStr(j) & " and " & CStr(k), MsgBoxStyle.OkOnly + MsgBoxStyle.Information)
                        End If

                        If VentWidth(j, k, i) > 2 * (RoomLength(j) + RoomWidth(j)) Then
                            'The vent size specified is too wide for the room size.
                            'The vent width has been reduced to fit."
                            MsgBox("Vent width not valid for Vent " & CStr(i) & " between rooms " & CStr(j) & " and " & CStr(k), MsgBoxStyle.OkOnly + MsgBoxStyle.Information)
                        End If
                    Next i

                    'check the total area of all vents does not exceed the wall area
                    totalarea = 0
                    ventarea1 = 0
                    For i = 1 To NumberVents(j, k)
                        ventarea1 = VentHeight(j, k, i) * VentWidth(j, k, i)
                        totalarea = totalarea + ventarea1
                    Next i

                    If totalarea > 2 * RoomHeight(j) * (RoomLength(j) + RoomWidth(j)) Then
                        MsgBox("The total area of all vents exceeds the total wall area. Please review your input data.")
                        Exit Sub
                    End If
                End If
            Next k
        Next j
    End Sub

    Function COMass_Rate2(ByVal GER As Double, ByVal T As Double, ByVal massloss As Double, ByVal plumeflow As Double, ByVal UT As Double) As Double
        '*  ==================================================================
        '*  This function returns the value of the sum of the fuel mass loss rates
        '*  multiplied by the species generation rate for all burning objects.
        '*  ==================================================================


        If comode = True Then 'manual entry from user
            If Flashover = False Then
                If VentilationLimitFlag = True And VM2 = True Then
                    'case 5 VM2 rules
                    COYield = postCO
                Else
                    COYield = preCO
                End If

            Else
                COYield = postCO
            End If

        Else 'default,  function to give conservative fit to gottuk's data
            If GER < 0.6 Then
                COYield = 0.04
            ElseIf GER > 1 Then
                COYield = 0.3
            Else
                COYield = 0.4 + (GER - 0.6) * (0.3 - 0.04) / (1 - 0.6)
            End If
        End If

        COMass_Rate2 = massloss * COYield 'kg-CO/s

    End Function

    Function COMass_Rate(ByVal GER As Double, ByVal T As Double, ByVal massloss As Double, ByVal plumeflow As Double, ByVal UT As Double) As Double
        '*  ==================================================================
        '*  This function returns the value of the sum of the fuel mass loss rates
        '*  multiplied by the species generation rate for all burning objects.
        '*  Revised 19/12/97 Colleen Wade
        '*  ==================================================================

        Dim total1, total, total2 As Single

        If comode = True Then 'manual entry from user
            If Flashover = False Then
                COYield = preCO
            Else
                COYield = postCO
            End If
        Else
            'from NIST GCR 92 619 Gottuk pg 183
            'use the upper envelope of the two curves for more conservative result
            '2005.1 25/07/05
            'If UT > 875 Then
            total1 = ((0.22 / PI) * Atan(10 * (GER - 1.25)) + 0.11)
            'Else
            total2 = ((0.19 / PI) * Atan(10 * (GER - 0.8)) + 0.095)
            'End If
            If total2 > total1 Then 'use the higher of the two
                total = total2
            Else
                total = total1
            End If
            COYield = total
        End If

        COMass_Rate = massloss * COYield 'kg-CO/s

    End Function

    Function Composite_HRR(ByVal tim As Double) As Double
        '*  ===================================================================
        '*  This function returns the total heat release from all the objects
        '*
        '*  tim =time (sec)
        '*
        '*  Revised 6 December 1996 Colleen Wade
        ' 1 feb 2008 Amanda
        '*  ===================================================================

        Dim dummy As Double = 0
        Dim total As Double = 0
        Dim QCeiling, QW, QWall, QFloor As Double
        Dim i As Integer
        Static post As Boolean

        If gb_first_time_vent = True Then post = frmOptions1.optPostFlashover.Checked

        If Flashover = True And post = True Then
            Composite_HRR = Get_HRR(1, tim, QW, Qburner, QFloor, QWall, QCeiling)
            Exit Function
        ElseIf post = False And useCLTmodel = True Then
            Composite_HRR = Get_HRR(1, tim, QW, Qburner, QFloor, QWall, QCeiling)
            Composite_HRR = Composite_HRR + QWall + QCeiling
            Exit Function

        Else
            For i = 1 To NumberObjects
                dummy = Get_HRR(i, tim, QW, Qburner, QFloor, QWall, QCeiling)

                'keep track of the cumulative heat release rate
                total = total + dummy
            Next i

            Composite_HRR = total
        End If
    End Function

    Function Confirm_File(ByRef opendatafile As String, ByRef operation As Short, ByRef errorsuppress As Short) As Short
        '*  ==================================================================
        '*  This checking routine uses the value returned by the fileerrors
        '*  procedure to decide what action to take in the event of a disk
        '*  related error.
        '*  ==================================================================

        Dim thefile As String = ""
        Dim nl, msg As String
        Dim Confirmation, Action As Short

        Try

            nl = Chr(13) & Chr(10)

            'On Error GoTo ConfirmFileError
            'If opendatafile <> "" Then thefile = Dir(LCase(opendatafile), 0)

            'this extracts the filename from the string, it doesn't check it exists
            If opendatafile <> "" Then thefile = Path.GetFileName(opendatafile)


            'if user is saving to file that already exists
            If thefile <> "" And operation = replacefile Then
                'If File.Exists(opendatafile) And operation = replacefile Then
                msg = "The file " & opendatafile & " already exists on disk." & nl
                msg = msg & " Choose OK to overwrite the file, Cancel to stop." & nl
                Confirmation = MsgBox(msg, 49, "ProgramTitle")

                'if user wants to load data from nonexistent file
            ElseIf thefile = "" And operation = readfile Then
                'ElseIf Not File.Exists(opendatafile) And operation = readfile Then
                'If InStr(opendatafile, "ISO9705_") = 0 Then
                If errorsuppress = 0 Then
                    msg = "The file " & opendatafile & " doesn't exist."
                    Confirmation = MsgBox(msg, 64, ProgramTitle) + 1
                Else
                    Confirmation = 2
                End If
                'End If
                'if the file doesn't exist
                'ElseIf Not File.Exists(opendatafile) = "" Then
            ElseIf thefile = "" Then
                If operation = randomfile Or operation = binaryfile Then
                    Confirmation = 2
                Else
                    Confirmation = 3
                End If
            End If

            If Confirmation > 1 Then Confirm_File = 0 Else Confirm_File = 1
            If Confirmation = 1 Then
                If operation = loadfile Then
                    operation = replacefile
                End If
            End If
            Exit Function

        Catch ex As Exception
            MsgBox(Err.Description & " Line " & Err.Erl, MsgBoxStyle.Exclamation, "Exception in Confirm_File ")

            Action = File_Errors(Err.Number)
            Select Case Action
                Case 0
                    Exit Function
                Case 1
                    'Stop
                Case 2
                    Exit Function
                Case Else
                    Error (Err.Number)
            End Select
        End Try
    End Function

    Sub Derived_Variables()
        '*  ==================================================================
        '*  Calculate variables derived directly from input data.
        '*  ==================================================================

        Dim radflux(20) As Double
        Dim IgnitionTime(20) As Double
        Dim Peak(20) As Double
        Dim IgnPoints, room As Integer
        Dim IO, slope As Double
        Dim Ignition(20) As Double
        ReDim AreaUnderCeilingCurve(NumberRooms)
        Dim Yintercept, RMax, Nmax, R2 As Double
        Dim n As Single
        Dim i As Integer
        Dim text As String

        On Error GoTo error_derived

        'Error_Control = error_control_param 'hardwire this instead of user changeable

        'calculate volume of the room
        Room_Volume()

        'calculate floor area of the room
        Room_FloorArea()

        'find the soffit height of the vents
        Soffit_Height()

        'find the highest vent
        Max_Vent()

        If usepowerlawdesignfire = True Then
            InitialFuelMass = FLED * RoomFloorArea(fireroom) / EnergyYield(1) 'original fuel mass
        End If
        'initial mass of fuel for postflashover option
        If frmOptions1.optPostFlashover.Checked = True Then
            'If g_post = True Then
            InitialFuelMass = FLED * RoomFloorArea(fireroom) / HoC_fuel 'original fuel mass
        End If


        'combustion chemistry
        'ReDim HCNYield(1 To NumberObjects)
        MolecularWeightFuel = nC * 12 + nH * 1 + nO * 16 + nN * 14

        chemistry()

        For room = 1 To NumberRooms
            ThermalInertiaCeiling(room) = 0.000001 * CeilingConductivity(room) * CeilingDensity(room) * CeilingSpecificHeat(room)
            ThermalInertiaWall(room) = 0.000001 * WallConductivity(room) * WallDensity(room) * WallSpecificHeat(room)
            ThermalInertiaFloor(room) = 0.000001 * FloorConductivity(room) * FloorDensity(room) * FloorSpecificHeat(room)
        Next

        If frmOptions1.optRCNone.Checked = False Then
            'use karlsson's or quintiere's model

            'get lining flame spread properties

            For room = 1 To NumberRooms
                If WallConeDataFile(room) = "" Then WallConeDataFile(room) = "null.txt"
                If WallConeDataFile(room) <> "null.txt" Then
                    Call Flame_Spread_Props(room, 1, wallhigh, WallConeDataFile(room), Surface_Emissivity(2, room), ThermalInertiaWall(room), IgTempW(room), WallEffectiveHeatofCombustion(room), WallHeatofGasification(room), AreaUnderWallCurve(room), WallFTP(room), WallQCrit(room), Walln(room))

                    '2012 test - confusion about n vs 1/n
                    'Call Flame_Spread_Props2(room, 1, wallhigh, WallConeDataFile(room), Surface_Emissivity(2, room), ThermalInertiaWall(room), IgTempW(room), WallEffectiveHeatofCombustion(room), WallHeatofGasification(room), AreaUnderWallCurve(room), WallFTP(room), WallQCrit(room), Walln(room))

                End If
                If CeilingConeDataFile(room) = "" Then CeilingConeDataFile(room) = "null.txt"
                If CeilingConeDataFile(room) <> "null.txt" Then
                    Call Flame_Spread_Props(room, 2, ceilinghigh, CeilingConeDataFile(room), Surface_Emissivity(1, room), ThermalInertiaCeiling(room), IgTempC(room), CeilingEffectiveHeatofCombustion(room), CeilingHeatofGasification(room), AreaUnderCeilingCurve(room), CeilingFTP(room), CeilingQCrit(room), Ceilingn(room))
                End If
                If FloorConeDataFile(room) = "" Then FloorConeDataFile(room) = "null.txt"
                If FloorConeDataFile(room) <> "null.txt" Then
                    Call Flame_Spread_Props(room, 3, floorhigh, FloorConeDataFile(room), Surface_Emissivity(4, room), ThermalInertiaFloor(room), IgTempF(room), FloorEffectiveHeatofCombustion(room), FloorHeatofGasification(room), AreaUnderFloorCurve(room), FloorFTP(room), FloorQCrit(room), Floorn(room))
                End If
            Next room

            'Wall area behind the burner - used with Karlsson's room corner model
            WallAreaBurner = 2 * BurnerWidth * (RoomHeight(fireroom) - FireHeight(1))

            'continuous flame height, hasemi and tokunaga, SFPE handbook pg 2-5
            BurnerDiameter = Sqrt(4 * BurnerWidth ^ 2 / PI)

            'make sure that the thermal properties are consistent with user entries
            'in the room material screen. User values for density and thermal conductivity
            'will be accepted and the specific heat value will be adjusted to
            'be consistent with the value of thermal inertia given here.
            For room = 1 To NumberRooms
                If CeilingConeDataFile(room) = "null.txt" Or ThermalInertiaCeiling(room) = 0 Then ThermalInertiaCeiling(room) = 10 ^ (-6) * CeilingConductivity(room) * CeilingDensity(room) * CeilingSpecificHeat(room)
                If WallConeDataFile(room) = "null.txt" Or ThermalInertiaWall(room) = 0 Then ThermalInertiaWall(room) = 10 ^ (-6) * CeilingConductivity(room) * CeilingDensity(room) * CeilingSpecificHeat(room)
                If FloorConeDataFile(room) = "null.txt" Or ThermalInertiaFloor(room) = 0 Then ThermalInertiaFloor(room) = 10 ^ (-6) * FloorConductivity(room) * FloorDensity(room) * FloorSpecificHeat(room)
            Next room

        End If

        'save vent opening times
        ReDim oldCVentOpenTime(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)
        ReDim oldVentOpenTime(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
        ReDim WAfloor_flag(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
        Dim j, vent As Short
        For room = 1 To NumberRooms
            For j = 1 To NumberRooms + 1
                If j > room Then
                    For vent = 1 To NumberVents(room, j)
                        oldVentOpenTime(room, j, vent) = VentOpenTime(room, j, vent)
                    Next vent
                End If
            Next j
        Next room

        'check emissivities
        For room = 1 To NumberRooms
            For i = 1 To 4
                If Surface_Emissivity(i, room) <= 1 And Surface_Emissivity(i, room) > 0 Then
                    'ok
                Else
                    'not allowed
                    If i = 1 Then
                        text = "The emissivity of the ceiling material in room " & VB6.Format(room) & " is not a valid value. Please check your input."
                        MsgBox(text)
                    ElseIf i = 2 Or i = 3 Then
                        text = "The emissivity of the wall material in room " & VB6.Format(room) & " is not a valid value. Please check your input."
                        MsgBox(text)
                    Else
                        text = "The emissivity of the floor material in room " & VB6.Format(room) & " is not a valid value. Please check your input."
                        MsgBox(text)
                    End If
                    flagstop = 1
                    Exit Sub
                End If
            Next i
        Next room

        'get material thermal properties
        implicit_thermal_props() 'implicit method

        ' If useCLTmodel = True And IntegralModel = True Then
        If useCLTmodel = True Then
            Lamella1 = Lamella 'initialise value
            Lamella2 = Lamella
            CLTwalldelamT = 0
            CLTceildelamT = 0
            result1 = 0
            result2 = 0
        End If
        Exit Sub

error_derived:
        If Err.Number = 75 Then
            MsgBox("Either the wall or ceiling lining material does not have a heat release rate input file. You need to select another material or supply the required data.")
        End If
        Exit Sub

    End Sub

    Sub EndPoint_Flags(ByRef tim As Double, ByRef Target As Double, ByRef UT As Double, ByRef Visibility1 As Double, ByRef sprink As Double, ByRef layer As Double, ByRef LT As Double, ByRef hrr As Double)
        '*  =====================================================
        '*      This procedure evaluates the endpoint conditions
        '*      and sets flag if the condition is reached.
        '*
        '*      Arguments passed to the procedure
        '*          tim = time (seconds)
        '*          Target = incident radiant flux on target (kw/m2)
        '*          UT = upper layer temperature (K)
        '*          Visibility = visibility in the upper layer (m)
        '*          Sprink = link temperature of sprinkler/detector (K)
        '*
        '*  Revised 3 December 1997 Colleen Wade
        '*  =====================================================

        Static UTFLag, ConvectFlag, TargetFlag, TempFlag, VisibilityFlag As Short

        If tim = 0 Then
            TargetFlag = 0
            TempFlag = 0
            VisibilityFlag = 0
            ConvectFlag = 0
            SprinklerFlag = 0
            HDFlag = 0
            UTflag = 0
        End If

        If Target >= TargetEndPoint And TargetFlag = 0 Then
            frmEndPoints.lblTarget.Text = "FED Thermal (incap) Exceeds " & Format(TargetEndPoint) & " at " & Format(tim, "0.0") & " Seconds."
            TargetFlag = 1
        End If

        If UT >= TempEndPoint And TempFlag = 0 Then
            frmEndPoints.lblTemp.Text = "Upper Layer Temperature Exceeds " & Format(TempEndPoint - 273) & " deg C at " & Format(tim, "0.0") & " Seconds."

            TempFlag = 1
        End If

        For room = 1 To NumberRooms

            If VM2 = True And UTFLag = 0 Then
                'If MonitorHeight > layerheight(room, stepcount) Then
                If uppertemp(room, stepcount) > TempEndPoint Then
                    Dim Message As String = CStr(tim) & " sec.  Upper layer temperature exceeds " & Format(TempEndPoint - 273) & " C in room " & room.ToString
                    If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                    UTFLag = 1
                End If
                'Else
                '    If lowertemp(room, stepcount) > 473 Then
                '        Dim Message As String = CStr(tim) & " sec.  Layer Temperature at " & CStr(MonitorHeight) & "m above floor exceeds " & Format(TempEndPoint - 273) & " C in room " & room.ToString
                '        If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                '        UTFLag = 1
                '    End If
                'End If
            End If

            If Visibility(room, stepcount) <= VisibilityEndPoint And VisibilityFlag = 0 Then
                frmEndPoints.lblVisibility.Text = "Visibility at " & CStr(MonitorHeight) & "m above floor reduced to " & VB6.Format(VisibilityEndPoint) & " m at " & VB6.Format(tim, "0.0") & " sec" & " in room " & room.ToString
                VisibilityFlag = 1
                If MonitorHeight > RoomHeight(room) Then
                    Dim Message As String = CStr(tim) & " sec. Visibility at " & CStr(RoomHeight(room)) & "m above floor reduced to " & VB6.Format(VisibilityEndPoint) & " m in room " & room.ToString
                    If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                Else
                    Dim Message As String = CStr(tim) & " sec. Visibility at " & CStr(MonitorHeight) & "m above floor reduced to " & VB6.Format(VisibilityEndPoint) & " m in room " & room.ToString
                    If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                End If
            End If
        Next
        'If Visibility1 <= VisibilityEndPoint And VisibilityFlag = 0 Then
        '    frmEndPoints.lblVisibility.Text = "Visibility at " & CStr(MonitorHeight) & "m above floor reduced to " & VB6.Format(VisibilityEndPoint) & " m at " & VB6.Format(tim, "0.0") & " Seconds."
        '    VisibilityFlag = 1
        '    Dim Message As String = CStr(tim) & " sec. Visibility at " & CStr(MonitorHeight) & "m above floor reduced to " & VB6.Format(VisibilityEndPoint) & " m."
        '    If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

        'End If

        If layer <= MonitorHeight Then
            If UT >= ConvectEndPoint And ConvectFlag = 0 Then
                frmEndPoints.lblConvect.Text = "Temperature at " & CStr(MonitorHeight) & "m above floor has reached " & VB6.Format(ConvectEndPoint - 273) & " deg C at " & VB6.Format(tim, "0.0") & " Seconds."
                ConvectFlag = 1
            End If
        Else
            If LT >= ConvectEndPoint And ConvectFlag = 0 Then
                frmEndPoints.lblConvect.Text = "Temperature at " & CStr(MonitorHeight) & "m above floor has reached " & VB6.Format(ConvectEndPoint - 273) & " deg C at " & VB6.Format(tim, "0.0") & " Seconds."
                ConvectFlag = 1
            End If
        End If

        'need to check if there is a sprinkler first !!!
        'If DetectorType(fireroom) > 0 Then
        '    If ActuationTemp(fireroom) <= sprink And SprinklerFlag = 0 Then
        '        frmEndPoints.lblSprinkler.Text = "Sprinkler/Detector Actuated at " & VB6.Format(tim, "0.0") & " Seconds."
        '        SprinklerFlag = 1
        '        SprinklerTime = tim 'time of actuation
        '        SprinklerHRR = hrr 'heat release at actuation


        '        'TwoZones(fireroom) = False
        '        'UpperVolume(fireroom, 1) = RoomVolume(fireroom)
        '    End If
        'End If
    End Sub
    Sub FED_CO_13571(ByVal room As Integer)

        Dim i, FEDFlag As Short
        Dim O2Sum As Single
        Dim COSum As Double
        Dim RMV2, RMV1, VCO2 As Single

        O2Sum = 0
        COSum = 0

        For i = 1 To NumberTimeSteps + 1
            If tim(i, 1) >= StartOccupied And tim(i, 1) <= EndOccupied Then
                If layerheight(room, i) <= MonitorHeight Then
                    'exposed to upper layer
                    If O2VolumeFraction(room, i, 1) < 0.13 Then 'ignore O2 effect where % volume > 13%
                        O2Sum = O2Sum + Timestep / 60 / (Exp(8.13 - 0.54 * (20.9 - (O2VolumeFraction(room, i, 1) * 100 + O2VolumeFraction(room, i + 1, 1) * 100) / 2)))
                    End If

                    If CO2VolumeFraction(room, i, 1) > 0 Then 'include hyperventilation for all co2 as per iso 13571:2012
                        RMV1 = Exp(0.2 * CO2VolumeFraction(room, i, 1) * 100)
                        RMV2 = Exp(0.2 * CO2VolumeFraction(room, i + 1, 1) * 100)
                    Else
                        RMV1 = 1
                        RMV2 = 1
                    End If
                    VCO2 = (RMV1 + RMV2) / 2

                    If COVolumeFraction(room, i, 1) > COVolumeFraction(room, 1, 1) Then
                        COSum = COSum + Timestep / (60 * 35000) * VCO2 * (((COVolumeFraction(room, i, 1) + COVolumeFraction(room, i + 1, 1)) / 2) * 10 ^ 6)
                    End If

                Else
                    'exposed to lower layer
                    If O2VolumeFraction(room, i, 2) < 0.13 Then 'ignore O2 effect where % volume > 13%
                        O2Sum = O2Sum + Timestep / 60 / (Exp(8.13 - 0.54 * (20.9 - (O2VolumeFraction(room, i, 2) * 100 + O2VolumeFraction(room, i + 1, 2) * 100) / 2)))
                    End If

                    If CO2VolumeFraction(room, i, 2) > 0 Then 'include hyperventilation for all co2 as per iso 13571:2012
                        RMV1 = Exp(0.2 * CO2VolumeFraction(room, i, 2) * 100)
                        RMV2 = Exp(0.2 * CO2VolumeFraction(room, i + 1, 2) * 100)
                    Else
                        RMV1 = 1
                        RMV2 = 1
                    End If
                    VCO2 = (RMV1 + RMV2) / 2

                End If

                If COVolumeFraction(room, i, 2) > COVolumeFraction(room, 1, 2) Then
                    COSum = COSum + Timestep / (60 * 35000) * VCO2 * (((COVolumeFraction(room, i, 2) + COVolumeFraction(room, i + 1, 2)) / 2) * 10 ^ 6)
                End If

            End If

            If O2Sum + COSum < 1 Then
                FEDSum(room, i) = O2Sum + COSum
            Else
                FEDSum(room, i) = 1
            End If

            If i = 1 Then FEDFlag = 0

            If FEDEndPoint <= FEDSum(fireroom, i) And FEDFlag = 0 Then
                frmEndPoints.lblFED.Text = "FED(CO) (incap) Exceeded " & VB6.Format(FEDEndPoint) & " at " & VB6.Format(tim(i, 1), "0.0") & " Seconds."
                FEDFlag = 1
            End If

        Next i

        If NumberTimeSteps < 1 Then Exit Sub

        If tim(NumberTimeSteps + 1, 1) >= StartOccupied And tim(NumberTimeSteps + 1, 1) <= EndOccupied Then
            If layerheight(room, NumberTimeSteps + 1) <= MonitorHeight Then
                If O2VolumeFraction(room, i, 1) < 0.13 Then
                    O2Sum = O2Sum + Timestep / 60 / (Exp(7.98 - 0.528 * (O2VolumeFraction(room, 1, 1) * 100 - O2VolumeFraction(room, NumberTimeSteps + 1, 1) * 100)))
                End If

                If CO2VolumeFraction(room, i, 1) > 0 Then 'include hyperventilation for all co2 as per iso 13571:2012
                    RMV1 = Exp(0.2 * CO2VolumeFraction(room, NumberTimeSteps + 1, 1) * 100)
                Else
                    RMV1 = 1
                End If
                VCO2 = RMV1

                If COVolumeFraction(room, i, 1) > COVolumeFraction(room, 1, 1) Then
                    COSum = COSum + Timestep / (60 * 35000) * VCO2 * (COVolumeFraction(room, NumberTimeSteps + 1, 1) * 10 ^ 6) * RMV1
                End If

            Else
                If O2VolumeFraction(room, i, 2) < 0.13 Then
                    O2Sum = O2Sum + Timestep / 60 / (Exp(7.98 - 0.528 * (O2VolumeFraction(room, 1, 2) * 100 - O2VolumeFraction(room, NumberTimeSteps + 1, 2) * 100)))
                End If

                If CO2VolumeFraction(room, i, 2) > 0 Then 'include hyperventilation for all co2 as per iso 13571:2012
                    RMV1 = Exp(0.2 * CO2VolumeFraction(room, NumberTimeSteps + 1, 2) * 100)
                Else
                    RMV1 = 1
                End If
                VCO2 = RMV1

                If COVolumeFraction(room, i, 2) > COVolumeFraction(room, 1, 2) Then
                    COSum = COSum + Timestep / (60 * 35000) * VCO2 * (COVolumeFraction(room, NumberTimeSteps + 1, 2) * 10 ^ 6) * RMV1
                End If

            End If
        End If

        If O2Sum + COSum < 4 Then
            FEDSum(room, NumberTimeSteps + 1) = O2Sum + COSum
        Else
            FEDSum(room, NumberTimeSteps + 1) = 4
        End If

        If room = fireroom Then
            If FEDEndPoint <= FEDSum(fireroom, NumberTimeSteps + 1) And FEDFlag = 0 Then
                frmEndPoints.lblFED.Text = "FED(CO) exceeded " & VB6.Format(FEDEndPoint) & " at " & VB6.Format(tim(i, 1), "0.0") & " Seconds."
                FEDFlag = 1
            End If
        End If
    End Sub
    Sub FED_Toxicity(ByVal room As Integer)
        '* ======================================================
        '*  This procedure calculates the Fractional Effective
        '*  Dose causing incapacitation due to oxygen hypoxia and carbon
        '*  monoxide and HCN exposure effects.
        '*  includes CO2 hyperventilation effect on CO and HCN uptake
        '*  allows for activity level
        '*
        '*  Revised: Colleen Wade 13/1/98
        '   edited 14/10/12 
        '* ======================================================

        Dim i, FEDFlag As Short
        Dim O2Sum As Single
        Dim COSum, HCNsum As Double
        Dim RMV2, RMVo, RMV1, COHb, VCO2, COppm As Single

        On Error GoTo toxicityhandler
        If IsNothing(tim) Then Exit Sub
        If NumberTimeSteps < 1 Then Exit Sub

        'Get activity level
        If frmOptions1.optAtRest.Checked = True Then Activity = "Rest"
        If frmOptions1.optLightWork.Checked = True Then Activity = "Light"
        If frmOptions1.optHeavyWork.Checked = True Then Activity = "Heavy"

        'Get values of COHb and RMVo appropriate to activity level
        Select Case Activity
            Case "Rest"
                COHb = 40 '% COHb dose for incapacitation
                RMVo = 8.5 'breathing rate l/min
            Case "Light"
                COHb = 30 '% COHb dose for incapacitation
                RMVo = 25 'breathing rate l/min
            Case "Heavy"
                COHb = 20 '% COHb dose for incapacitation
                RMVo = 50 'breathing rate l/min
        End Select

        O2Sum = 0
        COSum = 0
        HCNsum = 0

        For i = 1 To NumberTimeSteps
            If tim(i, 1) >= StartOccupied And tim(i, 1) <= EndOccupied Then 'user has option to select the time interval over which the calculation is done
                If layerheight(room, i) <= MonitorHeight Then
                    'exposed to upper layer
                    If O2VolumeFraction(room, i, 1) < 0.13 Then 'only do if O2 < 13% vol
                        O2Sum = O2Sum + Timestep / 60 / (Exp(8.13 - 0.54 * (20.9 - (O2VolumeFraction(room, i, 1) * 100 + O2VolumeFraction(room, i + 1, 1) * 100) / 2)))
                    End If

                    RMV1 = Exp(CO2VolumeFraction(room, i, 1) * 100 / 5)
                    RMV2 = Exp(CO2VolumeFraction(room, i + 1, 1) * 100 / 5)
                    VCO2 = (RMV1 + RMV2) / 2  'sfpe 2-122 eqn 15

                    If COVolumeFraction(room, i, 1) > COVolumeFraction(room, 1, 1) Then
                        COppm = ((COVolumeFraction(room, i, 1) + COVolumeFraction(room, i + 1, 1)) / 2) * 10 ^ 6 'CO in ppm
                        COSum = COSum + Timestep / 60 * VCO2 * 0.00003317 * RMVo / COHb * (COppm) ^ 1.036

                    End If

                    'hcn
                    If HCNVolumeFraction(room, i, 1) > HCNVolumeFraction(room, 1, 1) Then 'if higher than ambient level (~0)
                        If (HCNVolumeFraction(room, i, 1) + HCNVolumeFraction(room, i + 1, 1)) / 2 * 10 ^ 6 > 80 Then
                            HCNsum = HCNsum + Timestep / 60 * VCO2 * Exp(5.396 - 0.023 * (((HCNVolumeFraction(room, i, 1) + HCNVolumeFraction(room, i + 1, 1)) / 2) * 10 ^ 6))
                        End If
                    End If

                Else
                    'exposed to lower layer
                    If O2VolumeFraction(room, i, 2) < 0.13 Then
                        O2Sum = O2Sum + Timestep / 60 / (Exp(8.13 - 0.54 * (20.9 - (O2VolumeFraction(room, i, 2) * 100 + O2VolumeFraction(room, i + 1, 2) * 100) / 2)))
                    End If

                    RMV1 = Exp(CO2VolumeFraction(room, i, 2) * 100 / 5)
                    RMV2 = Exp(CO2VolumeFraction(room, i + 1, 2) * 100 / 5)
                    VCO2 = (RMV1 + RMV2) / 2  'sfpe 2-122 eqn 15

                    If COVolumeFraction(room, i, 2) > COVolumeFraction(room, 1, 2) Then
                        COppm = ((COVolumeFraction(room, i, 2) + COVolumeFraction(room, i + 1, 2)) / 2) * 10 ^ 6
                        COSum = COSum + Timestep / 60 * VCO2 * 0.00003317 * RMVo / COHb * (COppm) ^ 1.036
                    End If

                    If HCNVolumeFraction(room, i, 2) > HCNVolumeFraction(room, 1, 2) Then
                        If ((HCNVolumeFraction(room, i, 2) + HCNVolumeFraction(room, i + 1, 2)) / 2) * 10 ^ 6 > 80 Then
                            HCNsum = HCNsum + Timestep / 60 * VCO2 * Exp(5.396 - 0.023 * (((HCNVolumeFraction(room, i, 2) + HCNVolumeFraction(room, i + 1, 2)) / 2) * 10 ^ 6))
                        End If
                    End If
                End If
            End If

            If O2Sum + COSum < 1 Then
                FEDSum(room, i) = O2Sum + COSum + HCNsum
            Else
                FEDSum(room, i) = 1
            End If

            If i = 1 Then FEDFlag = 0

            If FEDEndPoint <= FEDSum(fireroom, i) And FEDFlag = 0 Then
                frmEndPoints.lblFED.Text = "FED asphyxiant gases (incap) Exceeded " & VB6.Format(FEDEndPoint) & " at " & VB6.Format(tim(i, 1), "0.0") & " Seconds."
                FEDFlag = 1
            End If

        Next i

        'following is just for the very last timstep, otherwise same as above.
        If tim(NumberTimeSteps + 1, 1) >= StartOccupied And tim(NumberTimeSteps + 1, 1) <= EndOccupied Then
            If layerheight(room, NumberTimeSteps + 1) <= MonitorHeight Then
                If O2VolumeFraction(room, i, 1) < 0.13 Then
                    O2Sum = O2Sum + Timestep / 60 / (Exp(8.13 - 0.54 * (20.9 - O2VolumeFraction(room, NumberTimeSteps + 1, 1) * 100)))
                End If

                VCO2 = Exp(CO2VolumeFraction(room, NumberTimeSteps + 1, 1) * 100 / 5)

                If COVolumeFraction(room, i, 1) > COVolumeFraction(room, 1, 1) Then
                    COSum = COSum + Timestep / 60 * VCO2 * 0.00003317 * RMVo / COHb * ((COVolumeFraction(room, NumberTimeSteps + 1, 1) * 10 ^ 6)) ^ 1.036
                End If
                If HCNVolumeFraction(room, i, 1) > HCNVolumeFraction(room, 1, 1) Then
                    If HCNVolumeFraction(room, i, 1) * 10 ^ 6 > 80 Then
                        HCNsum = HCNsum + Timestep / 60 * VCO2 * Exp(5.396 - 0.023 * (HCNVolumeFraction(room, i, 1) * 10 ^ 6))
                    End If
                End If
            Else
                If O2VolumeFraction(room, i, 2) < 0.13 Then
                    O2Sum = O2Sum + Timestep / 60 / (Exp(8.13 - 0.54 * (20.9 - O2VolumeFraction(room, NumberTimeSteps + 1, 2) * 100)))
                End If

                VCO2 = Exp(CO2VolumeFraction(room, NumberTimeSteps + 1, 2) * 100 / 5)

                If COVolumeFraction(room, i, 2) > COVolumeFraction(room, 1, 2) Then
                    COSum = COSum + Timestep / 60 * VCO2 * 0.00003317 * RMVo / COHb * ((COVolumeFraction(room, NumberTimeSteps + 1, 2) * 10 ^ 6)) ^ 1.036
                End If
                If HCNVolumeFraction(room, i, 2) > HCNVolumeFraction(room, 1, 2) Then
                    If HCNVolumeFraction(room, i, 2) * 10 ^ 6 > 80 Then
                        HCNsum = HCNsum + Timestep / 60 * VCO2 * Exp(5.396 - 0.023 * (((HCNVolumeFraction(room, i, 2))) * 10 ^ 6))
                    End If
                End If
            End If
        End If

        If O2Sum + COSum + HCNsum < 4 Then
            FEDSum(room, NumberTimeSteps + 1) = O2Sum + COSum + HCNsum
        Else
            FEDSum(room, NumberTimeSteps + 1) = 4
        End If

        If room = fireroom Then
            If FEDEndPoint <= FEDSum(fireroom, NumberTimeSteps + 1) And FEDFlag = 0 Then
                frmEndPoints.lblFED.Text = "FED Asphyxiant Gases (incap) Exceeded " & VB6.Format(FEDEndPoint) & " at " & VB6.Format(tim(i, 1), "0.0") & " Seconds."
                FEDFlag = 1
            End If
        End If

        Exit Sub

toxicityhandler:
        Exit Sub

    End Sub

    Function File_Errors(ByRef errval As Short) As Short
        '*  ===================================================================
        '*  This function is part of the error handling routine for file access
        '*  ===================================================================

        Dim msgtype, response As Short
        Dim msg As String

        msgtype = MB_EXCLAIM

        Select Case errval
            Case err_deviceunavailable
                msg = "That device appears unavailable."
                msgtype = MB_EXCLAIM + 4
            Case err_disknotready
                msg = "Insert a disk in the drive and close the door."
            Case err_deviceio
                msg = "internal disk error."
                msgtype = MB_EXCLAIM + 4
            Case err_diskfull
                msg = "Disk is full. Continue?"
                msgtype = 35
            Case err_badfilename, err_badfilenameornumber
                msg = "That filename is illegal."
            Case err_pathdoesnotexist
                msg = "That path doesn't exist."
            Case err_badfilemode
                msg = "Can't open your file for that type of access."
            Case err_filealreadyopen
                msg = "This file is already open."
            Case err_inputpastendoffile
                msg = "This file has a non standard end of file marker,"
                msg = msg & " or an attempt was made to read beyond"
                msg = msg & " the end of file marker."
                FileClose(1)
                msgtype = MB_EXCLAIM + 2
            Case err_overflow
                msg = "Overflow. Simulation will be Stopped."
                'Stop
                Exit Function
            Case Else
                File_Errors = 3
                msg = "An error has occurred reading the input data file."
                FileClose(1)
        End Select

        response = MsgBox(msg, msgtype, ProgramTitle)

        Select Case response
            Case 1, 4 'ok, retry buttons
                File_Errors = 0
                Exit Function
            Case 5 'ignore button
                File_Errors = 1
                Exit Function
            Case 2, 3 'cancel, abort buttons
                File_Errors = 2
                Exit Function
            Case Else
                File_Errors = 3
                Exit Function
        End Select

    End Function

    Function Fire_Diameter(ByVal massloss As Double, ByVal MLUA As Double) As Double
        '*  ==================================================================
        '*  This function returns the value of the fire diameter for use in
        '*  the plume equations.
        '*
        '*  Arguments passed to the function:
        '*      MassLoss    =   mass loss rate of the fuel (kg/s)
        '*
        '*  Other variables used:
        '*      MLUnitArea        =   mass burning rate per unit area (kg/s per m2)
        '*  ==================================================================

        If massloss < 0 Then massloss = 0

        'Fire_Diameter = Sqrt(4 * massloss / (PI * MLUnitArea))
        Fire_Diameter = Sqrt(4 * massloss / (PI * MLUA))

    End Function

    Function Flame_Height(ByVal diameter As Double, ByVal Q As Single) As Single
        '*  ================================================================
        '*  This function returns the height of the flame above the
        '*  virtual source using Heskestad's equation
        '*
        '*  Arguments passed to the function:
        '*      Diameter    =   fire diameter (m)
        '*      Q           =   heat release (kW)
        '*  ================================================================

        Flame_Height = -1.02 * diameter + 0.23 * Q ^ (2 / 5)

    End Function
    Function Get_HRR_powerlaw(ByVal tim As Double, ByRef Q As Double) As Double
        '*  ====================================================================
        '*  Returns the heat release rate (kW) at a given time (sec) for a single
        '*  object.
        '*
        '*  ID = the burning object number
        '*  tim = Time (seconds)
        '*  Q = hrr to use in plume correlation (kW/m2)
        '*
        '*  ====================================================================

        Dim hrr, hrr_fo, sphrr As Double
        Static post As Boolean

        If tim = 0 Then
            post = frmOptions1.optPostFlashover.Checked 'remember this value
        End If

        If Flashover = True And post = True Then
            'this never get called?
            hrr = MassLoss_Total(tim) * NewHoC_fuel * 1000
            Get_HRR_powerlaw = hrr
            Q = hrr
            Exit Function
        End If

        'check to see if any fuel left
        'If frmOptions1.optFuelLimited.Checked = True And post = False And stepcount > 1 Then
        If post = False And stepcount > 1 Then
            If TotalFuel(stepcount - 1) > InitialFuelMass Then
                Get_HRR_powerlaw = 0
                Q = 0
                'Q = Qlast
                Exit Function
            End If
        End If

        If sprink_mode = 2 Then 'suppression
            If SprinklerFlag = 1 Then
                'sprinkler has operated

                If WaterSprayDensity(fireroom) > 0 Then
                    sphrr = SprinklerHRR * Exp(-(tim - SprinklerTime) / (3 * (WaterSprayDensity(fireroom) / 60) ^ (-1.85)))
                End If
            End If
        ElseIf sprink_mode = 1 Then  'control
            If SprinklerFlag = 1 Then
                'sprinkler has operated

                sphrr = SprinklerHRR
            End If
        End If

        If useT2fire = True Then
            If Flashover = True And VM2 = True Then
                'increase the hrr to the peak HRR over 15 sec period as specified VM2
                If tim < flashover_time + 15 Then
                    hrr_fo = AlphaT * flashover_time ^ 2
                    If PeakHRR < hrr_fo Then hrr_fo = PeakHRR
                    hrr = hrr_fo + (tim - flashover_time) / 15 * (PeakHRR - hrr_fo)
                Else
                    hrr = PeakHRR
                End If
            Else
                hrr = AlphaT * tim ^ 2
            End If
        Else
            If Flashover = True And VM2 = True Then
                If tim < flashover_time + 15 Then
                    hrr_fo = AlphaT * flashover_time ^ 3 * StoreHeight
                    If PeakHRR < hrr_fo Then hrr_fo = PeakHRR
                    hrr = hrr_fo + (tim - flashover_time) / 15 * (PeakHRR - hrr_fo)
                Else
                    hrr = PeakHRR
                End If
            Else
                hrr = AlphaT * tim ^ 3 * StoreHeight
            End If
        End If

        If PeakHRR < hrr Then
            hrr = PeakHRR
        End If


        'for a sprinkler fire, use the lesser of the theoretical hrr and the sprinkler controlled rate
        If SprinklerFlag = 1 Then
            If sprink_mode = 0 Then sphrr = hrr
            If sprink_mode = 2 And WaterSprayDensity(fireroom) = 0 Then sphrr = hrr 'suppression mode
            If sprink_mode = 1 And sphrr > hrr Then sphrr = hrr 'control mode, added 29/09/2011
            hrr = sphrr
        End If

        If tim = 0 And hrr < 0.000001 Then hrr = 0.000001

        Get_HRR_powerlaw = hrr

        Q = hrr

    End Function
    '    Function Get_HRR(ByVal id As Integer, ByVal tim As Double, ByRef Q As Double, ByRef Qburner As Double, ByRef QFloor As Double, ByRef QWall As Double, ByRef QCeiling As Double) As Double
    '        '*  ====================================================================
    '        '*  Returns the heat release rate (kW) at a given time (sec) for a single
    '        '*  object. Linearly interpolates between data points.
    '        '*
    '        '*  ID = the burning object number
    '        '*  tim = Time (seconds)
    '        '*  Q = hrr to use in plume correlation (kW/m2)
    '        '*
    '        '*  ====================================================================

    '        Dim i As Short
    '        Dim deltaQ, FSA As Double
    '        Dim hrr, b, A, C, sphrr As Double
    '        Static QClast, QFlast, HRRlast, timlast, Qlast, QWlast, QBlast As Double
    '        Static IDLast As Short
    '        Static growth, post, quintiere As Boolean

    '        If FuelResponseEffects = True Then
    '            'won't work for multiple objects
    '            'only need free burn MLT here 
    '            'well ventilated
    '            Q = 1000 * EnergyYield(id) * FuelBurningRate(0, fireroom, stepcount)
    '            'kW freeburn
    '            ' If Q > 1500 Then Stop
    '            hrr = Q
    '            Get_HRR = Q

    '            'Exit Function
    '        End If


    '        If tim = 0 Then
    '            HRRlast = 0
    '            Qlast = 0
    '            QBlast = 0
    '            QWlast = 0
    '            QClast = 0
    '            QFlast = 0
    '            IDLast = id '2011.26
    '        End If

    '        If tim = timlast And id = IDLast And tim > 0 Then
    '            Get_HRR = HRRlast
    '            Q = Qlast
    '            Qburner = QBlast
    '            QWall = QWlast
    '            QCeiling = QClast
    '            QFloor = QFlast
    '            Exit Function
    '        End If

    '        If tim = 0 Then
    '            post = frmOptions1.optPostFlashover.Checked 'remember this value
    '            growth = frmOptions1.optRCNone.Checked
    '            quintiere = frmOptions1.optQuintiere.Checked
    '        End If

    '        If Flashover = True And post = True Then 'using the wood crib post-flashover model

    '            'only need to call this once per timestep when using the wood crib PF model
    '            If IEEERemainder(tim, Timestep) = 0 Then
    '                Dim dummy As Double 'returns contribution from wall, ceiling

    '                If frmCLT.chkWoodIntegralModel.Checked = True Then 'using MS & JQ integral model for timber 
    '                    hrr = MassLoss_Total_pluswood(tim, dummy, dummy) * NewHoC_fuel * 1000 'spearpoint quintiere
    '                Else
    '                    hrr = MassLoss_Total(tim) * NewHoC_fuel * 1000
    '                End If

    '                Get_HRR = hrr
    '                timlast = tim
    '                HRRlast = hrr
    '                IDLast = id
    '                Q = hrr
    '                Qlast = Q
    '                QBlast = Qburner
    '                QWlast = QWall
    '                QClast = QCeiling
    '                QFlast = QFloor
    '            Else
    '                Get_HRR = HRRlast
    '                Q = Qlast
    '                Qburner = QBlast
    '                QWall = QWlast
    '                QCeiling = QClast
    '                QFloor = QFlast
    '            End If

    '            Exit Function
    '        End If

    '        If usepowerlawdesignfire = True Then
    '            If id = 1 Then '1 item only
    '                hrr = Get_HRR_powerlaw(tim, Q)
    '            Else
    '                hrr = 0
    '                Q = 0
    '            End If
    '            Get_HRR = hrr
    '            timlast = tim
    '            HRRlast = hrr
    '            IDLast = id
    '            Q = hrr
    '            Qlast = Q
    '            QBlast = Qburner
    '            QWlast = QWall
    '            QClast = QCeiling
    '            QFlast = QFloor

    '            Exit Function
    '        End If

    '        'check to see if any fuel left
    '        If post = False And stepcount > 1 Then 'if not using wood crib post flashover model
    '            If growth = True Then  'and not a room corner test 

    '                If TotalFuel(stepcount - 1) > InitialFuelMass Then

    '                    'If useCLTmodel = True And TotalFuel(stepcount - 1) < fuelmasswithCLT Then
    '                    If useCLTmodel = True Then
    '                        'continue
    '                    Else
    '                        'no fuel left
    '                        Get_HRR = 0
    '                            Q = Qlast
    '                            Qburner = QBlast
    '                            QWall = QWlast
    '                            QCeiling = QClast
    '                            QFloor = QFlast
    '                            IDLast = id '2011.26

    '                            Exit Function

    '                        End If
    '                    End If
    '                End If
    '        End If

    '        If sprink_mode = 2 Then 'sprinkler suppression
    '            If SprinklerFlag = 1 Then
    '                'sprinkler has operated
    '                If id > 1 Then Exit Function 'only uses fire object no. 1
    '                If WaterSprayDensity(fireroom) > 0 Then
    '                    sphrr = SprinklerHRR * Exp(-(tim - SprinklerTime) / (3 * (WaterSprayDensity(fireroom) / 60) ^ (-1.85)))
    '                End If
    '            End If
    '        ElseIf sprink_mode = 1 Then  'sprinkler control
    '            If SprinklerFlag = 1 Then
    '                'sprinkler has operated
    '                If id > 1 Then Exit Function 'only uses fire object no. 1
    '                sphrr = SprinklerHRR
    '            End If
    '        End If

    '        If growth = True Then
    '            'this is not a room corner test
    'start:
    '            'read in the hrr versus time data for the object
    '            If FuelResponseEffects = False Then
    '                For i = 1 To NumberDataPoints(id) - 1

    '                    If i + 1 > 500 Then 'max 500 data points allowed
    '                        hrr = HRRlast
    '                        Get_HRR = hrr
    '                        timlast = tim
    '                        HRRlast = hrr
    '                        IDLast = id
    '                        Q = hrr
    '                        Qlast = Q
    '                        QBlast = Qburner
    '                        QWlast = QWall
    '                        QClast = QCeiling
    '                        QFlast = QFloor
    '                        Exit Function
    '                    End If

    '                    If tim >= (HeatReleaseData(1, i, id) + ObjectIgnTime(id)) Then
    '                        If tim < (HeatReleaseData(1, i + 1, id) + ObjectIgnTime(id)) Then
    '                            A = tim - (HeatReleaseData(1, i, id) + ObjectIgnTime(id))
    '                            b = HeatReleaseData(1, i + 1, id) - HeatReleaseData(1, i, id) 'sec
    '                            C = HeatReleaseData(2, i + 1, id) - HeatReleaseData(2, i, id) 'kw
    '                            hrr = HeatReleaseData(2, i, id) + (A / b) * C
    '                            If hrr < 0 Then hrr = 0

    '                            'for a sprinkler fire, use the lesser of the theoretical hrr and the sprinkler controlled rate
    '                            If SprinklerFlag = 1 Then
    '                                If sprink_mode = 0 Then sphrr = hrr
    '                                If sprink_mode = 2 And WaterSprayDensity(fireroom) = 0 Then sphrr = hrr 'suppression mode
    '                                If sprink_mode = 1 And sphrr > hrr Then sphrr = hrr 'control mode, added 29/09/2011
    '                                hrr = sphrr
    '                            End If

    '                            If tim = 0 And hrr < 0.000001 Then hrr = 0.000001

    '                            If useCLTmodel = True And post = False Then 'CLT mod on and wood crib PF model off. 
    '                                QWall1 = wall_char(stepcount, 1) * WallEffectiveHeatofCombustion(fireroom) * 1000
    '                                QCeiling1 = ceil_char(stepcount, 1) * CeilingEffectiveHeatofCombustion(fireroom) * 1000
    '                                Get_HRR = hrr

    '                                timlast = tim
    '                                HRRlast = hrr
    '                                IDLast = id
    '                                Q = hrr 'plume?
    '                                Qlast = Q
    '                                QBlast = Qburner
    '                                QWlast = QWall1
    '                                QClast = QCeiling1
    '                                QFlast = QFloor1

    '                                Exit Function

    '                            End If

    '                            Get_HRR = hrr
    '                            timlast = tim
    '                            HRRlast = hrr
    '                            IDLast = id
    '                            Q = hrr
    '                            Qlast = Q
    '                            QBlast = Qburner
    '                            QWlast = QWall
    '                            QClast = QCeiling
    '                            QFlast = QFloor

    '                            Exit Function
    '                        End If
    '                    End If
    '                Next i
    '            Else
    '                'fuel response effects is true
    '                If useCLTmodel = True And post = False Then 'CLT mod on and wood crib PF model off. 
    '                    QWall1 = wall_char(stepcount - 1, 1) * WallEffectiveHeatofCombustion(fireroom) * 1000
    '                    QCeiling1 = ceil_char(stepcount - 1, 1) * CeilingEffectiveHeatofCombustion(fireroom) * 1000
    '                    Get_HRR = hrr
    '                    Qburner = hrr
    '                    timlast = tim
    '                    HRRlast = hrr
    '                    IDLast = id
    '                    Q = hrr
    '                    Qlast = Q
    '                    QBlast = Qburner
    '                    QWlast = QWall1
    '                    QClast = QCeiling1
    '                    QFlast = QFloor1
    '                    'If QWall1 > 1000 Then Stop
    '                    Exit Function

    '                End If
    '            End If

    '            'check to see if time corresponds to last data point
    '            If NumberDataPoints(id) <> 0 Then
    '                    If tim >= (HeatReleaseData(1, NumberDataPoints(id), id) + ObjectIgnTime(id)) Then
    '                        hrr = HeatReleaseData(2, NumberDataPoints(id), id)
    '                        If SprinklerFlag = 1 Then
    '                            If sprink_mode = 0 Then sphrr = hrr
    '                            If sprink_mode = 2 And WaterSprayDensity(fireroom) = 0 Then sphrr = hrr 'suppression mode
    '                            If sprink_mode = 1 And sphrr > hrr Then sphrr = hrr 'control mode, added 29/09/2011

    '                            hrr = sphrr
    '                        End If
    '                    End If
    '                End If

    '            Else
    '                'this is a room corner test
    '                If id = Max(1, burner_id) Then
    '                If quintiere = True Then
    '                    'quintiere's room corner model
    '                    hrr = HRR_Quintiere2(id, tim, Q, Qburner, QFloor, QWall, QCeiling) 'hrr includes burner + linings
    '                Else
    '                    'Karlsson's room corner Model A
    '                    hrr = HRR_ModelA(id, tim, Q)
    '                End If
    '            Else
    '                'this item has not ignited the wall
    '                GoTo start
    '            End If

    '            'for a sprinkler fire, use the lesser of the theoretical hrr and the sprinkler controlled rate
    '            If SprinklerFlag = 1 Then
    '                If sprink_mode = 0 Then sphrr = hrr
    '                If sprink_mode = 2 And WaterSprayDensity(fireroom) = 0 Then sphrr = hrr 'suppression mode
    '                If sprink_mode = 1 And sphrr > hrr Then sphrr = hrr 'control mode, added 29/09/2011
    '                hrr = sphrr
    '            End If

    '        End If
    '        If flagstop = 1 Then Exit Function


    '        Get_HRR = hrr

    '        If id = Max(1, burner_id) Then '21/11/2011
    '            timlast = tim
    '            HRRlast = hrr
    '            IDLast = id
    '            Qlast = Q
    '            QBlast = Qburner
    '            QWlast = QWall
    '            QClast = QCeiling
    '            QFlast = QFloor
    '        End If

    '    End Function
    Function Get_HRR(ByVal id As Integer, ByVal tim As Double, ByRef Q As Double, ByRef Qburner As Double, ByRef QFloor As Double, ByRef QWall As Double, ByRef QCeiling As Double) As Double
        '*  ====================================================================
        '*  Returns the heat release rate (kW) at a given time (sec) for a single
        '*  object. Linearly interpolates between data points.
        '*
        '*  ID = the burning object number
        '*  tim = Time (seconds)
        '*  Q = hrr to use in plume correlation (kW/m2)
        '*
        '*  ====================================================================

        If FuelResponseEffects = True Then
            'won't work for multiple objects
            'only need free burn MLT here 
            'well ventilated
            Q = 1000 * EnergyYield(id) * FuelBurningRate(0, fireroom, stepcount) 'kW freeburn

            Dim mwall, mceiling, temp As Double 'returns contribution from wall, ceiling kg/s
            If useCLTmodel = True And IntegralModel = True Then 'using MS & JQ integral model for timber 
                ' hrr = MassLoss_Total_pluswood(tim, mwall, mceiling) * NewHoC_fuel * 1000 'spearpoint quintiere
            ElseIf useCLTmodel = True And KineticModel = True Then
                'If FuelResponseEffects = False Then
                temp = MassLoss_Total_Kinetic(tim, mwall, mceiling)

                Q = FuelMassLossRate(stepcount, fireroom) * 1000 * EnergyYield(1)
                QWall = WallEffectiveHeatofCombustion(fireroom) * mwall * 1000
                QCeiling = CeilingEffectiveHeatofCombustion(fireroom) * mceiling * 1000

            End If

            Get_HRR = Q

            Exit Function
        End If

        'broke bit
        '        If FuelResponseEffects = True Then
        '            'won't work for multiple objects
        '            'only need free burn MLT here 
        '            'well ventilated
        '            Q = 1000 * EnergyYield(id) * FuelBurningRate(0, fireroom, stepcount)
        '            'kW freeburn
        '            ' If Q > 1500 Then Stop
        '           hrr = Q
        '            Get_HRR = Q

        '            'Exit Function
        '        End If

        Dim i As Short
        Dim deltaQ, FSA As Double
        Dim hrr, b, A, C, sphrr As Double
        Static QClast, QFlast, HRRlast, timlast, Qlast, QWlast, QBlast As Double
        Static IDLast As Short
        Static growth, post, quintiere As Boolean

        If tim = 0 Then
            HRRlast = 0
            Qlast = 0
            QBlast = 0
            QWlast = 0
            QClast = 0
            QFlast = 0
            IDLast = id '2011.26
        End If

        If tim = timlast And id = IDLast And tim > 0 Then
            Get_HRR = HRRlast
            Q = Qlast
            Qburner = QBlast
            QWall = QWlast
            QCeiling = QClast
            QFloor = QFlast
            Exit Function
        End If

        If tim = 0 Then
            post = frmOptions1.optPostFlashover.Checked 'remember this value
            growth = frmOptions1.optRCNone.Checked
            quintiere = frmOptions1.optQuintiere.Checked
        End If

        If Flashover = True And post = True Then 'using the wood crib post-flashover model

            'only need to call this once per timestep when using the wood crib PF model
            If IEEERemainder(tim, Timestep) = 0 Then

                Dim mwall, mceiling As Double 'returns contribution from wall, ceiling kg/s
                If useCLTmodel = True And IntegralModel = True Then 'using MS & JQ integral model for timber 
                    hrr = MassLoss_Total_pluswood(tim, mwall, mceiling) * NewHoC_fuel * 1000 'spearpoint quintiere
                ElseIf useCLTmodel = True And kineticModel = True Then
                    If FuelResponseEffects = False Then
                        hrr = MassLoss_Total_Kinetic(tim, mwall, mceiling) * NewHoC_fuel * 1000
                    Else
                        'need the contents as well
                        hrr = WallEffectiveHeatofCombustion(fireroom) * mwall * 1000 + CeilingEffectiveHeatofCombustion(fireroom) * mceiling * 1000
                    End If

                Else
                        hrr = MassLoss_Total(tim) * NewHoC_fuel * 1000
                End If

                Get_HRR = hrr
                timlast = tim
                HRRlast = hrr
                IDLast = id
                Q = hrr
                Qlast = Q
                QBlast = Qburner
                QWlast = QWall
                QClast = QCeiling
                QFlast = QFloor
            Else
                Get_HRR = HRRlast
                Q = Qlast
                Qburner = QBlast
                QWall = QWlast
                QCeiling = QClast
                QFloor = QFlast
            End If

            Exit Function

            ' ElseIf Flashover = True And useCLTmodel = True And IntegralModel = True Then



        End If

        If usepowerlawdesignfire = True Then
            If id = 1 Then '1 item only
                hrr = Get_HRR_powerlaw(tim, Q)
            Else
                hrr = 0
                Q = 0
            End If
            Get_HRR = hrr
            timlast = tim
            HRRlast = hrr
            IDLast = id
            Q = hrr
            Qlast = Q
            QBlast = Qburner
            QWlast = QWall
            QClast = QCeiling
            QFlast = QFloor

            Exit Function
        End If

        'check to see if any fuel left
        If post = False And stepcount > 1 Then 'if not using wood crib post flashover model
            If growth = True Then  'and not a room corner test 

                If TotalFuel(stepcount - 1) > InitialFuelMass Then
                    If useCLTmodel = True And TotalFuel(stepcount - 1) < fuelmasswithCLT Then
                        'continue
                    Else
                        'no fuel left
                        Get_HRR = 0
                        Q = Qlast
                        Qburner = QBlast
                        QWall = QWlast
                        QCeiling = QClast
                        QFloor = QFlast
                        IDLast = id '2011.26

                        Exit Function

                    End If
                End If
            End If
        End If

        'broke bit
        '        'check to see if any fuel left
        '        If post = False And stepcount > 1 Then 'if not using wood crib post flashover model
        '            If growth = True Then  'and not a room corner test 

        '                If TotalFuel(stepcount - 1) > InitialFuelMass Then

        '                    'If useCLTmodel = True And TotalFuel(stepcount - 1) < fuelmasswithCLT Then
        '                    If useCLTmodel = True Then
        '                        'continue
        '                    Else
        '                        'no fuel left
        '                        Get_HRR = 0
        '                            Q = Qlast
        '                            Qburner = QBlast
        '                            QWall = QWlast
        '                            QCeiling = QClast
        '                            QFloor = QFlast
        '                            IDLast = id '2011.26

        '                            Exit Function

        '                        End If
        '                    End If
        '                End If
        '        End If


        If sprink_mode = 2 Then 'sprinkler suppression
            If SprinklerFlag = 1 Then
                'sprinkler has operated
                If id > 1 Then Exit Function 'only uses fire object no. 1
                If WaterSprayDensity(fireroom) > 0 Then
                    sphrr = SprinklerHRR * Exp(-(tim - SprinklerTime) / (3 * (WaterSprayDensity(fireroom) / 60) ^ (-1.85)))
                End If
            End If
        ElseIf sprink_mode = 1 Then  'sprinkler control
            If SprinklerFlag = 1 Then
                'sprinkler has operated
                If id > 1 Then Exit Function 'only uses fire object no. 1
                sphrr = SprinklerHRR
            End If
        End If

        If growth = True Then
            'this is not a room corner test
start:
            ' broke bit

            '            'read in the hrr versus time data for the object
            '            If FuelResponseEffects = False Then
            '                For i = 1 To NumberDataPoints(id) - 1

            '                    If i + 1 > 500 Then 'max 500 data points allowed
            '                        hrr = HRRlast
            '                        Get_HRR = hrr
            '                        timlast = tim
            '                        HRRlast = hrr
            '                        IDLast = id
            '                        Q = hrr
            '                        Qlast = Q
            '                        QBlast = Qburner
            '                        QWlast = QWall
            '                        QClast = QCeiling
            '                        QFlast = QFloor
            '                        Exit Function
            '                    End If

            '                    If tim >= (HeatReleaseData(1, i, id) + ObjectIgnTime(id)) Then
            '                        If tim < (HeatReleaseData(1, i + 1, id) + ObjectIgnTime(id)) Then
            '                            A = tim - (HeatReleaseData(1, i, id) + ObjectIgnTime(id))
            '                            b = HeatReleaseData(1, i + 1, id) - HeatReleaseData(1, i, id) 'sec
            '                            C = HeatReleaseData(2, i + 1, id) - HeatReleaseData(2, i, id) 'kw
            '                            hrr = HeatReleaseData(2, i, id) + (A / b) * C
            '                            If hrr < 0 Then hrr = 0

            '                            'for a sprinkler fire, use the lesser of the theoretical hrr and the sprinkler controlled rate
            '                            If SprinklerFlag = 1 Then
            '                                If sprink_mode = 0 Then sphrr = hrr
            '                                If sprink_mode = 2 And WaterSprayDensity(fireroom) = 0 Then sphrr = hrr 'suppression mode
            '                                If sprink_mode = 1 And sphrr > hrr Then sphrr = hrr 'control mode, added 29/09/2011
            '                                hrr = sphrr
            '                            End If

            '                            If tim = 0 And hrr < 0.000001 Then hrr = 0.000001

            '                            If useCLTmodel = True And post = False Then 'CLT mod on and wood crib PF model off. 
            '                                QWall1 = wall_char(stepcount, 1) * WallEffectiveHeatofCombustion(fireroom) * 1000
            '                                QCeiling1 = ceil_char(stepcount, 1) * CeilingEffectiveHeatofCombustion(fireroom) * 1000
            '                                Get_HRR = hrr

            '                                timlast = tim
            '                                HRRlast = hrr
            '                                IDLast = id
            '                                Q = hrr 'plume?
            '                                Qlast = Q
            '                                QBlast = Qburner
            '                                QWlast = QWall1
            '                                QClast = QCeiling1
            '                                QFlast = QFloor1

            '                                Exit Function

            '                            End If

            '                            Get_HRR = hrr
            '                            timlast = tim
            '                            HRRlast = hrr
            '                            IDLast = id
            '                            Q = hrr
            '                            Qlast = Q
            '                            QBlast = Qburner
            '                            QWlast = QWall
            '                            QClast = QCeiling
            '                            QFlast = QFloor

            '                            Exit Function
            '                        End If
            '                    End If
            '                Next i
            '            Else
            '                'fuel response effects is true
            '                If useCLTmodel = True And post = False Then 'CLT mod on and wood crib PF model off. 
            '                    QWall1 = wall_char(stepcount - 1, 1) * WallEffectiveHeatofCombustion(fireroom) * 1000
            '                    QCeiling1 = ceil_char(stepcount - 1, 1) * CeilingEffectiveHeatofCombustion(fireroom) * 1000
            '                    Get_HRR = hrr
            '                    Qburner = hrr
            '                    timlast = tim
            '                    HRRlast = hrr
            '                    IDLast = id
            '                    Q = hrr
            '                    Qlast = Q
            '                    QBlast = Qburner
            '                    QWlast = QWall1
            '                    QClast = QCeiling1
            '                    QFlast = QFloor1
            '                    'If QWall1 > 1000 Then Stop
            '                    Exit Function

            '                End If
            '            End If


            'read in the hrr versus time data for the object
            For i = 1 To NumberDataPoints(id) - 1

                If i + 1 > 500 Then 'max 500 data points allowed
                    hrr = HRRlast
                    Get_HRR = hrr
                    timlast = tim
                    HRRlast = hrr
                    IDLast = id
                    Q = hrr
                    Qlast = Q
                    QBlast = Qburner
                    QWlast = QWall
                    QClast = QCeiling
                    QFlast = QFloor
                    Exit Function
                End If

                If tim >= (HeatReleaseData(1, i, id) + ObjectIgnTime(id)) Then
                    If tim < (HeatReleaseData(1, i + 1, id) + ObjectIgnTime(id)) Then
                        A = tim - (HeatReleaseData(1, i, id) + ObjectIgnTime(id))
                        b = HeatReleaseData(1, i + 1, id) - HeatReleaseData(1, i, id) 'sec
                        C = HeatReleaseData(2, i + 1, id) - HeatReleaseData(2, i, id) 'kw
                        hrr = HeatReleaseData(2, i, id) + (A / b) * C
                        If hrr < 0 Then hrr = 0

                        'for a sprinkler fire, use the lesser of the theoretical hrr and the sprinkler controlled rate
                        If SprinklerFlag = 1 Then
                            If sprink_mode = 0 Then sphrr = hrr
                            If sprink_mode = 2 And WaterSprayDensity(fireroom) = 0 Then sphrr = hrr 'suppression mode
                            If sprink_mode = 1 And sphrr > hrr Then sphrr = hrr 'control mode, added 29/09/2011
                            hrr = sphrr
                        End If

                        ''enhance option should be removed
                        'If Enhance = True And stepcount > 1 Then
                        '    'account for enhancement of burning due to room effects by adding to first object only
                        '    If FuelSurfaceArea = 0 Then
                        '        ' If FuelHoC(stepcount - 1) > 0 Then
                        '        If (EnergyYield(id) > 0) Then
                        '            'FSA = HeatRelease(fireroom, stepcount - 1, 1) / (FuelHoC(stepcount - 1) * MLUnitArea * 1000) 'sq.m
                        '            'FSA = hrr / (FuelHoC(stepcount - 1) * MLUnitArea * 1000) 'sq.m

                        '            If ObjectMLUA(2, id) > 0 Then 'HRRUA used
                        '                FSA = hrr / ObjectMLUA(2, id)  'm2
                        '            ElseIf ObjectMLUA(0, id) > 0 Or ObjectMLUA(1, id) > 0 Then
                        '                Dim mlua = ObjectMLUA(0, id) * (Target(fireroom, stepcount - 1) - Target(fireroom, 1)) + ObjectMLUA(1, id)
                        '                FSA = hrr / (EnergyYield(id) * mlua * 1000) 'sq.m
                        '            Else
                        '                'FSA = hrr / (EnergyYield(id) * MLUnitArea * 1000) 'sq.m
                        '                'm2 = kW / (kJ/g * kg/s/m2 * 1000)
                        '            End If
                        '        Else
                        '            FSA = 0
                        '        End If
                        '    Else
                        '        FSA = FuelSurfaceArea
                        '    End If
                        '    'deltaQ = FSA * EnergyYield(id) / FuelHeatofGasification * (Target(fireroom, stepcount - 1) - Target(fireroom, 1))
                        '    If ObjectLHoG(id) Then
                        '        deltaQ = FSA * EnergyYield(id) / ObjectLHoG(id) * (Target(fireroom, stepcount - 1) - Target(fireroom, 1))
                        '    End If
                        '    hrr = hrr + deltaQ
                        'End If

                        If tim = 0 And hrr < 0.000001 Then hrr = 0.000001

                        If useCLTmodel = True And post = False Then 'CLT mod on and wood crib PF model off. 
                            QWall = wall_char(stepcount, 1) * WallEffectiveHeatofCombustion(fireroom) * 1000
                            QCeiling = ceil_char(stepcount, 1) * CeilingEffectiveHeatofCombustion(fireroom) * 1000
                            Get_HRR = hrr + QWall + QCeiling

                            timlast = tim
                            HRRlast = hrr
                            IDLast = id
                            Q = hrr
                            Qlast = Q
                            QBlast = Qburner
                            QWlast = QWall
                            QClast = QCeiling
                            QFlast = QFloor

                            Exit Function

                        End If

                        Get_HRR = hrr
                        timlast = tim
                        HRRlast = hrr
                        IDLast = id
                        Q = hrr
                        Qlast = Q
                        QBlast = Qburner
                        QWlast = QWall
                        QClast = QCeiling
                        QFlast = QFloor

                        Exit Function
                    End If
                End If
            Next i

            'check to see if time corresponds to last data point
            If NumberDataPoints(id) <> 0 Then
                If tim >= (HeatReleaseData(1, NumberDataPoints(id), id) + ObjectIgnTime(id)) Then
                    hrr = HeatReleaseData(2, NumberDataPoints(id), id)
                    If SprinklerFlag = 1 Then
                        If sprink_mode = 0 Then sphrr = hrr
                        If sprink_mode = 2 And WaterSprayDensity(fireroom) = 0 Then sphrr = hrr 'suppression mode
                        If sprink_mode = 1 And sphrr > hrr Then sphrr = hrr 'control mode, added 29/09/2011

                        hrr = sphrr
                    End If
                End If
            End If

        Else
            'this is a room corner test
            If id = Max(1, burner_id) Then
                If quintiere = True Then
                    'quintiere's room corner model
                    hrr = HRR_Quintiere2(id, tim, Q, Qburner, QFloor, QWall, QCeiling) 'hrr includes burner + linings
                Else
                    'Karlsson's room corner Model A
                    hrr = HRR_ModelA(id, tim, Q)
                End If
            Else
                'this item has not ignited the wall
                GoTo start
            End If

            'for a sprinkler fire, use the lesser of the theoretical hrr and the sprinkler controlled rate
            If SprinklerFlag = 1 Then
                If sprink_mode = 0 Then sphrr = hrr
                If sprink_mode = 2 And WaterSprayDensity(fireroom) = 0 Then sphrr = hrr 'suppression mode
                If sprink_mode = 1 And sphrr > hrr Then sphrr = hrr 'control mode, added 29/09/2011
                hrr = sphrr
            End If

        End If
        If flagstop = 1 Then Exit Function


        Get_HRR = hrr

        If id = Max(1, burner_id) Then '21/11/2011
            timlast = tim
            HRRlast = hrr
            IDLast = id
            Qlast = Q
            QBlast = Qburner
            QWlast = QWall
            QClast = QCeiling
            QFlast = QFloor
        End If

    End Function
    '    Function Get_HRR_old(ByVal id As Integer, ByVal tim As Double, ByRef Q As Double, ByRef Qburner As Double, ByRef QFloor As Double, ByRef QWall As Double, ByRef QCeiling As Double) As Double
    '        '*  ====================================================================
    'jan 18 broke version

    '        '*  Returns the heat release rate (kW) at a given time (sec) for a single
    '        '*  object. Linearly interpolates between data points.
    '        '*
    '        '*  ID = the burning object number
    '        '*  tim = Time (seconds)
    '        '*  Q = hrr to use in plume correlation (kW/m2)
    '        '*
    '        '*  ====================================================================

    '        If FuelResponseEffects = True Then
    '            'won't work for multiple objects
    '            'only need free burn MLT here 
    '            'well ventilated
    '            Q = 1000 * EnergyYield(id) * FuelBurningRate(0, fireroom, stepcount) 'kW freeburn
    '            'Q = 1000 * EnergyYield(id) * FuelBurningRate(4, fireroom, stepcount) 'kW burning rate
    '            Get_HRR = Q
    '            Exit Function

    '            'Dim o2lower As Double = O2MassFraction(fireroom, stepcount, 2)
    '            'Dim idg As Integer = 1
    '            'Dim MLR As Double
    '            'Dim S As Double = 15.1 'stoichiometric air to fuel ratio for heptane

    '            'MLR = MassLoss_ObjectwithFuelResponse(idg, tim, Qburner, QFloor, QWall, QCeiling, o2lower, massplumeflow(stepcount, fireroom)) 'kg/s
    '            'If GlobalER(stepcount) > 1 Then
    '            '    'under ventilated
    '            '    Q = 1000 * EnergyYield(1) * massplumeflow(stepcount, fireroom) * o2lower / 0.231 / S 'kW
    '            'Else
    '            '    'well ventilated
    '            '    Q = 1000 * EnergyYield(1) * MLR 'kW
    '            'End If
    '            'Get_HRR = Q
    '            'Exit Function
    '        End If

    '        Dim i As Short
    '        Dim deltaQ, FSA As Double

    '        Dim hrr, b, A, C, sphrr As Double
    '        Static QClast, QFlast, HRRlast, timlast, Qlast, QWlast, QBlast As Double
    '        Static IDLast As Short
    '        Static growth, post, quintiere As Boolean

    '        If tim = 0 Then
    '            HRRlast = 0
    '            Qlast = 0
    '            QBlast = 0
    '            QWlast = 0
    '            QClast = 0
    '            QFlast = 0
    '            IDLast = id '2011.26
    '        End If

    '        If tim = timlast And id = IDLast And tim > 0 Then
    '            Get_HRR = HRRlast
    '            Q = Qlast
    '            Qburner = QBlast
    '            QWall = QWlast
    '            QCeiling = QClast
    '            QFloor = QFlast
    '            Exit Function
    '        End If

    '        If tim = 0 Then
    '            post = frmOptions1.optPostFlashover.Checked 'remember this value
    '            growth = frmOptions1.optRCNone.Checked
    '            quintiere = frmOptions1.optQuintiere.Checked
    '        End If

    '        If Flashover = True And post = True Then

    '            'only need to call this once per timestep when using the wood crib PF model
    '            If IEEERemainder(tim, Timestep) = 0 Then
    '                Dim dummy As Double 'returns contribution from wall, ceiling

    '                If frmCLT.chkWoodIntegralModel.Checked = True Then
    '                    hrr = MassLoss_Total_pluswood(tim, dummy, dummy) * NewHoC_fuel * 1000 'spearpoint quintiere
    '                Else
    '                    hrr = MassLoss_Total(tim) * NewHoC_fuel * 1000
    '                End If

    '                Get_HRR = hrr
    '                timlast = tim
    '                HRRlast = hrr
    '                IDLast = id
    '                Q = hrr
    '                Qlast = Q
    '                QBlast = Qburner
    '                QWlast = QWall
    '                QClast = QCeiling
    '                QFlast = QFloor
    '            Else
    '                Get_HRR = HRRlast
    '                Q = Qlast
    '                Qburner = QBlast
    '                QWall = QWlast
    '                QCeiling = QClast
    '                QFloor = QFlast
    '            End If

    '            Exit Function

    '        End If

    '        'not using wood crib postflashover model
    '        If usepowerlawdesignfire = True Then
    '            If id = 1 Then '1 item only
    '                hrr = Get_HRR_powerlaw(tim, Q)
    '            Else
    '                hrr = 0
    '                Q = 0
    '            End If
    '            Get_HRR = hrr
    '            timlast = tim
    '            HRRlast = hrr
    '            IDLast = id
    '            Q = hrr
    '            Qlast = Q
    '            QBlast = Qburner
    '            QWlast = QWall
    '            QClast = QCeiling
    '            QFlast = QFloor
    '            Exit Function
    '        End If

    '        'If stepcount = 100 Then Stop
    '        'check to see if any fuel left
    '        'If frmOptions1.optFuelLimited.Checked = True And post = False And stepcount > 1 Then
    '        If post = False And stepcount > 1 Then 'if not using wood crib post flashover model
    '            If growth = True Then
    '                'not a room corner test

    '                If TotalFuel(stepcount - 1) > InitialFuelMass Then
    '                    If useCLTmodel = True And TotalFuel(stepcount - 1) < fuelmasswithCLT Then
    '                        'continue
    '                    Else
    '                        Get_HRR = 0
    '                        Q = Qlast
    '                        Qburner = QBlast
    '                        QWall = QWlast
    '                        QCeiling = QClast
    '                        QFloor = QFlast
    '                        IDLast = id '2011.26
    '                        Exit Function
    '                    End If

    '                End If
    '            End If
    '        End If

    '        If sprink_mode = 2 Then 'suppression
    '            If SprinklerFlag = 1 Then
    '                'sprinkler has operated
    '                If id > 1 Then Exit Function 'only uses fire object no. 1
    '                If WaterSprayDensity(fireroom) > 0 Then
    '                    sphrr = SprinklerHRR * Exp(-(tim - SprinklerTime) / (3 * (WaterSprayDensity(fireroom) / 60) ^ (-1.85)))
    '                End If
    '            End If
    '        ElseIf sprink_mode = 1 Then  'control
    '            If SprinklerFlag = 1 Then
    '                'sprinkler has operated
    '                If id > 1 Then Exit Function
    '                sphrr = SprinklerHRR
    '            End If
    '        End If

    '        If growth = True Then
    '            'this is not a room corner test
    'start:
    '            '14/10/2002
    '            For i = 1 To NumberDataPoints(id) - 1


    '                If i + 1 > 500 Then
    '                    hrr = HRRlast
    '                    Get_HRR = hrr
    '                    timlast = tim
    '                    HRRlast = hrr
    '                    IDLast = id
    '                    Q = hrr
    '                    Qlast = Q
    '                    QBlast = Qburner
    '                    QWlast = QWall
    '                    QClast = QCeiling
    '                    QFlast = QFloor
    '                    Exit Function
    '                End If

    '                If tim >= (HeatReleaseData(1, i, id) + ObjectIgnTime(id)) Then
    '                    If tim < (HeatReleaseData(1, i + 1, id) + ObjectIgnTime(id)) Then
    '                        A = tim - (HeatReleaseData(1, i, id) + ObjectIgnTime(id))
    '                        b = HeatReleaseData(1, i + 1, id) - HeatReleaseData(1, i, id) 'sec
    '                        C = HeatReleaseData(2, i + 1, id) - HeatReleaseData(2, i, id) 'kw
    '                        hrr = HeatReleaseData(2, i, id) + (A / b) * C
    '                        If hrr < 0 Then hrr = 0

    '                        'for a sprinkler fire, use the lesser of the theoretical hrr and the sprinkler controlled rate
    '                        If SprinklerFlag = 1 Then
    '                            If sprink_mode = 0 Then sphrr = hrr
    '                            If sprink_mode = 2 And WaterSprayDensity(fireroom) = 0 Then sphrr = hrr 'suppression mode
    '                            If sprink_mode = 1 And sphrr > hrr Then sphrr = hrr 'control mode, added 29/09/2011
    '                            hrr = sphrr
    '                        End If

    '                        ' If Enhance = True And id = 1 And stepcount > 1 Then
    '                        If Enhance = True And stepcount > 1 Then
    '                            'account for enhancement of burning due to room effects by adding to first object only
    '                            If FuelSurfaceArea = 0 Then
    '                                ' If FuelHoC(stepcount - 1) > 0 Then
    '                                If (EnergyYield(id) > 0) Then
    '                                    'FSA = HeatRelease(fireroom, stepcount - 1, 1) / (FuelHoC(stepcount - 1) * MLUnitArea * 1000) 'sq.m
    '                                    'FSA = hrr / (FuelHoC(stepcount - 1) * MLUnitArea * 1000) 'sq.m

    '                                    If ObjectMLUA(2, id) > 0 Then 'HRRUA used
    '                                        FSA = hrr / ObjectMLUA(2, id)  'm2
    '                                    ElseIf ObjectMLUA(0, id) > 0 Or ObjectMLUA(1, id) > 0 Then
    '                                        Dim mlua = ObjectMLUA(0, id) * (Target(fireroom, stepcount - 1) - Target(fireroom, 1)) + ObjectMLUA(1, id)
    '                                        FSA = hrr / (EnergyYield(id) * mlua * 1000) 'sq.m
    '                                    Else
    '                                        'FSA = hrr / (EnergyYield(id) * MLUnitArea * 1000) 'sq.m
    '                                        'm2 = kW / (kJ/g * kg/s/m2 * 1000)
    '                                    End If
    '                                Else
    '                                    FSA = 0
    '                                End If
    '                            Else
    '                                FSA = FuelSurfaceArea
    '                            End If
    '                            'deltaQ = FSA * EnergyYield(id) / FuelHeatofGasification * (Target(fireroom, stepcount - 1) - Target(fireroom, 1))
    '                            If ObjectLHoG(id) Then
    '                                deltaQ = FSA * EnergyYield(id) / ObjectLHoG(id) * (Target(fireroom, stepcount - 1) - Target(fireroom, 1))
    '                            End If
    '                            hrr = hrr + deltaQ
    '                        End If


    '                        If tim = 0 And hrr < 0.000001 Then hrr = 0.000001

    '                        If useCLTmodel = True And post = False Then 'CLT mod on and wood crib PF model off. 
    '                            QWall = wall_char(stepcount, 1) * WallEffectiveHeatofCombustion(fireroom) * 1000
    '                            QCeiling = ceil_char(stepcount, 1) * CeilingEffectiveHeatofCombustion(fireroom) * 1000

    '                        End If

    '                        Get_HRR = hrr
    '                        timlast = tim
    '                        HRRlast = hrr
    '                        IDLast = id
    '                        Q = hrr
    '                        Qlast = Q
    '                        QBlast = Qburner
    '                        QWlast = QWall
    '                        QClast = QCeiling
    '                        QFlast = QFloor

    '                        Exit Function
    '                    End If
    '                End If
    '            Next i

    '            'check to see if time corresponds to last data point
    '            If NumberDataPoints(id) <> 0 Then
    '                If tim >= (HeatReleaseData(1, NumberDataPoints(id), id) + ObjectIgnTime(id)) Then
    '                    hrr = HeatReleaseData(2, NumberDataPoints(id), id)
    '                    If SprinklerFlag = 1 Then
    '                        If sprink_mode = 0 Then sphrr = hrr
    '                        If sprink_mode = 2 And WaterSprayDensity(fireroom) = 0 Then sphrr = hrr 'suppression mode
    '                        If sprink_mode = 1 And sphrr > hrr Then sphrr = hrr 'control mode, added 29/09/2011

    '                        hrr = sphrr
    '                    End If
    '                End If
    '            End If




    '        Else
    '            'this is a room corner test
    '            If id = Max(1, burner_id) Then
    '                If quintiere = True Then
    '                    'quintiere's room corner model
    '                    hrr = HRR_Quintiere2(id, tim, Q, Qburner, QFloor, QWall, QCeiling) 'hrr includes burner + linings
    '                Else
    '                    'Karlsson's room corner Model A
    '                    hrr = HRR_ModelA(id, tim, Q)
    '                End If
    '            Else
    '                'this item has not ignited the wall
    '                GoTo start
    '            End If

    '            'for a sprinkler fire, use the lesser of the theoretical hrr and the sprinkler controlled rate
    '            If SprinklerFlag = 1 Then
    '                If sprink_mode = 0 Then sphrr = hrr
    '                If sprink_mode = 2 And WaterSprayDensity(fireroom) = 0 Then sphrr = hrr 'suppression mode
    '                If sprink_mode = 1 And sphrr > hrr Then sphrr = hrr 'control mode, added 29/09/2011
    '                hrr = sphrr
    '            End If

    '        End If
    '        If flagstop = 1 Then Exit Function


    '        Get_HRR = hrr

    '        If id = Max(1, burner_id) Then '21/11/2011
    '            timlast = tim
    '            HRRlast = hrr
    '            IDLast = id
    '            Qlast = Q
    '            QBlast = Qburner
    '            QWlast = QWall
    '            QClast = QCeiling
    '            QFlast = QFloor
    '        End If

    '    End Function


    Sub Get_Plume()
        '*  ===================================================================
        '*  This subprocedure determines which plume correlation model
        '*  the user has selected in the program.
        '*  ===================================================================
        Static gen As Boolean
        'If stepcount = 1 Then gen = frmOptions1.optStrongPlume.Checked
        gen = frmOptions1.optStrongPlume.Checked

        If gen = True Then
            'use general strong plume and Delichatsios correlation
            PlumeModel = "General"
        Else
            'use McCaffrey's correlation
            PlumeModel = "McCaffrey"
        End If

    End Sub

    Function Get_Visibility(ByVal MW_layer As Double, ByVal pressure As Double, ByVal T As Double, ByVal Y As Double, ByVal MT As Double, ByVal MSoot As Double, ByRef OpticalDensity As Single) As Single
        '*  =====================================================
        '*      This function returns the value of visibility in
        '*      the upper layer. Maximum visibility is assumed
        '*      to be 20 metres.
        '*
        '*      Arguments passed to the function:
        '*          UT      = layer temperature (K)
        '*          Y       = mass fraction of soot in layer
        '*                      (kg soot / kg mixture)
        '*          MT      = total mass loss rate of fuel (kg fuel/s)
        '*          MSoot   = total soot generation rate (kg soot/s)
        '*
        '*      Global variables used:
        '*          ReferenceDensity
        '*          ReferenceTemp
        '*  =====================================================

        Dim Concentration As Single
        Dim AvgExtinction As Single

        'Concentration = ReferenceDensity * ReferenceTemp * Y / UT
        'Concentration = Density * Y
        'Concentration = (pressure / 1000 + Atm_Pressure) / (Gas_Constant_Air * UT) * Y
        If T > 0 Then
            Concentration = (pressure / 1000 + Atm_Pressure) / (Gas_Constant / MW_layer * T) * Y 'kg-soot per kg-gas layer
        End If

        'If MT > 0 Then
        '    WeightedSootYield = MSoot / MT
        'Else
        '    WeightedSootYield = 0
        'End If
        'ParticleExtinction = 10750 * Exp(-4.95 * WeightedSootYield)

        'ParticleExtinction = 7600 'm2/kg Seider and Einhorn, flaming combustion
        'ParticleExtinction = 8790 'm2/kg Mulholland and Choi
        AvgExtinction = Concentration * ParticleExtinction '1/m
        OpticalDensity = AvgExtinction / 2.3

        If AvgExtinction > 0 Then
            If frmOptions1.optReflectiveSign.Checked = True Then
                If 3 / AvgExtinction < 20 Then
                    Get_Visibility = 3 / AvgExtinction
                    Exit Function
                Else
                    Get_Visibility = 20
                    Exit Function
                End If
            ElseIf frmOptions1.optIlluminatedSign.Checked = True Then
                If 8 / AvgExtinction < 20 Then
                    Get_Visibility = 8 / AvgExtinction
                    Exit Function
                Else
                    Get_Visibility = 20
                    Exit Function
                End If
            End If
        Else
            Get_Visibility = 20
            Exit Function
        End If

    End Function

    Sub Graph_Data(ByRef index As Short, ByRef YAxisTitle As String, ByRef datatobeplotted(,) As Double, ByRef DataShift As Double, ByRef multiplier As Double, ByRef Style As Short, ByRef ymax As Double)
        '		Dim frmGraphs As Object
        '		'*  ====================================================================
        '		'*  This function takes data for a variable from a one-dimensional array
        '		'*  and displays it in a graph
        '		'*
        '		'*  Revised 6 September 1996 Colleen Wade
        '		'*  ====================================================================


        'Dim i As Short
        Dim ydata(0 To NumberTimeSteps) As Double
        Dim room As Integer


        'if no data exists
        If NumberTimeSteps < 1 Then
            MsgBox("There is no data to plot, please run the simulation first.", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Dim chdata(0 To NumberTimeSteps, 0 To 2 * NumberRooms - 1) As Object
        Dim j As Integer

        frmPlot.Chart1.Series.Clear()
        room = 1
        For room = 1 To NumberRooms

            frmPlot.Chart1.Series.Add("Room " & room)

            frmPlot.Chart1.Series("Room " & room).ChartType = SeriesChartType.FastLine

            For j = 1 To NumberTimeSteps
                ydata(j) = datatobeplotted(j, index) * multiplier + DataShift 'data to be plotted
                frmPlot.Chart1.Series("Room " & room).Points.AddXY(tim(j, 1), ydata(j))
            Next

        Next room



        'room = 1
        'For i = 0 To 2 * NumberRooms - 1 Step 2

        '    If room = fireroom Then chdata(0, i) = "Room " & CStr(room)

        '    For j = 1 To NumberTimeSteps
        '        chdata(j, i) = tim(j, 1) 'time
        '        chdata(j, i + 1) = datatobeplotted(j, index) * multiplier + DataShift 'data to be plotted
        '    Next

        '    room = room + 1
        'Next i

        '.ChartData = chdata

        'End With
        frmPlot.Chart1.BackColor = Color.AliceBlue
        frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
        frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
        frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = YAxisTitle
        'frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
        frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
        frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
        frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
        frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
        frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
        frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
        frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
        'frmPlot.Chart1.Titles("Title1").Text = Title

        frmPlot.Chart1.Visible = True
        frmPlot.BringToFront()
        frmPlot.Show()


        'frmGraph.Show()


    End Sub

    Sub Graph_Data2(ByRef graph As System.Windows.Forms.Control, ByRef YAxisTitle As String, ByRef datatobeplotted() As Double, ByRef DataShift As Single, ByRef multiplier As Single, ByRef Style As Short, ByRef ymax As Single)
        '		'*  ====================================================================
        '		'*  This function takes data for a variable from a one-dimensional array
        '		'*  and displays it in a run-time graph
        '		'*
        '		'*  Revised 9 September 1996 Colleen Wade
        '		'*  ====================================================================

        '		Static i As Short
        '		Static factor As Short

        '		On Error GoTo graphhandler1

        '		'if no data exists
        '		If NumberTimeSteps < 1 Then Exit Sub
        '		If stepcount < 2 Then Exit Sub

        '		'frmGraphs.Show

        '		With graph
        '			'.NumPoints = stepcount
        '			'.NumSets = 1
        '			If Style = 2 Then
        '				'UPGRADE_WARNING: Couldn't resolve default property of object graph.YAxisMax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '				.YAxisMax = ymax
        '			End If
        '		End With

        'here1: 

        '		factor = 1
        '		Do While stepcount / factor > 1200
        '			factor = factor * 2
        '		Loop 

        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.NumPoints = Int(stepcount / factor)

        '		With graph
        '			If i = 1 Then
        '				'UPGRADE_WARNING: Couldn't resolve default property of object graph.Graphdata. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '				.Graphdata = (multiplier * datatobeplotted(1) + DataShift)
        '			End If
        '			'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '			For i = 2 To .NumPoints
        '				'UPGRADE_WARNING: Couldn't resolve default property of object graph.Graphdata. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '				.Graphdata = (multiplier * datatobeplotted(factor * i) + DataShift)
        '			Next i

        '			'UPGRADE_WARNING: Couldn't resolve default property of object graph.XPosData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '			.XPosData = (tim(1, 1))
        '			'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '			For i = 2 To .NumPoints
        '				'UPGRADE_WARNING: Couldn't resolve default property of object graph.XPosData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '				.XPosData = (tim(factor * i, 1))
        '			Next i
        '			'UPGRADE_ISSUE: Control method graph.DrawMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        '			.DrawMode = 3

        '		End With
        '		Exit Sub

        'graphhandler1: 
        '		'graph.NumPoints = graph.NumPoints - 1
        '		GoTo here1

    End Sub


    Sub Graph_Data_fuel(ByRef room As Short, ByRef Title As String, ByRef datatobeplotted() As Double, ByRef DataShift As Double, ByRef DataMultiplier As Double, ByRef GraphSty As Short, ByRef MaxYValue As Double, ByVal timeunit As Integer)
        '*  ====================================================================
        '*  This function takes data for a variable from a two-dimensional array
        '*  and displays it in a graph
        '*  ====================================================================
        Dim j As Integer
        Dim ydata(0 To NumberTimeSteps) As Double

        'if no data exists
        If NumberTimeSteps < 1 Then
            MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Sub
        End If

        Try

            frmPlot.Chart1.Series.Clear()
            room = 1
            For room = 1 To NumberRooms

                frmPlot.Chart1.Series.Add("Room " & room)

                frmPlot.Chart1.Series("Room " & room).ChartType = SeriesChartType.FastLine

                If Not roomcolor(room - 1).IsEmpty Then
                    frmPlot.Chart1.Series("Room " & room).Color = roomcolor(room - 1) '=line color
                End If

                For j = 1 To NumberTimeSteps
                    ydata(j) = datatobeplotted(j) * DataMultiplier + DataShift 'data to be plotted
                    frmPlot.Chart1.Series("Room " & room).Points.AddXY(tim(j, 1) / timeunit, ydata(j))
                Next

            Next room
            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            'frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            If timeunit = 60 Then
                frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (min)"
            Else
                frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            End If
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            'frmPlot.Chart1.Titles("Title1").Text = Title

            'frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Graph_Data_fuel")
        End Try

    End Sub
    Sub Graph_Data_furnaceNHL(ByRef room As Short, ByRef Title As String, ByRef datatobeplotted(,) As Double, ByRef DataShift As Double, ByRef DataMultiplier As Double, ByRef GraphSty As Short, ByRef MaxYValue As Double, ByRef MaxXValue As Double)
        '*  ====================================================================
        '*  This function takes data for a variable from a one-dimensional array
        '*  and displays it in a graph
        '*  ====================================================================
        Dim j, material As Integer
        Dim ydata(0 To MaxXValue / Timestep + 1) As Double
        Dim descr1 As String = "material"

        'if no data exists
        'If NumberTimeSteps < 1 Then
        '    MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
        '    Exit Sub
        'End If

        Try

            frmPlot.Chart1.Series.Clear()
            room = fireroom

            For material = 0 To 2
                If material = 0 Then descr1 = "ceiling material"
                If material = 1 Then descr1 = "wall material"
                If material = 2 Then descr1 = "floor material"

                frmPlot.Chart1.Series.Add(descr1)

                frmPlot.Chart1.Series(descr1).ChartType = SeriesChartType.FastLine

                If Not roomcolor(material).IsEmpty Then
                    frmPlot.Chart1.Series(descr1).Color = roomcolor(material) '=line color
                End If

                For j = 0 To MaxXValue / Timestep
                    ydata(j) = datatobeplotted(material, j) * DataMultiplier + DataShift 'data to be plotted
                    frmPlot.Chart1.Series(descr1).Points.AddXY(j * Timestep, ydata(j))

                Next

            Next material


            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            'frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = MaxXValue
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Minimum = 0
            'frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Maximum = MaxYValue

            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Interval = 3600
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.MajorGrid.Interval = 3600
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            'frmPlot.Chart1.Titles("Title1").Text = Title

            'frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Graph_Data_furnaceNHL")
        End Try

    End Sub
    Sub Graph_Data_furnaceST(ByRef material As Integer, ByRef room As Short, ByRef Title As String, ByRef datatobeplotted(,) As Double, ByRef DataShift As Double, ByRef DataMultiplier As Double, ByRef GraphSty As Short, ByRef MaxYValue As Double, ByRef MaxXValue As Double)
        '*  ====================================================================
        '*  This function takes data for a variable from a one-dimensional array
        '*  and displays it in a graph
        '*  ====================================================================
        Dim j As Integer
        Dim ydata(0 To MaxXValue / Timestep + 1) As Double
        Dim descr1 As String = "material"

        Try

            frmPlot.Chart1.Series.Clear()
            room = fireroom

            For material = 0 To 2
                If material = 0 Then descr1 = "ceiling material"
                If material = 1 Then descr1 = "wall material"
                If material = 2 Then descr1 = "floor material"

                frmPlot.Chart1.Series.Add(descr1)

                frmPlot.Chart1.Series(descr1).ChartType = SeriesChartType.FastLine

                If Not roomcolor(material).IsEmpty Then
                    frmPlot.Chart1.Series(descr1).Color = roomcolor(material) '=line color
                End If


                For j = 0 To MaxXValue / Timestep
                    ydata(j) = datatobeplotted(material, j) * DataMultiplier + DataShift 'data to be plotted
                    frmPlot.Chart1.Series(descr1).Points.AddXY(j * Timestep, ydata(j))

                Next

            Next material


            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            'frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = MaxXValue
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Minimum = 0

            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Interval = 600
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.MajorGrid.Interval = 600
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            'frmPlot.Chart1.Titles("Title1").Text = Title

            'frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Graph_Data_furnaceST")
        End Try

    End Sub

    Sub Ignition_Correlation2(ByRef radflux() As Double, ByRef Ignition() As Double, ByRef IgnitionTime() As Double, ByRef IgnPoints As Short, ByRef matemm As Double, ByRef ThermalInertia As Double, ByRef IgnitionTemp As Double, ByRef Qcrit As Double)
        '******************************************************
        '*  Find the ignition temperature and thermal inertia
        '*  from a correlation of cone ignition data using the
        '*  method of Marc Janssens. Assumes the material behaves
        '*  as a semi-infinite solid, thermally thick.
        '*
        '*  2012
        '******************************************************

        Dim slope, Yintercept, R2 As Double
        Dim upperbound, hc, var, lowerbound As Double
        Dim TMax, hig, Amb, RMax As Double
        Dim n As Single
        Dim i As Short
        Dim Newslope As Double

        On Error GoTo IgnitionHandler

        'convective heat transfer coefficient in cone
        hc = 0.0135 'kW/m2K

        'assumed ambient temp for cone tests
        Amb = 293 'K

        upperbound = 2000
        lowerbound = Amb

        'fit a linear regression line to the data, plotting vs external flux
        Call RegressionL(radflux, Ignition, IgnPoints, Yintercept, slope, R2)

        'find X-intercept = Q critical
        'Xintercept = -Yintercept / slope

        'the X-intercept should not be too small - fudge
        'If Xintercept < 1 Then
        '    Xintercept = 1 'kw/m2
        '    MsgBox("The ignition data provided is showing poor correlation and any results obtained using this material are unlikely to be valid. Additional data may be required.")
        'End If

        'guess ignition temp
        IgnitionTemp = 2 * Amb

        'find ignition temperature by iteration
        var = 1000 * hc * (IgnitionTemp - Amb) + matemm * StefanBoltzmann * (IgnitionTemp ^ 4 - Amb ^ 4)

        Do Until Abs(var - 1000 * Qcrit * matemm) < 0.0001

            If var > 1000 * Qcrit * matemm Then
                'reduce guess
                upperbound = IgnitionTemp
                IgnitionTemp = (IgnitionTemp + lowerbound) / 2
            Else
                'increase guess
                lowerbound = IgnitionTemp
                IgnitionTemp = (IgnitionTemp + upperbound) / 2
            End If

            'find new ignition temperature by iteration
            var = 1000 * hc * (IgnitionTemp - Amb) + matemm * StefanBoltzmann * (IgnitionTemp ^ 4 - Amb ^ 4)
        Loop

        'find total heat transfer coefficient
        hig = (Qcrit * matemm) / (IgnitionTemp - Amb)

        TMax = 0
        For i = 1 To IgnPoints
            If IgnitionTime(i) > 0 Then
                If (1 / IgnitionTime(i)) ^ 0.547 > TMax Then
                    TMax = (1 / IgnitionTime(i)) ^ 0.547
                    RMax = radflux(i)
                End If
            End If
        Next

        Newslope = TMax / (RMax - Qcrit)

        'find thermal inertia
        ThermalInertia = hig ^ 2 * (1 / (0.73 * Newslope * Qcrit)) ^ (1.828)

        Exit Sub

IgnitionHandler:
        flagstop = 1
        Exit Sub

    End Sub
    Sub Ignition_Correlation(ByRef radflux() As Double, ByRef Ignition() As Double, ByRef IgnitionTime() As Double, ByRef IgnPoints As Short, ByRef matemm As Double, ByRef ThermalInertia As Double, ByRef IgnitionTemp As Double, ByRef Xintercept As Double)
        '******************************************************
        '*  Find the ignition temperature and thermal inertia
        '*  from a correlation of cone ignition data using the
        '*  method of Marc Janssens. Assumes the material behaves
        '*  as a semi-infinite solid, thermally thick.
        '*
        '*  4 April 1998 Colleen Wade
        '******************************************************

        Dim slope, Yintercept, R2 As Double
        Dim upperbound, hc, var, lowerbound As Double
        Dim TMax, hig, Amb, RMax As Double
        Dim n As Single
        Dim i As Short
        Dim Newslope As Double

        On Error GoTo IgnitionHandler

        'convective heat transfer coefficient in cone
        hc = 0.0135 'kW/m2K

        'assumed ambient temp for cone tests
        Amb = 293 'K

        upperbound = 2000
        lowerbound = Amb

        'fit a linear regression line to the data, plotting vs external flux
        Call RegressionL(radflux, Ignition, IgnPoints, Yintercept, slope, R2)

        'find X-intercept = Q critical
        Xintercept = -Yintercept / slope

        'the X-intercept should not be too small - fudge
        If Xintercept < 1 Then
            Xintercept = 1 'kw/m2
            MsgBox("The ignition data provided is showing poor correlation and any results obtained using this material are unlikely to be valid. Additional data may be required.")
        End If

        'guess ignition temp
        IgnitionTemp = 2 * Amb

        'find ignition temperature by iteration
        var = 1000 * hc * (IgnitionTemp - Amb) + matemm * StefanBoltzmann * (IgnitionTemp ^ 4 - Amb ^ 4)

        Do Until Abs(var - 1000 * Xintercept * matemm) < 0.0001

            If var > 1000 * Xintercept * matemm Then
                'reduce guess
                upperbound = IgnitionTemp
                IgnitionTemp = (IgnitionTemp + lowerbound) / 2
            Else
                'increase guess
                lowerbound = IgnitionTemp
                IgnitionTemp = (IgnitionTemp + upperbound) / 2
            End If

            'find new ignition temperature by iteration
            var = 1000 * hc * (IgnitionTemp - Amb) + matemm * StefanBoltzmann * (IgnitionTemp ^ 4 - Amb ^ 4)
        Loop

        'find total heat transfer coefficient
        hig = (Xintercept * matemm) / (IgnitionTemp - Amb)

        TMax = 0
        For i = 1 To IgnPoints
            If IgnitionTime(i) > 0 Then
                If (1 / IgnitionTime(i)) ^ 0.547 > TMax Then
                    TMax = (1 / IgnitionTime(i)) ^ 0.547
                    RMax = radflux(i)
                End If
            End If
        Next

        Newslope = TMax / (RMax - Xintercept)

        'find thermal inertia
        'ThermalInertia = hig ^ 2 * (1 / (.73 * slope * XIntercept)) ^ (1.828)
        ThermalInertia = hig ^ 2 * (1 / (0.73 * Newslope * Xintercept)) ^ (1.828)

        Exit Sub

IgnitionHandler:
        flagstop = 1
        Exit Sub

    End Sub

    Sub Implicit_thermal_props()
        '*************************************************************
        '*  Determines heat transfer parameters required for an
        '*  implicit finite difference scheme
        '*
        '*  Revised 3 June 1997 Colleen Wade
        '*************************************************************

        Dim room As Short

        On Error GoTo Prophandler

        For room = 1 To NumberRooms

            'Thermal Diffusivities
            AlphaWall(room) = WallConductivity(room) / (WallSpecificHeat(room) * WallDensity(room))
            AlphaCeiling(room) = CeilingConductivity(room) / (CeilingSpecificHeat(room) * CeilingDensity(room))
            AlphaFloor(room) = FloorConductivity(room) / (FloorSpecificHeat(room) * FloorDensity(room))

            'Find DeltaX
            WallDeltaX(room) = (WallThickness(room) / 1000) / (Wallnodes - 1)
            CeilingDeltaX(room) = (CeilingThickness(room) / 1000) / (Ceilingnodes - 1)
            FloorDeltaX(room) = (FloorThickness(room) / 1000) / (Floornodes - 1)

            'Find Biot Numbers
            WallOutsideBiot(room) = OutsideConvCoeff * WallDeltaX(room) / WallConductivity(room)
            CeilingOutsideBiot(room) = OutsideConvCoeff * CeilingDeltaX(room) / CeilingConductivity(room)
            FloorOutsideBiot(room) = OutsideConvCoeff * FloorDeltaX(room) / FloorConductivity(room)

            'Find Fourier Numbers
            WallFourier(room) = AlphaWall(room) * Timestep / (WallDeltaX(room)) ^ 2
            CeilingFourier(room) = AlphaCeiling(room) * Timestep / (CeilingDeltaX(room) ^ 2)
            FloorFourier(room) = AlphaFloor(room) * Timestep / (FloorDeltaX(room)) ^ 2
        Next room
        Exit Sub

Prophandler:
        Exit Sub

    End Sub

    Sub Initialize_EndPointFlags()
        '*  =====================================================
        '*      This procedure initializes strings relating to
        '*      endpoint conditions.
        '*
        '*  Revised 19 October 1996 Colleen Wade
        '*  =====================================================

        frmEndPoints.lblTarget.Text = ""
        frmEndPoints.lblTemp.Text = ""
        frmEndPoints.lblVisibility.Text = ""
        frmEndPoints.lblConvect.Text = ""
        frmEndPoints.lblSprinkler.Text = ""
        frmEndPoints.lblFED.Text = ""
        OtherMessages = ""

    End Sub

    Sub Interpolate_S(ByRef X() As Double, ByRef Y() As Double, ByVal n As Integer, ByVal Xint As Double, ByRef yinterp As Double)
        '****************************************************************************
        '*  From ProMath 2.0
        '*  This Subroutine Performs Linear Interpolation Within A Set Of X(),Y()
        '*  Pairs To Give The Y Value Corresponding To Xint.
        '*       X       Array Of Values Of The Independent Variable
        '*       Y       Array Of Values Of The Dependent Variable Corresponding To X
        '*       N       Number Of Points To Interpolate.
        '*       Xint    The X-value For Which Estimate Of Y Is Desired.
        '*       Yinterp    The Y-value Returned To The User.
        '*
        '*  Revised Colleen Wade 12 October 1996
        '****************************************************************************

        Dim i As Integer
        Dim y1, x1, x2, y2 As Double

        yinterp = 0

        For i = 1 To n - 1
            If X(i) <= Xint And X(i + 1) >= Xint Then
                x1 = X(i)
                x2 = X(i + 1)
                y1 = Y(i)
                y2 = Y(i + 1)
                yinterp = (Xint - x1) * (y2 - y1) / (x2 - x1) + y1
                Exit Sub
            End If
        Next i

    End Sub

    'Function Mass_Plume(ByVal tim As Double, ByVal Z1 As Double, ByVal qmax As Double, ByVal UT As Double, ByVal LT As Double) As Double
    '    '*  ===================================================================
    '    '*  This function returns the values of the mass flow of air entrained
    '    '*  from the lower layer into the fire plume.
    '    '*
    '    '*  Arguments passed to the function are:
    '    '*      tim = time (sec)
    '    '*      Z1 = layer height above the floor (m)
    '    '*      QMax = max heat release that than be supported by air supply
    '    '*
    '    '*  Functions or subprocedures called:
    '    '*      Get_HRR
    '    '*      MassLoss_Object
    '    '*      Fire_Diameter
    '    '*      Virtual_Source
    '    '*  ===================================================================

    '    Dim Region, zo, diameter As Double
    '    Dim L As Single
    '    Dim i As Integer
    '    Dim Q, Qc As Double
    '    Dim MLoss, total As Double
    '    Dim NumberObjects1 As Short
    '    Dim ZF, massplume, Qtotal As Double
    '    Dim QP As Double
    '    Dim QI, QWall, QCeiling, QFloor As Double
    '    Static post, rcnone As Boolean

    '    If gb_first_time_vent = True Then 'If tim = 0 Then
    '        post = frmOptions1.optPostFlashover.Checked
    '        rcnone = frmOptions1.optRCNone.Checked
    '    End If

    '    'z1 cannot be negative
    '    If Z1 <= 0 Then
    '        Mass_Plume = 0
    '        Exit Function
    '    End If

    '    If Flashover = True And post = True Then

    '        If qmax > 0 Then
    '            Mass_Plume = 0.011 * qmax * ((Z1) / (qmax ^ (2 / 5))) ^ 0.566
    '        Else
    '            Mass_Plume = 0
    '        End If
    '        Exit Function

    '    Else
    '        NumberObjects1 = NumberObjects
    '    End If

    '    'determine the mass of air entrained into the plume
    '    Qtotal = 0

    '    For i = 1 To NumberObjects1

    '        'find heat release rate for each object
    '        QI = Get_HRR(i, tim, QP, Qburner, QFloor, QWall, QCeiling)

    '        If rcnone = False Then
    '            'this is a room corner test
    '            'hrr from the burner and the wall below the layer height
    '            'Q = QP
    '            Q = QI
    '        Else
    '            Q = QI
    '        End If

    '        Qtotal = Qtotal + QI 'total to use in plume equations
    '        'qmax = qmax + QI 'total theoretical hrr

    '        'find mass loss rate for each object
    '        'If EnergyYield(i) > 0 Then
    '        If NewEnergyYield(i) > 0 Then
    '            'MLoss = Q / (EnergyYield(i) * 1000)
    '            MLoss = Q / (NewEnergyYield(i) * 1000) 'used to get plume diameter
    '        Else
    '            MLoss = 0
    '        End If

    '        'check to see if heat release will be limited by burning
    '        If Qtotal > qmax Then
    '            'xxx If i > 1 Then Q = Qtotal - qmax  '19/7/2011 commented out
    '            Qtotal = qmax
    '            'individual object cannot be greater than QMax commented out
    '            'xxx If Q > qmax Then Q = qmax '19/7/2011
    '        End If

    '        'find height of layer above the base of the fire
    '        ZF = Z1 - FireHeight(i)
    '        If ZF < 0 Then ZF = 0

    '        If FireLocation(i) = 0 Then
    '            'fire in centre of room
    '            Select Case PlumeModel

    '                Case "General"
    '                    'general correlation from strong plume theory for buoyant plume
    '                    'and Delichatsios's correlations for the flaming regions

    '                    If ObjectMLUA(2, i) > 0 Then
    '                        'hrrua
    '                        diameter = 2 * Sqrt((Q / ObjectMLUA(2, i)) / PI)
    '                    Else

    '                        'get mass loss rate for object
    '                        Dim mlua As Single = ObjectMLUA(0, i) * (Target(fireroom, stepcount - 1) - Target(fireroom, 1)) + ObjectMLUA(1, i)

    '                        If ObjectMLUA(0, i) > 0 Or ObjectMLUA(1, i) > 0 Then
    '                        Else
    '                            mlua = ObjectMLUA(1, i) / NewEnergyYield(i) / 1000 'not specified for object
    '                        End If

    '                        'find fire diameter for each object
    '                        diameter = Fire_Diameter(MLoss, mlua)
    '                        'diameter = Fire_Diameter(MLoss, MLUnitArea)
    '                    End If

    '                    'call function to get virtual source
    '                    zo = Virtual_Source(Q, diameter)

    '                    'call function to get flame height
    '                    L = Flame_Height(diameter, Q) 'for a fire in centre of room

    '                    Qc = (1 - NewRadiantLossFraction(i)) * Q

    '                    If Z1 - FireHeight(i) - zo >= L Then
    '                        'buoyant plume, classical theory
    '                        massplume = EntrainmentCoefficient * Qc ^ (1 / 3) * (ZF - zo) ^ (5 / 3)

    '                        'Heskestad's strong plume
    '                        massplume = 0.071 * Qc ^ (1 / 3) * (ZF - zo) ^ (5 / 3) * (1 + 0.026 * (Qc) ^ (2 / 3) * (ZF - zo) ^ (-5 / 3))

    '                    ElseIf ZF - zo >= 5.16 * diameter Then
    '                        'flame, Delichatsios
    '                        massplume = 0.18 * ((ZF - zo)) ^ (5 / 2)
    '                    ElseIf ZF - zo >= diameter Then
    '                        'flame, Delichatsios
    '                        massplume = 0.93 * ((ZF - zo) / diameter) ^ (3 / 2) * diameter ^ (5 / 2)
    '                    ElseIf ZF - zo > 0 Then
    '                        'flame, Delichatsios
    '                        massplume = 0.86 * ((ZF - zo) / diameter) ^ (1 / 2) * diameter ^ (5 / 2)
    '                    Else
    '                        massplume = 0
    '                    End If

    '                    If ZF <= 0 Then
    '                        massplume = 0
    '                    End If

    '                Case "McCaffrey"
    '                    'McCaffrey's correlation

    '                    If Q > 0 Then
    '                        Region = (ZF) / (Q ^ (2 / 5))
    '                    Else
    '                        massplume = 0
    '                    End If

    '                    If ZF >= 0 And Q > 0 Then
    '                        If Region < 0.08 Then
    '                            'flaming region
    '                            massplume = 0.011 * Q * (Region) ^ 0.566
    '                        ElseIf Region >= 0.2 Then
    '                            'buoyant plume
    '                            massplume = 0.124 * Q * (Region) ^ 1.895
    '                        Else
    '                            'intermittent region
    '                            massplume = 0.026 * Q * (Region) ^ 0.909
    '                        End If
    '                    Else
    '                        massplume = 0
    '                    End If
    '            End Select


    '        ElseIf FireLocation(i) = 1 Then
    '            'fire against wall
    '            'massplume = massplume * 0.77 'following mowrer and williamson
    '            'Total = Total * .63   'zukoski
    '            massplume = 0.045 * Q ^ (1 / 3) * ZF ^ (5 / 3)
    '        ElseIf FireLocation(i) = 2 Then
    '            'entrainment in corner reduced
    '            'massplume = massplume * 0.59 'following mowrer and williamson
    '            'Total = Total * .4   'zukoski
    '            massplume = 0.028 * Q ^ (1 / 3) * ZF ^ (5 / 3)
    '        End If

    '        If UT > LT And UT < 400 Then
    '            If (1 - NewRadiantLossFraction(i)) * Qtotal / (SpecificHeat_air * (UT - LT)) < total Then
    '                massplume = (1 - NewRadiantLossFraction(i)) * Qtotal / (SpecificHeat_air * (UT - LT))
    '            End If
    '        ElseIf UT = LT Then
    '            'total = 0 '20/5/2102
    '        End If

    '        total = total + massplume


    '    Next i

    '    Mass_Plume = total

    'End Function
    Function Mass_Plume_2012(ByVal tim As Double, ByVal Z1 As Double, ByVal qmax As Double, ByVal UT As Double, ByVal LT As Double) As Double
        '*  ===================================================================
        '*  This function returns the values of the mass flow of air entrained
        '*  from the lower layer into the fire plume.
        '*
        '*  Arguments passed to the function are:
        '*      tim = time (sec)
        '*      Z1 = layer height above the floor (m)
        '*      QMax = max heat release that than be supported by air supply
        '*
        '*  Functions or subprocedures called:
        '*      Get_HRR
        '*      MassLoss_Object
        '*      Fire_Diameter
        '*      Virtual_Source
        '*  ===================================================================

        'this function will use the Heskestad plume model for the buoyant plume where it is valid,
        'otherwise the McCaffrey correlations will be used, except for wood crib model where only MCC is used.

        If FuelResponseEffects = True Then
            'if CLT/kinetic then this will be applicable for the contents + linings but not larger than Qmax sent here.
            Mass_Plume_2012 = Mass_Plume_fuelresponse(tim, Z1, qmax, UT, LT)
            Exit Function
        End If

        Dim Region As Double = 0
        Dim zo As Double = 0
        Dim diameter As Double = 0
        Dim L As Single = 0
        Dim i As Integer = 0
        Dim Q As Double = 0
        Dim Qc As Double = 0
        Dim MLoss As Double = 0
        Dim total As Double = 0
        Dim NumberObjects1 As Short = 0
        Dim ZF As Double = 0
        Dim massplume As Double = 0
        Dim Qtotal As Double = 0
        Dim QP As Double = 0
        Dim QI As Double = 0
        Dim QWall As Double = 0
        Dim QCeiling As Double = 0
        Dim QFloor As Double = 0

        'Static post, rcnone As Boolean
        Static post As Boolean

        Try

            If gb_first_time_vent = True Then 'If tim = 0 Then
                post = frmOptions1.optPostFlashover.Checked
                RCNone = frmOptions1.optRCNone.Checked
            End If

            'z1 cannot be negative
            If Z1 <= 0 Then
                Mass_Plume_2012 = 0
                Exit Function
            End If

            If Flashover = True And post = True Then

                If qmax > 0 Then
                    'Mass_Plume_2012 = 0.011 * qmax * ((Z1) / (qmax ^ (2 / 5))) ^ 0.566
                    If tim - flashover_time < 15 Then
                        If qmax > HRRatFO Then
                            'restricts the rate of increase to no more than 0.1MW per sec
                            qmax = Min(qmax, HRRatFO + 1500 * (tim - flashover_time) / 15)
                        End If
                    End If

                    'find height of layer above the base of the fire
                    ZF = Z1 - FireHeight(1)
                    If ZF < 0 Then ZF = 0

                    If Z1 <= 0.005 Then
                        ZF = 0 'a minimum layer height of 5mm is enforced for numerical reasons
                        Mass_Plume_2012 = 0
                        Exit Function
                    End If

                    'McCaffrey
                    Mass_Plume_2012 = 0.011 * qmax * ((ZF) / (qmax ^ (2 / 5))) ^ 0.566

                    'Mass_Plume_2012 = Mass_Plume_2012 * 2
                    'Thomas - flaming region of large fires
                    'Mass_Plume_2012 = 0.188 * 2 * PI * Sqrt(RoomFloorArea(fireroom) / PI) * ZF ^ (3 / 2)

                Else
                    Mass_Plume_2012 = 0
                End If

                'Mass_Plume_2012 = 2 * Mass_Plume_2012 'noisy plume option

                Exit Function

            Else
                NumberObjects1 = NumberObjects
            End If

            'determine the mass of air entrained into the plume
            Qtotal = 0
            total = 0

            For i = 1 To NumberObjects1
                If FuelResponseEffects = True Then
                    QI = qmax 'only works 1 item for now
                Else
                    'find heat release rate for each object
                    QI = Get_HRR(i, tim, QP, Qburner, QFloor, QWall, QCeiling)
                End If

                If useCLTmodel = True Then
                    Q = QI + QFloor + QWall + QCeiling
                    Qtotal = Qtotal + Q
                    'If QWall > 0 Then Stop
                Else
                    Q = QI
                    Qtotal = Qtotal + QI 'total to use in plume equations
                End If

                'find mass loss rate for each object
                If NewEnergyYield(i) > 0 Then
                    MLoss = Q / (NewEnergyYield(i) * 1000) 'used to get plume diameter
                Else
                    MLoss = 0
                End If

                'check to see if heat release will be limited by burning
                If Qtotal > qmax Then
                    Qtotal = qmax
                    '************** 4/9/2014
                    'Q = qmax
                End If

                'find height of layer above the base of the fire
                ZF = Z1 - FireHeight(i)
                If ZF < 0 Then ZF = 0

                If Z1 <= 0.005 Then
                    ZF = 0 'a minimum layer height of 5mm is enforced for numerical reasons
                    Mass_Plume_2012 = 0
                    Exit Function
                End If

                If FireLocation(i) = 0 Then
                    'fire in centre of room

                    If ObjectMLUA(2, i) > 0 Then
                        'hrrua
                        diameter = 2 * Sqrt((Q / ObjectMLUA(2, i)) / PI)
                    Else

                        'get mass loss rate for object
                        Dim mlua As Single = ObjectMLUA(0, i) * (Target(fireroom, stepcount - 1) - Target(fireroom, 1)) + ObjectMLUA(1, i)

                        'find fire diameter for each object
                        If mlua > 0 Then diameter = Fire_Diameter(MLoss, mlua) '19082015
                    End If

                    'call function to get virtual source
                    zo = Virtual_Source(Q, diameter)

                    'call function to get flame height
                    L = Flame_Height(diameter, Q) 'for a fire in centre of room

                    Qc = (1 - NewRadiantLossFraction(i)) * Q

                    If Qc > 0 And ZF >= L And RoomHeight(fireroom) - FireHeight(i) > 2 * L And post = False Then '2018 don't use heskestad with the wood crib model
                        'If ZF >= L Then

                        'Heskestad's strong plume
                        massplume = 0.071 * Qc ^ (1 / 3) * (ZF - zo) ^ (5 / 3) * (1 + 0.026 * (Qc) ^ (2 / 3) * (ZF - zo) ^ (-5 / 3))

                    Else
                        'mCaffrey

                        If Q > 0 Then
                            Region = (ZF) / (Q ^ (2 / 5))
                        Else
                            massplume = 0
                        End If

                        If ZF >= 0 And Q > 0 Then
                            If Region < 0.08 Then
                                'flaming region
                                massplume = 0.011 * Q * (Region) ^ 0.566
                            ElseIf Region >= 0.2 Then
                                'buoyant plume
                                massplume = 0.124 * Q * (Region) ^ 1.895
                            Else
                                'intermittent region
                                massplume = 0.026 * Q * (Region) ^ 0.909
                            End If
                        Else
                            massplume = 0
                        End If

                    End If

                ElseIf FireLocation(i) = 1 Then
                    'fire against wall
                    massplume = 0.045 * Q ^ (1 / 3) * ZF ^ (5 / 3)

                    'If Q > 0 Then massplume = Q * 1 * 0.076 * (ZF / Q ^ (2 / 5)) ^ (5 / 3)
                ElseIf FireLocation(i) = 2 Then
                    'entrainment in corner reduced
                    massplume = 0.028 * Q ^ (1 / 3) * ZF ^ (5 / 3)
                End If

                If ObjectWindEffect(i) > 1 Then
                    massplume = 2 * massplume 'noisy plume option
                End If

                If UT > LT And UT < 400 Then
                    'If (1 - NewRadiantLossFraction(i)) * Qtotal / (SpecificHeat_air * (UT - LT)) < total Then
                    If (1 - NewRadiantLossFraction(i)) * Q / (SpecificHeat_air * (UT - LT)) < massplume Then
                        massplume = (1 - NewRadiantLossFraction(i)) * Q / (SpecificHeat_air * (UT - LT))
                        'this is eqn 46 in NIST TN 1299 - upper limit on entrainment such that no more
                        'is entrained that would be needed to penetrate the upper layer
                        'stated to give better match with exp dat for the inital descent of the smoke layer
                    End If
                Else
                    'Stop
                End If

                total = total + massplume
            Next i

            Mass_Plume_2012 = total

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in Module1.vb Mass_Plume_fuelresponse")
        End Try

    End Function
    Function Mass_Plume_fuelresponse(ByVal tim As Double, ByVal Z1 As Double, ByVal qmax As Double, ByVal UT As Double, ByVal LT As Double) As Double
        '*  ===================================================================
        '*  This function returns the values of the mass flow of air entrained
        '*  from the lower layer into the fire plume.
        '*
        '*  Arguments passed to the function are:
        '*      tim = time (sec)
        '*      Z1 = layer height above the floor (m)
        '*      QMax = max heat release that than be supported by air supply
        '*
        '*  Functions or subprocedures called:
        '*      Get_HRR
        '*      MassLoss_Object
        '*      Fire_Diameter
        '*      Virtual_Source
        '*  ===================================================================

        'this function will use the Heskestad plume model for the buoyant plume where it is valid,
        'otherwise the McCaffrey correlations will be used.

        Dim Region As Double = 0
        Dim zo As Double = 0
        Dim diameter As Double = 0
        Dim L As Single = 0
        Dim i As Integer = 0
        Dim Q As Double = 0
        Dim Qc As Double = 0
        Dim MLoss As Double = 0
        Dim total As Double = 0
        Dim NumberObjects1 As Short = 0
        Dim ZF As Double = 0
        Dim massplume As Double = 0
        Dim Qtotal As Double = 0
        Dim QP As Double = 0
        Dim QI As Double = 0
        Dim QWall As Double = 0
        Dim QCeiling As Double = 0
        Dim QFloor As Double = 0

        'Static post, rcnone As Boolean
        Static post As Boolean

        Try
            ' If tim = 300 Then Stop
            If gb_first_time_vent = True Then 'If tim = 0 Then
                post = frmOptions1.optPostFlashover.Checked
                RCNone = frmOptions1.optRCNone.Checked
            End If

            'z1 cannot be negative
            If Z1 <= 0 Then
                Mass_Plume_fuelresponse = 0
                Exit Function
            End If

            NumberObjects1 = NumberObjects

            'determine the mass of air entrained into the plume
            Qtotal = 0
            total = 0

            For i = 1 To NumberObjects1

                'find heat release rate for each object
                QI = Get_HRR(i, tim, QP, Qburner, QFloor, QWall, QCeiling)
                ' If tim > 150 Then Stop

                If i = 1 Then QI = QI + QWall + QCeiling
                Q = QI

                Qtotal = Qtotal + QI 'total to use in plume equations

                'find mass loss rate for each object
                'If NewEnergyYield(i) > 0 Then
                '    MLoss = Q / (NewEnergyYield(i) * 1000) 'used to get plume diameter
                'Else
                '    MLoss = 0
                'End If

                'dummy = MassLoss_ObjectwithFuelResponse(i, tim, Qburner, QFloor, QWall, QCeiling, o2lower, ltemp, mplume, incidentflux, 0) * postSoot


                'check to see if heat release will be limited by burning
                If Qtotal > qmax Then
                    Qtotal = qmax
                End If

                'find height of layer above the base of the fire
                ZF = Z1 - FireHeight(i)
                If ZF < 0 Then ZF = 0

                If FireLocation(i) = 0 Then
                    'fire in centre of room
                    'diameter = ObjectPoolDiameter(i)

                    If ObjectPyrolysisOption(i) = 1 Then 'pool fire
                        'reduce the diameter accounting for the fuel area shrinkage
                        diameter = Sqrt(ObjectPoolDiameter(i) ^ 2 * FuelBurningRate(5, fireroom, stepcount))

                    Else   'user data and wood cribs
                        'equiv diameter based on top surface
                        diameter = Sqrt(4 * ObjLength(i) * ObjWidth(i) / PI)
                        If stepcount > 1 Then diameter = Sqrt(diameter ^ 2 * FuelBurningRate(5, fireroom, stepcount - 1))

                    End If

                    'call function to get virtual source
                    zo = Virtual_Source(Q, diameter)

                    'call function to get flame height
                    L = Flame_Height(diameter, Q) 'for a fire in centre of room

                    Qc = (1 - NewRadiantLossFraction(i)) * Q

                    If Qc > 0 And ZF >= L And RoomHeight(fireroom) - FireHeight(i) > 2 * L Then
                        'If ZF >= L Then

                        'Heskestad's strong plume
                        massplume = 0.071 * Qc ^ (1 / 3) * (ZF - zo) ^ (5 / 3) * (1 + 0.026 * (Qc) ^ (2 / 3) * (ZF - zo) ^ (-5 / 3))

                    Else
                        'mCaffrey
                        If Q > 0 Then
                            Region = (ZF) / (Q ^ (2 / 5))
                        Else
                            massplume = 0
                        End If

                        If ZF >= 0 And Q > 0 Then
                            If Region < 0.08 Then
                                'flaming region
                                massplume = 0.011 * Q * (Region) ^ 0.566
                            ElseIf Region >= 0.2 Then
                                'buoyant plume
                                massplume = 0.124 * Q * (Region) ^ 1.895
                            Else
                                'intermittent region
                                massplume = 0.026 * Q * (Region) ^ 0.909
                            End If
                        Else
                            massplume = 0
                        End If

                    End If

                ElseIf FireLocation(i) = 1 Then
                    'fire against wall
                    massplume = 0.045 * Q ^ (1 / 3) * ZF ^ (5 / 3)
                ElseIf FireLocation(i) = 2 Then
                    'entrainment in corner reduced
                    massplume = 0.028 * Q ^ (1 / 3) * ZF ^ (5 / 3)
                End If

                If ObjectWindEffect(i) > 1 Then
                    massplume = 2 * massplume 'noisy plume option
                End If

                If UT > LT And UT < 400 Then
                    If (1 - NewRadiantLossFraction(i)) * Q / (SpecificHeat_air * (UT - LT)) < massplume Then
                        massplume = (1 - NewRadiantLossFraction(i)) * Q / (SpecificHeat_air * (UT - LT))
                        'this is eqn 46 in NIST TN 1299 - upper limit on entrainment such that no more
                        'is entrained that would be needed to penetrate the upper layer
                        'stated to give better match with exp dat for the inital descent of the smoke layer
                    End If
                Else
                    'Stop
                End If

                total = total + massplume
            Next i
            Mass_Plume_fuelresponse = total

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in Module1.vb Mass_Plume_fuelresponse")
        End Try

    End Function
    Function MassLoss_ObjectwithCLT(ByVal id As Integer, ByVal tim As Double, ByRef mburner As Double, ByRef mFloor As Double, ByRef mWall As Double, ByRef mCeiling As Double) As Double
        '*  ===================================================================
        '*  This function return the value of the fuel mass loss rate for wood cribs and CLT elements
        '*
        '*  tim = time (sec)
        '*  ID = object identifier
        '*  call  If Flashover = True And g_post = True Then
        '*  ===================================================================

        mWall = 0
        mCeiling = 0
        mFloor = 0

        If IntegralModel = True Then
            MassLoss_ObjectwithCLT = MassLoss_Total_pluswood(tim, mWall, mCeiling) 'kg/s contents + wall + ceiling
        ElseIf KineticModel = True Then
            MassLoss_ObjectwithCLT = MassLoss_Total_Kinetic(tim, mWall, mCeiling) 'kg/s contents + wall + ceiling
        End If

    End Function
    '    Function MassLoss_ObjectwithFuelResponse_delete(ByVal id As Integer, ByVal tim As Double, ByRef Qburner As Double, ByVal O2lower As Double, ByVal ltemp As Double, ByVal mplume As Double, ByVal incidentflux As Double, ByRef burningrate As Double) As Double
    '        '*  ===================================================================
    '        '*  This function return the value of the fuel mass loss rate for a
    '        '*  specified burning object at a specified time including fuel response effects
    '        '*
    '        '*  tim = time (sec)
    '        '*  ID = object identifier
    '        '*
    '        '*  ===================================================================
    '        ' incidentflux includes contribution from gas layer, room surfaces, and point source fire onto the surface of the floor
    '        'for heptane fuel

    '        Dim QW As Double
    '        Dim k As Integer
    '        QW = 0
    '        QWall = 0
    '        QCeiling = 0
    '        QFloor = 0
    '        Dim Leff As Double = ObjectLHoG(id) * 1000 'kJ/kg heptane on water, effective heat of gasification from Mizukami et al FSJ (2016)
    '        Dim ec As Double = EmissionCoefficient 'emission coefficient for heptane
    '        'Dim r = S * 0.2313 'stoichiometric oxygen to fuel mass ratio for heptane
    '        'Dim r = 3.513 'stoichiometric oxygen to fuel mass ratio for heptane
    '        'Dim S As Double = 9.5 'stoichiometric air to fuel mass ratio for PU foam
    '        'Dim S As Double = 15.1 'stoichiometric air to fuel mass ratio for heptane
    '        Dim S As Double = StoichAFratio 'stoichiometric air to fuel mass ratio 

    '        StoichAFratio = 0
    '        If StoichAFratio = 0 Then
    '            S = EnergyYield(id) / deltaHair 'stoichiometric air to fuel mass ratio 
    '        End If

    '        Dim r = S * O2Ambient 'stoichiometric oxygen to fuel mass ratio for heptane

    '        Dim FBMLR As Double = 0
    '        Dim FBMLRuser As Double = 0
    '        Dim diameter As Double = ObjectPoolDiameter(id)
    '        Dim ventilation_contrib As Double = 0
    '        Dim thermal_contrib As Double = 0
    '        Dim peak As Double = 0

    '        Try


    '            If ObjectPyrolysisOption(id) = 0 Then 'use this for user freeburn MLR data provided
    '                If EnergyYield(id) > 0 Then

    '                    'Static mlrpeak As Double 'find peak MLR from free burn data
    '                    'If mlrpeak = 0 Then
    '                    '    For i = 1 To NumberDataPoints(id) - 1
    '                    '        If MLRData(2, i + 1, id) > mlrpeak Then
    '                    '            mlrpeak = MLRData(2, i + 1, id)

    '                    '        End If
    '                    '    Next
    '                    'End If
    '                    'peak = mlrpeak

    '                    'get free burn item MLR
    '                    Dim A, B, C As Double
    '                    For i = 1 To NumberDataPoints(id) - 1
    '                        If tim >= (MLRData(1, i, id) + ObjectIgnTime(id)) Then
    '                            If tim < (MLRData(1, i + 1, id) + ObjectIgnTime(id)) Then
    '                                A = tim - (MLRData(1, i, id) + ObjectIgnTime(id))
    '                                B = MLRData(1, i + 1, id) - MLRData(1, i, id) 'sec
    '                                C = MLRData(2, i + 1, id) - MLRData(2, i, id) 'kw
    '                                FBMLRuser = MLRData(2, i, id) + (A / B) * C
    '                                If FBMLRuser < 0 Then FBMLRuser = 0
    '                                FBMLR = FBMLRuser 'kg/s
    '                                Exit For
    '                            End If
    '                        End If
    '                    Next
    '                End If
    '            End If

    '            'this only applies to pool fire option
    '            If ObjectPyrolysisOption(id) = 1 Then



    '            ElseIf ObjectPyrolysisOption(id) = 0 Then
    '                If TotalFuel(stepcount - 1) > ObjectMass(id) Then
    '                    If CDec(tim) Mod CDec(Timestep) = 0 Then
    '                        FuelBurningRate(0, fireroom, stepcount) = FBMLR 'free burn mass loss rate
    '                    End If
    '                    Exit Function
    '                End If
    '            End If

    '            Dim Pathlength As Double = 0


    '            Dim ef As Double = 0 'flame emissivity

    '            If ObjectPyrolysisOption(id) = 1 Then 'poolfire equations
    '                ec = 1.1 / 0.6 'heptane
    '                FBMLRuser = ObjectPoolFBMLR(id) * (1 - Exp(-ec * 0.6 * ObjectPoolDiameter(id))) 'kg/m2/s
    '                FBMLR = FBMLRuser * 0.25 * PI * ObjectPoolDiameter(id) ^ 2 'kg/s

    '                'new diameter based on AFB/AF from the previous timestep
    '                If stepcount > 1 Then diameter = Sqrt(diameter ^ 2 * FuelBurningRate(5, fireroom, stepcount - 1))
    '                Pathlength = 0.6 * diameter
    '                ef = (1 - Exp(-ec * Pathlength)) 'flame emissivity

    '                If stepcount * Timestep < ObjectPoolRamp(id) Then
    '                    FBMLR = stepcount * Timestep * FBMLR / ObjectPoolRamp(id) 'ramping function
    '                End If

    '                If TotalFuel(stepcount - 1) > ObjectPoolDensity(id) * ObjectPoolVolume(id) / 1000 Then
    '                    If CDec(tim) Mod CDec(Timestep) = 0 Then
    '                        FuelBurningRate(0, fireroom, stepcount) = FBMLR 'free burn mass loss rate
    '                    End If
    '                    Exit Function
    '                End If

    '            ElseIf ObjectPyrolysisOption(id) = 0 Then
    '                'equiv diameter based on top surface
    '                diameter = Sqrt(4 * ObjLength(id) * ObjWidth(id) / PI)
    '                If stepcount > 1 Then diameter = Sqrt(diameter ^ 2 * FuelBurningRate(5, fireroom, stepcount - 1))

    '                Pathlength = 0.6 * diameter
    '                ef = (1 - Exp(-ec * Pathlength)) 'flame emissivity
    '            End If

    '            Dim Tf As Double
    '            Dim cp As Double = cplx(fireroom) 'specific heat of lower layer gases kJ/kg/K
    '            Dim Qexternal As Double = 0
    '            Dim Qext As Double = 0 'net heat feedback to the nonflaming area
    '            Dim Qextb As Double = 0 'net heat feedback to the flaming area
    '            Dim GER As Double = 0
    '            Dim AFB As Double = 0 'm2 burning surface area of the fuel
    '            Dim AF As Double = 0 'm2 total surface area of the fuel 

    '            If ObjectPyrolysisOption(id) = 0 Then
    '                'user data
    '                ' AF = 1000 * EnergyYield(id) * peak / ObjectMLUA(2, 1) 'm2
    '                AF = 1000 * EnergyYield(id) * FBMLR / ObjectMLUA(2, 1) 'm2

    '                'MLR data, represent by rectangular object W, L, H
    '                'AF = ObjWidth(id) * ObjLength(id) + 2 * (ObjHeight(id) * ObjWidth(id)) + 2 * (ObjHeight(id) * ObjLength(id)) 'm2 total surface area of the fuel excludes botton surface

    '            ElseIf ObjectPyrolysisOption(id) = 1 Then
    '                'pool fire
    '                AF = PI * ObjectPoolDiameter(id) ^ 2 / 4 'm2 total surface area of the fuel 
    '            ElseIf ObjectPyrolysisOption(id) = 2 Then
    '                'wood cribs to do
    '                AF = PI * ObjectPoolDiameter(id) ^ 2 / 4 'm2 total surface area of the fuel 
    '            End If

    '            ventilation_contrib = FBMLR * O2lower / O2Ambient 'kg/s
    '            ' If stepcount = 180 Then Stop

    '            Dim masslosstempold As Double
    '            Dim masslosstemp As Double = 0
    '            If stepcount > 1 Then masslosstemp = FuelBurningRate(3, fireroom, stepcount - 1)

    'here:

    '            If mplume > 0 Then
    '                GER = EnergyYield(id) * masslosstemp / (13.1 * mplume * O2lower)
    '            End If

    '            Dim delta As Double = 0
    '            Dim delta2 As Double
    '            If GER < 1 - delta Then
    '                AFB = AF
    '            ElseIf GER > 1 + delta Then
    '                AFB = (O2lower * mplume / O2Ambient) / S / (FBMLR / AF * O2lower / O2Ambient + (1 - ef) * incidentflux / Leff)


    '            Else
    '                'linear transition
    '                AFB = (O2lower * mplume / O2Ambient) / S / (FBMLR / AF * O2lower / O2Ambient + (1 - ef) * incidentflux / Leff)
    '                delta2 = (AF - AFB) * (GER - (1 - delta)) / (2 * delta)
    '                AFB = AF - delta2
    '            End If

    '            If AFB > AF Then Stop

    '            'ignore fuel shrinkage for now
    '            'AFB = AF

    '            'If GER > 1 Then
    '            '    'AFB = (O2lower * mplume / O2Ambient) / S / (ventilation_contrib / AF + (1 - ef) * incidentflux / Leff)

    '            '    'possibly should be using AFB instead of AF on RHS and solve iteratively.
    '            '    AFB = (O2lower * mplume / O2Ambient) / S / (FBMLR / AF * O2lower / O2Ambient + (1 - ef) * incidentflux / Leff)

    '            '    'below introduces oscillation in the GER/pressure.
    '            '    'Dim RHS As Double = AFB
    '            '    'Dim AFBX As Double = AF
    '            '    'Dim count As Integer = 0
    '            '    'While Abs(RHS - AFBX) / RHS > 0.001
    '            '    '    AFBX = RHS
    '            '    '    'update ef
    '            '    '    ef = (1 - Exp(-ec * 0.6 * Sqrt(ObjectPoolDiameter(id) ^ 2 * AFBX / AF)))
    '            '    '    RHS = (O2lower * mplume / O2Ambient) / S / (FBMLR / AFBX * O2lower / O2Ambient + (1 - ef) * incidentflux / Leff)
    '            '    '    count = count + 1
    '            '    '    If count > 100 Then Exit While
    '            '    'End While
    '            '    'AFB = RHS

    '            'Else
    '            '    AFB = AF
    '            'End If
    '            If AF > 0 Then
    '                ventilation_contrib = AFB / AF * FBMLR * O2lower / O2Ambient 'kg/s 
    '            End If

    '            Qext = incidentflux * (AF - AFB)
    '            Qextb = (1 - ef) * incidentflux * AFB 'kW
    '            Qexternal = Qext + Qextb

    '            thermal_contrib = Qexternal / Leff 'kg/s

    '            masslosstempold = masslosstemp
    '            masslosstemp = ventilation_contrib + thermal_contrib 'kg/s
    '            k = k + 1
    '            If k < 3 Then GoTo here
    '            'If Abs(masslosstemp - masslosstempold) / masslosstemp > 0.015 Then
    '            '    GoTo here
    '            'End If

    '            If GER > 1 Then
    '                burningrate = O2lower * mplume / (O2Ambient * S)
    '            Else
    '                burningrate = ventilation_contrib + thermal_contrib 'same as MLR
    '            End If

    '            If FlameExtinctionModel = True And burningrate > 0 Then 'if using flame extinction model
    '                'extinction model
    '                Dim RHS As Double = 0
    '                Dim TTemp As Double
    '                'TTemp = (ltemp + InteriorTemp) / 2
    '                'TTemp = ltemp
    '                TTemp = InteriorTemp
    '                'TTemp = (InteriorTemp + 1.28 * uppertemp(fireroom, stepcount)) / 2.28

    '                'RHS = (1000 * EnergyYield(id) - Leff + cp * (ObjectPoolVapTemp(id) - (TTemp - 273)) + Qexternal / masslosstemp) / (1 + r / O2lower)

    '                'this should use Qext,b not Qext?
    '                RHS = (1000 * EnergyYield(id) - Leff + cp * (ObjectPoolVapTemp(id) - (TTemp - 273)) + Qextb / burningrate) / (1 + r / O2lower)

    '                'simplify
    '                'RHS = (1000 * EnergyYield(id) + cp * (ObjectPoolVapTemp(id) - (TTemp - 273))) / (1 + r / O2lower)

    '                Tf = RHS / cp + (TTemp - 273)
    '                If Tf < 1300 Then 'C
    '                    'extinction
    '                    burningrate = 0
    '                    ventilation_contrib = 0
    '                    masslosstemp = thermal_contrib
    '                End If
    '            End If

    '            MassLoss_ObjectwithFuelResponse = masslosstemp 'kg/s

    '            If CDec(tim) Mod CDec(Timestep) = 0 Then
    '                FuelBurningRate(0, fireroom, stepcount) = FBMLR 'free burn mass loss rate
    '                FuelBurningRate(1, fireroom, stepcount) = ventilation_contrib 'contrib to mass loss rate
    '                FuelBurningRate(2, fireroom, stepcount) = thermal_contrib 'contrib to mass loss rate
    '                FuelBurningRate(3, fireroom, stepcount) = ventilation_contrib + thermal_contrib 'total mass loss
    '                FuelBurningRate(4, fireroom, stepcount) = burningrate
    '                If AF > 0 Then
    '                    FuelBurningRate(5, fireroom, stepcount) = AFB / AF 'fuel area shrinkage ratio
    '                End If
    '                FuelMassLossRate(stepcount, fireroom) = masslosstemp
    '            End If
    '            'Else
    '            'MassLoss_ObjectwithFuelResponse = 0
    '            'End If

    '        Catch ex As Exception
    '            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in Module1.vb MassLoss_ObjectwithFuelResponse")
    '        End Try

    '    End Function
    Function MassLoss_ObjectwithFuelResponse(ByVal id As Integer, ByVal tim As Double, ByRef Qburner As Double, ByVal O2lower As Double, ByVal ltemp As Double, ByVal mplume As Double, ByVal incidentflux As Double, ByRef burningrate As Double) As Double
        '*  ===================================================================
        '*  This function return the value of the fuel mass loss rate for a
        '*  specified burning object at a specified time including fuel response effects
        '*
        '*  tim = time (sec)
        '*  ID = object identifier
        '*
        '*  ===================================================================
        ' incidentflux includes contribution from gas layer, room surfaces, and point source fire onto the surface of the floor
        'for heptane fuel

        'according to Utiskul's thesis the incidentflux passed to this routine should be the net radiant feedback to the fuel 
        'we are generally using the Target() variable which is incident?
        'QFloorAST(room, 3, stepcount)  is the net radiant flux to floor

        incidentflux = QFloorAST(fireroom, 0, stepcount) 'incident flux on floor includes radiation from gas, surfaces and fire/flame source
        'since the flame radiation is accounted for in the ventilation effect, try using an incident external radiation that only includes the gas and surface radiation

        ''incidentflux = QFloorAST(fireroom, 3, stepcount)  'net radiant heat flux to floor + reradiation term to get incident 
        ' incidentflux = incidentflux + StefanBoltzmann / 1000 * Surface_Emissivity(4, fireroom) * FloorTemp(fireroom, stepcount) ^ 4
        'get radiant heat fluxes to room surfaces -EXCLUDING fire source Q=0
        'Dim matz(4, 1) As Double
        'Dim prodz(4, 1) As Double
        'Call FourWallRad(fireroom, 0, CeilingTemp(fireroom, stepcount), Upperwalltemp(fireroom, stepcount), LowerWallTemp(fireroom, stepcount), FloorTemp(fireroom, stepcount), uppertemp(fireroom, stepcount), ltemp, layerheight(fireroom, stepcount), prodz, 0, 0, 0, 0, matz, CO2MassFraction(fireroom, stepcount, 1), H2OMassFraction(fireroom, stepcount, 1), CO2MassFraction(fireroom, stepcount, 2), H2OMassFraction(fireroom, stepcount, 2), UpperVolume(fireroom, stepcount))
        'incidentflux = matz(4, 1) 'incident radiant flux on floor including gas layer, excluding Q 

        Dim QW As Double
        Dim k As Integer
        QW = 0
        'QWall = 0
        'QCeiling = 0
        'QFloor = 0
        Dim Leff As Double = ObjectLHoG(id) * 1000 'kJ/kg heptane on water, effective heat of gasification from Mizukami et al FSJ (2016)
        Dim ec As Double = EmissionCoefficient 'emission coefficient for heptane
        'Dim r = S * 0.2313 'stoichiometric oxygen to fuel mass ratio for heptane
        'Dim r = 3.513 'stoichiometric oxygen to fuel mass ratio for heptane
        'Dim S As Double = 9.5 'stoichiometric air to fuel mass ratio for PU foam
        'Dim S As Double = 15.1 'stoichiometric air to fuel mass ratio for heptane
        Dim S As Double = StoichAFratio 'stoichiometric air to fuel mass ratio 

        StoichAFratio = 0
        If StoichAFratio = 0 Then
            S = EnergyYield(id) / deltaHair 'stoichiometric air to fuel mass ratio 
        End If

        Dim r = S * O2Ambient 'stoichiometric oxygen to fuel mass ratio for heptane

        Dim FBMLR As Double = 0
        Dim FBMLRuser As Double = 0
        Dim diameter As Double = ObjectPoolDiameter(id)
        Dim ventilation_contrib As Double = 0
        Dim thermal_contrib As Double = 0
        Dim peak As Double = 0
        Dim ef As Double = 0 'flame emissivity
        Dim Pathlength As Double = 0
        Dim AF As Double = 0 'm2 total surface area of the fuel 

        Try

            If ObjectPyrolysisOption(id) = 0 Then 'use this for user freeburn MLR data provided
                If EnergyYield(id) > 0 Then

                    'get free burn item MLR
                    Dim A, B, C As Double
                    For i = 1 To NumberDataPoints(id) - 1
                        If tim >= (MLRData(1, i, id) + ObjectIgnTime(id)) Then
                            If tim < (MLRData(1, i + 1, id) + ObjectIgnTime(id)) Then
                                A = tim - (MLRData(1, i, id) + ObjectIgnTime(id))
                                B = MLRData(1, i + 1, id) - MLRData(1, i, id) 'sec
                                C = MLRData(2, i + 1, id) - MLRData(2, i, id) 'kw
                                FBMLRuser = MLRData(2, i, id) + (A / B) * C
                                If FBMLRuser < 0 Then FBMLRuser = 0
                                FBMLR = FBMLRuser 'kg/s

                                Exit For
                            End If
                        End If
                    Next
                End If

                If TotalFuel(stepcount - 1) > ObjectMass(id) Then
                    If CDec(tim) Mod CDec(Timestep) = 0 Then
                        FuelBurningRate(0, fireroom, stepcount) = FBMLR 'free burn mass loss rate
                    End If
                    Exit Function
                End If

                'equiv diameter based on top surface
                diameter = Sqrt(4 * ObjLength(id) * ObjWidth(id) / PI)
                If stepcount > 1 Then diameter = Sqrt(diameter ^ 2 * FuelBurningRate(5, fireroom, stepcount - 1))

                Pathlength = 0.6 * diameter
                ef = (1 - Exp(-ec * Pathlength)) 'flame emissivity


            ElseIf ObjectPyrolysisOption(id) = 1 Then 'pool fire

                ec = 1.1 / 0.6 'heptane
                FBMLRuser = ObjectPoolFBMLR(id) * (1 - Exp(-ec * 0.6 * ObjectPoolDiameter(id))) 'kg/m2/s
                FBMLR = FBMLRuser * 0.25 * PI * ObjectPoolDiameter(id) ^ 2 'kg/s

                'new diameter based on AFB/AF from the previous timestep
                If stepcount > 1 Then diameter = Sqrt(diameter ^ 2 * FuelBurningRate(5, fireroom, stepcount - 1))
                Pathlength = 0.6 * diameter
                ef = (1 - Exp(-ec * Pathlength)) 'flame emissivity

                If stepcount * Timestep < ObjectPoolRamp(id) Then
                    FBMLR = stepcount * Timestep * FBMLR / ObjectPoolRamp(id) 'ramping function
                End If

                If TotalFuel(stepcount - 1) > ObjectPoolDensity(id) * ObjectPoolVolume(id) / 1000 Then
                    If CDec(tim) Mod CDec(Timestep) = 0 Then
                        FuelBurningRate(0, fireroom, stepcount) = FBMLR 'free burn mass loss rate
                    End If
                    Exit Function
                End If

            ElseIf ObjectPyrolysisOption(id) = 2 Then 'cribs
                'eqn 5.43 in Utiskul's thesis

                'Cw is the empirical wood crib coefficient given by Block [64] for some species of wood as 1.03 mg/cm1.5 for Ponderosa pine, 1.33 for Oak, and 0.88 for Sugar pine.
                Dim Cw As Double = 1.03 'mg cm^(-1.5) 'pine

                'Dim b As Double = Fuel_Thickness 'b is thickness dimension of a stick
                'b = 0.01905

                'Dim Aco As Double 'Aco is the cross-sectional area of the vertical crib shafts

                'Dim N As Integer = 5 'N  is the number of stick layers (N>1)

                'Dim Li As Double = 0.15 'stick length (m)
                'Dim Lj As Double = 0.15 'stick length (m)

                'Dim mi As Double = Li / b
                'Dim mj As Double = Lj / b

                'Dim ni As Integer = 4 'number of sticks per layer
                'Dim nj As Integer = 4 'number of sticks per layer



                Dim b As Double = Fuel_Thickness 'b is thickness dimension of a stick
                b = 0.067

                Dim Aco As Double 'Aco is the cross-sectional area of the vertical crib shafts

                Dim N As Integer = 10 'N  is the number of stick layers (N>1)

                Dim Li As Double = 0.9 'stick length (m)
                Dim Lj As Double = 0.9 'stick length (m)

                Dim mi As Double = Li / b
                Dim mj As Double = Lj / b

                Dim ni As Integer = 8 'number of sticks per layer
                Dim nj As Integer = 8 'number of sticks per layer


                Dim si As Double = (Li - nj * b) / (nj - 1) 's is the spacing between sticks
                Dim sj As Double = (Lj - ni * b) / (ni - 1) 's is the spacing between sticks
                Dim sp As Double = (si + sj) / 2

                Aco = si * sj * (nj - 1) * (ni - 1)

                'rectangular footprint
                If (N / 2) Mod 2 = 0 Then 'N is even number
                    AF = b ^ 2 * (mi * ni * (2 * N - 1) + N * (ni + nj + 2 * mj * nj - 2 * ni * nj) + nj * (3 * ni - nj))
                Else 'N is odd number
                    AF = b ^ 2 * ((N - 1) * (2 * mj + 1) * nj + ni * (1 + mi + 2 * nj + N + 2 * N * (mi - nj)))
                End If

                FBMLR = 0.968 * Cw * (b) ^ (-1 / 2) * (1 - Exp(-Sqrt(sp * b) * Aco / (0.02 * AF))) 'kg/s/m2
                'mg/s/m2?
                'FBMLR = FBMLR / 1000 / 1000 'kg/s/m2?

                FBMLR = FBMLR * AF 'kg/s
            End If


            Dim Tf As Double
            Dim cp As Double = cplx(fireroom) 'specific heat of lower layer gases kJ/kg/K
            Dim Qexternal As Double = 0
            Dim Qext As Double = 0 'net heat feedback to the nonflaming area
            Dim Qextb As Double = 0 'net heat feedback to the flaming area
            Dim GER As Double = 0
            Dim AFB As Double = 0 'm2 burning surface area of the fuel
            Dim AF2 As Double

            If ObjectPyrolysisOption(id) = 0 Then
                'user data
                ' AF = 1000 * EnergyYield(id) * peak / ObjectMLUA(2, 1) 'm2
                AF = 1000 * EnergyYield(id) * FBMLR / ObjectMLUA(2, 1) 'm2  'okay for 2D 

                'MLR data, represent by rectangular object W, L, H
                AF2 = ObjWidth(id) * ObjLength(id) + 2 * (ObjHeight(id) * ObjWidth(id)) + 2 * (ObjHeight(id) * ObjLength(id)) 'm2 total surface area of the fuel excludes botton surface
                AF2 = ObjWidth(id) * ObjLength(id) + 1 * (ObjHeight(id) * ObjWidth(id)) + 2 * (ObjHeight(id) * ObjLength(id)) 'top surface only + 50% of vertical surfaces see the fire
                'use the lesser of the two?
                If AF2 < AF Then
                    AF = AF2
                End If

            ElseIf ObjectPyrolysisOption(id) = 1 Then
                'pool fire
                AF = PI * ObjectPoolDiameter(id) ^ 2 / 4 'm2 total surface area of the fuel 
            ElseIf ObjectPyrolysisOption(id) = 2 Then
                'wood cribs to do

            End If

            ventilation_contrib = FBMLR * O2lower / O2Ambient 'kg/s
            ' If stepcount = 180 Then Stop

            Dim masslosstempold As Double
            Dim masslosstemp As Double = 0
            If stepcount > 1 Then masslosstemp = FuelBurningRate(3, fireroom, stepcount - 1)


here:

            If mplume > 0 Then
                If useCLTmodel = True And KineticModel = False Then
                    GER = EnergyYield(id) * masslosstemp / (13.1 * mplume * O2lower)
                Else
                    GER = EnergyYield(id) * (masslosstemp + WoodBurningRate(stepcount)) / (13.1 * mplume * O2lower)
                End If



            End If

            Dim delta As Double = 0
            Dim delta2 As Double
            If GER < 1 - delta Then
                AFB = AF
            ElseIf GER > 1 + delta Then
                AFB = (O2lower * mplume / O2Ambient) / S / (FBMLR / AF * O2lower / O2Ambient + (1 - ef) * incidentflux / Leff)
            Else
                'linear transition
                AFB = (O2lower * mplume / O2Ambient) / S / (FBMLR / AF * O2lower / O2Ambient + (1 - ef) * incidentflux / Leff)
                delta2 = (AF - AFB) * (GER - (1 - delta)) / (2 * delta)
                AFB = AF - delta2
            End If

            If AFB > AF Then AFB = AF '05012018

            'ignore fuel shrinkage for now
            'AFB = AF

            'If GER > 1 Then
            '    'AFB = (O2lower * mplume / O2Ambient) / S / (ventilation_contrib / AF + (1 - ef) * incidentflux / Leff)

            '    'possibly should be using AFB instead of AF on RHS and solve iteratively.
            '    AFB = (O2lower * mplume / O2Ambient) / S / (FBMLR / AF * O2lower / O2Ambient + (1 - ef) * incidentflux / Leff)

            '    'below introduces oscillation in the GER/pressure.
            '    'Dim RHS As Double = AFB
            '    'Dim AFBX As Double = AF
            '    'Dim count As Integer = 0
            '    'While Abs(RHS - AFBX) / RHS > 0.001
            '    '    AFBX = RHS
            '    '    'update ef
            '    '    ef = (1 - Exp(-ec * 0.6 * Sqrt(ObjectPoolDiameter(id) ^ 2 * AFBX / AF)))
            '    '    RHS = (O2lower * mplume / O2Ambient) / S / (FBMLR / AFBX * O2lower / O2Ambient + (1 - ef) * incidentflux / Leff)
            '    '    count = count + 1
            '    '    If count > 100 Then Exit While
            '    'End While
            '    'AFB = RHS

            'Else
            '    AFB = AF
            'End If
            If AF > 0 Then
                ventilation_contrib = AFB / AF * FBMLR * O2lower / O2Ambient 'kg/s 
            End If

            Qext = incidentflux * (AF - AFB)
            Qextb = (1 - ef) * incidentflux * AFB 'kW
            Qexternal = Qext + Qextb

            thermal_contrib = Qexternal / Leff 'kg/s

            masslosstempold = masslosstemp
            masslosstemp = ventilation_contrib + thermal_contrib 'kg/s
            k = k + 1
            If k < 3 Then GoTo here
            'If Abs(masslosstemp - masslosstempold) / masslosstemp > 0.015 Then
            '    GoTo here
            'End If

            If GER > 1 Then
                burningrate = O2lower * mplume / (O2Ambient * S)
            Else
                burningrate = ventilation_contrib + thermal_contrib 'same as MLR
            End If

            If FlameExtinctionModel = True And burningrate > 0 Then 'if using flame extinction model
                'extinction model
                Dim RHS As Double = 0
                Dim TTemp As Double
                'TTemp = (ltemp + InteriorTemp) / 2
                'TTemp = ltemp
                TTemp = InteriorTemp
                'TTemp = (InteriorTemp + 1.28 * uppertemp(fireroom, stepcount)) / 2.28

                'RHS = (1000 * EnergyYield(id) - Leff + cp * (ObjectPoolVapTemp(id) - (TTemp - 273)) + Qexternal / masslosstemp) / (1 + r / O2lower)

                'this should use Qext,b not Qext?
                RHS = (1000 * EnergyYield(id) - Leff + cp * (ObjectPoolVapTemp(id) - (TTemp - 273)) + Qextb / burningrate) / (1 + r / O2lower)

                'simplify
                'RHS = (1000 * EnergyYield(id) + cp * (ObjectPoolVapTemp(id) - (TTemp - 273))) / (1 + r / O2lower)

                Tf = RHS / cp + (TTemp - 273)
                If Tf < 1300 Then 'C
                    'extinction
                    burningrate = 0
                    ventilation_contrib = 0
                    masslosstemp = thermal_contrib
                End If
            End If

            MassLoss_ObjectwithFuelResponse = masslosstemp 'kg/s

            If CDec(tim) Mod CDec(Timestep) = 0 Then
                FuelBurningRate(0, fireroom, stepcount) = FBMLR 'free burn mass loss rate
                FuelBurningRate(1, fireroom, stepcount) = ventilation_contrib 'contrib to mass loss rate
                FuelBurningRate(2, fireroom, stepcount) = thermal_contrib 'contrib to mass loss rate
                FuelBurningRate(3, fireroom, stepcount) = ventilation_contrib + thermal_contrib 'total mass loss
                FuelBurningRate(4, fireroom, stepcount) = burningrate
                If AF > 0 Then
                    FuelBurningRate(5, fireroom, stepcount) = AFB / AF 'fuel area shrinkage ratio
                End If
                FuelMassLossRate(stepcount, fireroom) = masslosstemp
            End If
            'Else
            'MassLoss_ObjectwithFuelResponse = 0
            'End If

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in Module1.vb MassLoss_ObjectwithFuelResponse")
        End Try

    End Function


    Function MassLoss_Object(ByVal id As Integer, ByVal tim As Double, ByRef Qburner As Double, ByRef QFloor As Double, ByRef QWall As Double, ByRef QCeiling As Double) As Double
        '*  ===================================================================
        '*  This function return the value of the fuel mass loss rate for a
        '*  specified burning object at a specified time.
        '*
        '*  tim = time (sec)
        '*  ID = object identifier
        '*
        '*  Revised 6 December 1996 Colleen Wade
        '*  ===================================================================
        Dim QW As Double
        If useCLTmodel = True And IntegralModel = True Then
            'using CLT integral model
            If Flashover = True Then
                'surfaces burn after flashover
                MassLoss_Object = MassLoss_ObjectwithCLT(id, tim, Qburner, QFloor, QWall, QCeiling)
            Else
                'mass loss rate from theoretical hrr divided by heat of combustion
                MassLoss_Object = Get_HRR(id, tim, QW, Qburner, QFloor, QWall, QCeiling) / (NewEnergyYield(id) * 1000)
            End If

            Exit Function

        ElseIf useCLTmodel = True And KineticModel = True Then
            'new
            If Flashover = True Then
                'surfaces burn after flashover
                MassLoss_Object = MassLoss_ObjectwithCLT(id, tim, Qburner, QFloor, QWall, QCeiling)
            Else
                'mass loss rate from theoretical hrr divided by heat of combustion
                MassLoss_Object = Get_HRR(id, tim, QW, Qburner, QFloor, QWall, QCeiling) / (NewEnergyYield(id) * 1000)
                QCeiling = CeilingWoodMLR_tot(stepcount) * CeilingEffectiveHeatofCombustion(fireroom) * 1000
            End If

            Exit Function

        ElseIf useCLTmodel = True And g_post = False And integralmodel = False And kineticmodel = False Then
            MassLoss_Object = MassLoss_Total(tim)
            MassLoss_Object = MassLoss_Object + wall_char(stepcount, 1) + ceil_char(stepcount, 1)
            QWall = wall_char(stepcount, 1) * WallEffectiveHeatofCombustion(fireroom) * 1000
            QCeiling = ceil_char(stepcount, 1) * CeilingEffectiveHeatofCombustion(fireroom) * 1000

            Exit Function
        End If

        QW = 0
        QWall = 0
        QCeiling = 0
        QFloor = 0

        'If EnergyYield(id) > 0 Then
        If NewEnergyYield(id) > 0 Then
            If Flashover = True And g_post = True Then
                MassLoss_Object = MassLoss_Total(tim)

                'Dim mwall, mceiling As Double
                'MassLoss_Object = MassLoss_Total_pluswood(tim, mwall, mceiling) 'spearpoint quintiere
            Else
                'mass loss rate from theoretical hrr divided by heat of combustion
                MassLoss_Object = Get_HRR(id, tim, QW, Qburner, QFloor, QWall, QCeiling) / (NewEnergyYield(id) * 1000)
            End If
        Else
            MassLoss_Object = 0
        End If

    End Function

    Function MassLoss_Total(ByVal T As Double) As Double
        '*  ====================================================================
        '*  This function return the value of the total fuel mass loss rate for
        '*  all burning items at a specified time.
        '*
        '*  t = time (sec)
        '*
        '*  Revised 6 December 1996 Colleen Wade
        '   this function is only called when using the wood crib postflashover model
        '*  ====================================================================

        Dim dummy As Double
        Dim QCeiling, QW, QWall, QFloor, Qtemp As Double
        Dim totalarea, total As Double
        Dim i As Integer
        Dim mass As Double

        QW = 0 'baggage
        totalarea = 0
        burnmode = False

        mass = InitialFuelMass 'kg

        'if clt model then update this value
        If useCLTmodel = True And KineticModel = False Then
            mass = fuelmasswithCLT 'kg
        End If

        'get HRR for all burning objects
        Dim total1, vp, total2, total3 As Double

        If Flashover = True And g_post = True Then
            'in this case, we should only need this subroutine once per timestep

            'postflashover burning
            If TotalFuel(stepcount - 1) >= mass Then
                total = 0 'fuel is fully consumed
            Else


                'fuel surface control
                vp = 0.0000022 * Fuel_Thickness ^ (-0.6) 'wood crib fire regression rate m/s

                total1 = 4 / Fuel_Thickness * mass * vp * ((mass - TotalFuel(stepcount - 1)) / mass) ^ (0.5) 'kg/s

                'crib porosity control

                total2 = 0.00044 * (Stick_Spacing / Cribheight) * (mass / Fuel_Thickness)
                'total2 = 1000 'disable crib porosity control

                If total2 < total1 Then
                    total = total2 'use the lesser, crib porosity control
                Else
                    total = total1 'use the lesser, fuel surface control
                End If

                Qtemp = massplumeflow(stepcount - 1, fireroom) * O2MassFraction(fireroom, stepcount - 1, 2) * 13100
                total3 = 1 / 1000 * Qtemp / NewHoC_fuel  'kJ/s / kJ/g = g/s
                'total3 = 1 / 1000 * HeatRelease(fireroom, stepcount - 1, 2) / NewHoC_fuel  'kJ/s / kJ/g = g/s

                'kg/s

                If total3 < total Then
                    total = total3 'use the lesser, ventilation control
                    burnmode = True
                Else
                    ' Stop
                End If

                'better to apply excess fuel factor after determining mode of burning and then only if vent-limited
                If burnmode = True Then
                    total = total * ExcessFuelFactor
                Else

                End If

            End If
        Else
            'preflashover burning
            For i = 1 To NumberObjects
                If NewEnergyYield(i) <> 0 Then
                    dummy = Get_HRR(i, T, QW, Qburner, QFloor, QWall, QCeiling) / (NewEnergyYield(i) * 1000)
                End If
                total = total + dummy
            Next i

        End If

        MassLoss_Total = total


    End Function
    Function MassLoss_Total2(ByVal T As Double) As Double
        '*  ====================================================================
        '*  This function return the value of the total fuel mass loss rate for
        '*  all burning items at a specified time.
        '*
        '*  t = time (sec)
        '*
        '*  Revised 11/11/2017 Colleen Wade
        '   this function is only called when using the wood crib postflashover model
        '*  ====================================================================

        Dim dummy As Double
        Dim QCeiling, QW, QWall, QFloor, Qtemp As Double
        Dim totalarea, total As Double
        Dim i As Integer
        Dim mass As Double

        QW = 0 'baggage
        totalarea = 0
        burnmode = False

        mass = InitialFuelMass 'kg

        'if clt model then update this value
        If useCLTmodel = True Then
            mass = fuelmasswithCLT 'kg
        End If

        'get HRR for all burning objects
        Dim total1, vp, total2, total3 As Double

        If Flashover = True And g_post = True Then
            'in this case, we should only need this subroutine once per timestep

            'postflashover burning
            If TotalFuel(stepcount - 1) >= mass Then
                total = 0 'fuel is fully consumed
            Else

                'fuel surface control
                vp = 0.0000022 * Fuel_Thickness ^ (-0.6) 'wood crib fire regression rate m/s

                total1 = 4 / Fuel_Thickness * mass * vp * ((mass - TotalFuel(stepcount - 1)) / mass) ^ (0.5) 'kg/s

                'crib porosity control

                total2 = 0.00044 * (Stick_Spacing / Cribheight) * (mass / Fuel_Thickness)
                ' total2 = 1000 'disable crib porosity control

                If total2 < total1 Then
                    total = total2 'use the lesser, crib porosity control
                Else
                    total = total1 'use the lesser, fuel surface control
                End If

                'this should use the max Q given the oxygen in the plume flow.
                'this used only after flashover
                'doesn't work if we have surface area control at flashover
                Qtemp = massplumeflow(stepcount - 1, fireroom) * O2MassFraction(fireroom, stepcount - 1, 2) * 13100
                total3 = 1 / 1000 * Qtemp / NewHoC_fuel  'kJ/s / kJ/g = g/s
                'kg/s

                If total3 < total Then
                    total = ExcessFuelFactor * total3 'use the lesser, ventilation control
                    burnmode = True
                Else
                    'total = ExcessFuelFactor * total 'use the lesser, ventilation control
                    'If total > total3 Then total = total3
                End If


            End If
        Else
            'preflashover burning
            For i = 1 To NumberObjects
                If NewEnergyYield(i) <> 0 Then
                    dummy = Get_HRR(i, T, QW, Qburner, QFloor, QWall, QCeiling) / (NewEnergyYield(i) * 1000)
                End If
                total = total + dummy
            Next i

        End If

        MassLoss_Total2 = total


    End Function
    Sub Debond_test()

        Try

            Dim j As Integer
            Dim idr As Integer = fireroom
            Dim crittemp As Double = DebondTemp + 273 'debond temp

            'define variables
            Dim depth As Single = 0
            Dim maxtemp As Double = 0
            Dim NumberwallNodes As Integer
            Dim NumberCeilNodes As Integer

            j = stepcount 'current timestep
            NumberwallNodes = UBound(UWallNode, 2)
            NumberCeilNodes = UBound(CeilingNode, 2)

            Dim mydepth As Double
            Dim X(0 To NumberwallNodes) As Double
            Dim Y(0 To NumberwallNodes) As Double
            Dim T(0 To NumberwallNodes) As Double

            If CLTwallpercent > 0 Then

                For curve = 1 To NumberwallNodes
                    X(NumberwallNodes - curve + 1) = UWallNode(idr, curve, j) 'descending order
                    depth = (curve - 1) * WallThickness(idr) / (NumberwallNodes - 1)
                    Y(NumberwallNodes - curve + 1) = depth
                    T(NumberwallNodes - curve + 1) = j * Timestep
                Next
                ' depth by interpolation
                Interpolate_D(X, Y, NumberwallNodes, crittemp, mydepth)

                If mydepth > WallThickness(idr) Then
                    ' Dim Message As String = CStr(tim(stepcount, 1)) & " sec. Charring in the wall layer has penetrated full depth of the wall."
                ElseIf mydepth / 1000 > Lamella2 Then  'delaminate based on 1D heat conduction into surface - 1st layer

                    Dim Message As String = CStr(tim(stepcount, 1)) & " sec. Wall layer at " & Format(Lamella2, "0.000") & " m delaminates."
                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                    Lamella2 = Lamella2 + Lamella

                    If IntegralModel = True Or KineticModel = True Then
                        CLTwalldelamT = tim(stepcount, 1) - flashover_time
                    End If

                End If
            End If

            ReDim X(0 To NumberCeilNodes)
            ReDim Y(0 To NumberCeilNodes)
            ReDim T(0 To NumberCeilNodes)

            If CLTceilingpercent > 0 Then

                For curve = 1 To NumberCeilNodes
                    X(NumberCeilNodes - curve + 1) = CeilingNode(idr, curve, j) 'descending order
                    depth = (curve - 1) * CeilingThickness(idr) / (NumberCeilNodes - 1)
                    Y(NumberCeilNodes - curve + 1) = depth
                    T(NumberCeilNodes - curve + 1) = j * Timestep
                Next
                ' depth by interpolation
                Interpolate_D(X, Y, NumberCeilNodes, crittemp, mydepth)

                If mydepth > CeilingThickness(idr) Then
                    'Dim Message As String = CStr(tim(stepcount, 1)) & " sec. Charring in the ceiling layer has penetrated full depth of the ceiling."
                ElseIf mydepth / 1000 > Lamella1 Then  'delaminate based on 1D heat conduction into surface - 1st layer

                    Dim Message As String = CStr(tim(stepcount, 1)) & " sec. Ceiling layer at " & Format(Lamella1, "0.000") & " m delaminates."
                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                    Lamella1 = Lamella1 + Lamella

                    If IntegralModel = True Or KineticModel = True Then
                        CLTceildelamT = tim(stepcount, 1) - flashover_time
                    End If

                End If
            End If

        Catch ex As Exception
            Dim errormessage As String = "Exception in class " & GetCurrentMethod().DeclaringType().Name & " " & GetCurrentMethod().Name & " " & Err.Description & " (Line " & Err.Erl & ")"
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, errormessage)
        End Try
    End Sub
    Function MassLoss_Total_pluswood(ByVal T As Double, ByRef mwall As Double, ByRef mceiling As Double) As Double
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
            MassLoss_Total_pluswood = MLRsave
            Exit Function
        End If

        Dim dw As Double = 502 'kg/m3 density
        Dim kw As Double = 0.23 'W/mK thermal conductivity
        Dim cw As Double = 2200 'J/kgK specific heat


        Dim qcrit As Double = 16 'kW/m2 critical flux
        qcrit = CLTQcrit

        Dim Tign As Double = 384 'C ignition temp
        Tign = CLTigtemp

        Dim Lw As Double = 1.6 'kJ/g latent heat of gasification
        Lw = CLTLoG

        Dim Qff As Double = 17 'kW/m2 flame flux
        Qff = CLTflameflux

        'Douglas Fir across the grain
        'Dim dw As Double = 455 'kg/m3 density
        'Dim kw As Double = 0.8 'W/mK thermal conductivity
        'Dim cw As Double = 4000 'J/kgK specific heat
        'Dim qcrit As Double = 8.4 'kW/m2 critical flux
        'Dim Tign As Double = 258 'C igntion temp
        'Dim Lw As Double = 2.9 'kJ/g latent heat of gasification
        'Dim Qff As Double = 46 'kW/m2 flame flux

        'maple across the grain
        'Dim dw As Double = 742 'kg/m3 density
        'Dim kw As Double = 2.08 'W/mK thermal conductivity
        'Dim cw As Double = 7100 'J/kgK specific heat
        'Dim qcrit As Double = 3.8 'kW/m2 critical flux
        'Dim Tign As Double = 150 'C igntion temp
        'Dim Lw As Double = 3.5 'kJ/g latent heat of gasification
        'Dim Qff As Double = 36 'kW/m2 flame flux


        Dim QW, Qtemp As Double
        Dim QCeiling As Double = 0
        Dim QWall As Double = 0
        Dim QFloor As Double = 0
        Dim totalarea, total As Double
        Dim mass, wood_MLR As Double

        'this procedure only called if flashover is true and postflashover model is selected.
        'assume wood surface linings start burning at flashover

        QW = 0 'baggage
        totalarea = 0
        burnmode = False

        mass = InitialFuelMass 'kg wood cribs

        Dim thickness_ceil As Double = CeilingThickness(fireroom) / 1000 'm thick of wood
        Dim thickness_wall As Double = WallThickness(fireroom) / 1000  'm thick of wood
        Dim MLR_wood, MLR_woodC, MLR_woodW As Double
        Dim TfromIgn As Double
        Dim TfromIgn1 As Double = T - flashover_time - CLTceildelamT 's burning time of current lamellae
        Dim TfromIgn2 As Double = T - flashover_time - CLTwalldelamT 's burning time of current lamellae


        Dim Tinit As Double = 20 'C initial temperature
        'Dim hc As Double = 0.02 'kW/m2K
        Dim dchar As Double = 200 'kg/m3 density of char
        'Dim dchar As Double = 300 'kg/m3 density of char

        Dim QCinc As Double = QCeilingAST(fireroom, 0, T / Timestep) 'kW/2 incident radiant heat flux
        Dim QWinc As Double = QUpperWallAST(fireroom, 0, T / Timestep) 'kW/2 incident radiant heat flux
        Dim Qinc As Double

        For counter = 1 To 2
            If counter = 2 And CLTwallpercent = 0 Then
                Exit For
            End If
            If counter = 1 And CLTceilingpercent = 0 Then
                GoTo 2020
            End If

            If counter = 1 Then Qinc = QCinc
            If counter = 2 Then Qinc = QWinc

            'If Qinc < 5 Then 'self-entinguishment below 5 kW/m2
            '    Qinc = 0 '

            '    If counter = 1 Then
            '        MLR_woodC = 0
            '    End If
            '    If counter = 2 Then
            '        MLR_woodW = 0
            '    End If
            '    Continue For
            'End If

            'Douglas Fir along the grain

            If counter = 2 Then
                dw = WallDensity(fireroom)
                kw = WallConductivity(fireroom)
                cw = WallSpecificHeat(fireroom)
            Else
                dw = CeilingDensity(fireroom)
                kw = CeilingConductivity(fireroom)
                cw = CeilingSpecificHeat(fireroom)
            End If

            Dim thetas As Double = ((Qinc + Qff) / (StefanBoltzmann / 1000 * (Tign + 273) ^ 4)) ^ 0.25
            Dim Tsurf As Double = (Tign + 273) * thetas - 273 'C max theoretical surface temp 
            Dim hvap As Double = Lw - cw * (Tign - Tinit) * 10 ^ -6 'kJ/g heat of vaporisation
            Dim alpha As Double = kw / dw / cw 'm2/s thermal diffusivity
            Dim charfraction As Double = 0.74 * (Qinc / qcrit) ^ (-0.64)

            'lets put an upper limit on char fraction ****************************
            If charfraction > 1 Then charfraction = 1

            Dim Qplus = Qff + Qinc - StefanBoltzmann / 1000 * ((Tign + 273) ^ 4 - (Tinit + 273) ^ 4) 'kW/m2 heat flux after ignition
            Dim deltas As Double = 2 * kw * Lw / (cw / 1000 * Qplus)

            If counter = 1 Then TfromIgn = TfromIgn1
            If counter = 2 Then TfromIgn = TfromIgn2

            Dim ND_time As Double = TfromIgn * alpha / deltas ^ 2
            Dim ND_longtime_MLR As Double = Sqrt(cw * 10 ^ -6 * (Tign + 273) / hvap * charfraction * (thetas - 1) / 8 / ND_time)


            'restrict the ND_longtime_MLR to not less than 0.04 **********************************
            'If ND_longtime_MLR < 0.04 Then
            '    ND_longtime_MLR = 0.04
            'End If

            Dim MLR_longtime As Double = ND_longtime_MLR / Lw * (1 - charfraction) * Qplus 'g/s/m2

            'Dim calibration As Double = 1
            MLR_longtime = MLR_longtime * CLTcalibrationfactor

            Dim deltaign As Double = cw * 10 ^ -6 * (Tign - Tinit) / Lw * Qplus / Qinc
            Dim delta As Double = deltaign + 6 * Lw / hvap * (1 - deltaign) * ND_time
            Dim ND_shorttime_MLR As Double = Lw / hvap * (1 - cw * 10 ^ -6 * (Tign - Tinit) / Lw / delta)
            Dim MLR_shorttime As Double = ND_shorttime_MLR / Lw * (1 - charfraction) * Qplus 'g/s/m2
            If charfraction >= 1 Then
                MLR_longtime = 0
                MLR_shorttime = 0
            End If

            'If MLR_longtime < 3.5 Then MLR_longtime = 3.5

            MLR_wood = Min(MLR_longtime, MLR_shorttime) 'g/s/m2
            If MLR_wood < 0 Then MLR_wood = 0

            If counter = 1 Then MLR_woodC = MLR_wood
            If counter = 2 Then MLR_woodW = MLR_wood

2020:
        Next

        'get HRR for all burning objects
        Dim total1, vp, total2, total3 As Double
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

                'kg/s

                If total3 < total Then
                    total = total3 'use the lesser, ventilation control
                    burnmode = True
                End If

                'better to apply excess fuel factor after determining mode of burning and then only if vent-limited
                If burnmode = True Then
                    total = total * ExcessFuelFactor

                    'shouldn't really let this be more than total1 or total2
                    'total = Min(total, total1)
                    'total = Min(total, total2)
                End If

            End If


            'add the surface linings
            mwall = CLTwallpercent / 100 * (RoomLength(fireroom) + RoomWidth(fireroom)) * 2 * RoomHeight(fireroom) * MLR_woodW / 1000 'kg/s
            'add the surface linings
            ' wood_MLR = surfacearea * MLR_wood / 1000 'kg/s  total MLR
            mceiling = CLTceilingpercent / 100 * RoomFloorArea(fireroom) * MLR_woodC / 1000 'kg/s

            'If ceil_char(stepcount - 1, 2) > thickness_ceil Then
            '    'no more fuel
            '    mceiling = 0 'g/s/m2
            'End If

            'If wall_char(stepcount - 1, 2) > thickness_wall Then
            '    'no more fuel
            '    mwall = 0 'g/s/m2
            'End If


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


            'If CLTwallpercent > 0 Then

            '    Dim D As Func(Of Double, Double) = AddressOf MyFunction
            '    Dim F As New OneVariableFunction(D)
            '    'Dim result As Double = H.Integrate(0, stepcount) 'kg consumed so far
            '    'ceil_char(stepcount, 2) = result / ((dw - dchar) * CLTceilingpercent / 100 * RoomFloorArea(fireroom)) 'ceiling depth burned

            '    If wall_char(stepcount, 2) > Lamella2 And Lamella2 = Lamella Then  'delaminate based on 1D heat conduction into surface - 1st layer
            '        'If wall_char(stepcount, 2) > Lamella2 Then  'delaminate based on 1D heat conduction into surface - 1st layer

            '        CLTwalldelamT = T - flashover_time

            '        Dim Message As String = CStr(tim(stepcount, 1)) & " sec. Wall layer at " & Format(Lamella2, "0.000") & " m delaminates."
            '        frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

            '        result2 = F.Integrate(0, stepcount) 'kg consumed in the first layer
            '        'next layer will delmainate when the same kg are 

            '        Lamella2 = Lamella2 + Lamella
            '    ElseIf Lamella2 > thickness_wall + gcd_Machine_Error Then
            '        'no more fuel
            '        mwall = 0
            '    ElseIf result2 > 0 Then 'first layer already fallen off
            '        Dim result As Double = F.Integrate(0, stepcount)  'kg consumed so far

            '            If result >= Lamella2 / Lamella * result2 Then 'next layer fallen off

            '                Dim Message As String = CStr(tim(stepcount, 1)) & " sec. Wall layer at " & Format(Lamella2, "0.000") & " m delaminates."
            '                frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

            '                Lamella2 = Lamella2 + Lamella
            '                CLTwalldelamT = T - flashover_time
            '            End If
            '        End If
            '    End If

            '    If CLTceilingpercent > 0 Then

            '    Dim G As Func(Of Double, Double) = AddressOf MyFunction2
            '    Dim H As New OneVariableFunction(G)
            '    'Dim result As Double = H.Integrate(0, stepcount) 'kg consumed so far
            '    'ceil_char(stepcount, 2) = result / ((dw - dchar) * CLTceilingpercent / 100 * RoomFloorArea(fireroom)) 'ceiling depth burned

            '    If ceil_char(stepcount, 2) > Lamella1 And Lamella1 = Lamella Then  'delaminate based on 1D heat conduction into surface - 1st layer

            '        CLTceildelamT = T - flashover_time

            '        Dim Message As String = CStr(tim(stepcount, 1)) & " sec. Ceiling layer at " & Format(Lamella1, "0.000") & " m delaminates."
            '        frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

            '        result1 = H.Integrate(0, stepcount) 'kg consumed in the first layer
            '        'next layer will delmainate when the same kg are 

            '        Lamella1 = Lamella1 + Lamella
            '    ElseIf Lamella1 > thickness_ceil + gcd_Machine_Error Then
            '        'no more fuel
            '        mceiling = 0
            '    ElseIf result1 > 0 Then 'first layer already fallen off
            '        Dim result As Double = H.Integrate(0, stepcount)  'kg consumed so far

            '        If result >= Lamella1 / Lamella * result1 Then 'next layer fallen off

            '            Dim Message As String = CStr(tim(stepcount, 1)) & " sec. Ceiling layer at " & Format(Lamella1, "0.000") & " m delaminates."
            '            frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

            '            Lamella1 = Lamella1 + Lamella
            '            CLTceildelamT = T - flashover_time

            '        End If
            '    End If
            'End If

            ''test for delamination
            'Dim NumberceilingNodes As Integer = UBound(CeilingNode, 2)
            'Dim depth, tempatnode As Double
            'Dim crittemp As Double = 300 + 273
            'Dim X(0 To NumberceilingNodes) As Double 'node temps
            'Dim Y(0 To NumberceilingNodes) As Double 'node depths
            'For j = 0 To stepcount
            '    For curve = 1 To NumberceilingNodes
            '        'what is my depth?
            '        If HaveCeilingSubstrate(fireroom) = True Then
            '            depth = (curve - 1) * CeilingThickness(fireroom) / (Ceilingnodes - 1)
            '            If depth > CeilingThickness(fireroom) Then
            '                depth = CeilingThickness(fireroom) + (curve - Ceilingnodes) * CeilingSubThickness(fireroom) / (Ceilingnodes - 1)
            '            End If
            '        Else
            '            depth = (curve - 1) * CeilingThickness(fireroom) / (NumberceilingNodes - 1)
            '        End If
            '        If depth >= Lamella1 Then
            '            tempatnode = CeilingNode(fireroom, curve, j) 'temp at this node
            '            If tempatnode >= crittemp Then
            '                'delamination
            '                Dim Message As String = CStr(tim(stepcount, 1)) & " sec. Ceiling layer at " & Format(Lamella1, "0") & " m delaminates."
            '                frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
            '                Lamella1 = Lamella1 + Lamella
            '            End If
            '        End If
            '    Next
            'Next


        End If

        wall_char(stepcount, 1) = mwall 'kg/s
        ceil_char(stepcount, 1) = mceiling

        wood_MLR = mceiling + mwall 'kg/s  total MLR

        mwallsave = mwall
        mceilingsave = mceiling

        MassLoss_Total_pluswood = total + wood_MLR 'includes contents+surfaces
        MLRsave = MassLoss_Total_pluswood

        Tsave = T


    End Function
    Function MassLoss_CLTinstant(ByVal T As Double, ByRef mwall As Double, ByRef mceiling As Double) As Double
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
            MassLoss_CLTinstant = MLRsave
            Exit Function
        End If

        Dim QW, Qtemp As Double
        Dim QCeiling As Double = 0
        Dim QWall As Double = 0
        Dim QFloor As Double = 0
        Dim totalarea, total As Double
        Dim mass, wood_MLR As Double

        QW = 0 'baggage
        totalarea = 0
        burnmode = False

        mass = InitialFuelMass 'kg wood cribs

        Dim thickness_ceil As Double = CeilingThickness(fireroom) / 1000 'm thick of wood
        Dim thickness_wall As Double = WallThickness(fireroom) / 1000  'm thick of wood
        Dim MLR_wood, MLR_woodC, MLR_woodW As Double


        For counter = 1 To 2
            If counter = 2 And CLTwallpercent = 0 Then
                Exit For
            End If
            If counter = 1 And CLTceilingpercent = 0 Then
                GoTo 2020
            End If



            If counter = 1 Then MLR_woodC = MLR_wood
            If counter = 2 Then MLR_woodW = MLR_wood

2020:
        Next

        Dim woodtotal As Double = 0


        'add the surface linings
        mwall = CLTwallpercent / 100 * (RoomLength(fireroom) + RoomWidth(fireroom)) * 2 * RoomHeight(fireroom) * MLR_woodW / 1000 'kg/s
        'add the surface linings
        mceiling = CLTceilingpercent / 100 * RoomFloorArea(fireroom) * MLR_woodC / 1000 'kg/s


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

        MassLoss_CLTinstant = total + wood_MLR 'includes contents+surfaces
        MLRsave = MassLoss_CLTinstant

        Tsave = T

    End Function

    Private Function MyFunction(ByVal X As Double) As Double
        'Return Math.Sqrt(X)
        Return wall_char(X, 1)
    End Function
    Private Function MyFunction2(ByVal X As Double) As Double
        'Return Math.Sqrt(X)
        Return ceil_char(X, 1)
    End Function

    Sub Max_Vent()
        '*  ==============================================================
        '*  Find the vents with the highest soffit height above floor,
        '*  and the vent with the lowest sill.
        '*
        '*  Revised 17 April 1996 Colleen Wade
        '*  ==============================================================

        'Dim SoffitMax As Single, j As Integer, SillMin  As Single, i As Integer, k As Integer
        Dim SoffitMax, SillMin As Double
        Dim i, j, k As Short '28/9/2005

        SoffitMax = 0

        For i = 1 To NumberRooms
            For k = 2 To NumberRooms + 1
                If NumberVents(i, k) > 0 Then
                    For j = 1 To NumberVents(i, k)
                        If soffitheight(i, k, j) > SoffitMax Then
                            SoffitMax = soffitheight(i, k, j)
                            Highest_Vent(i) = j
                        End If
                    Next j
                End If
                SillMin = RoomHeight(i)

                If NumberVents(i, k) > 0 Then
                    For j = 1 To NumberVents(i, k)
                        If VentSillHeight(i, k, j) < SillMin Then
                            SillMin = VentSillHeight(i, k, j)
                            Lowest_Vent(i) = j
                        End If
                    Next j
                End If
            Next k
        Next i
    End Sub


    Sub New_File()
        '*  ================================================================
        '*      Set up initial defaults for a new model
        '*  ================================================================
        ReDim NumberVents(MaxNumberRooms + 1, MaxNumberRooms + 1)
        ReDim NumberCVents(MaxNumberRooms + 1, MaxNumberRooms + 1)
        Dim i, j As Short

        If (DataDirectory) = "" Then DataDirectory = UserPersonalDataFolder & gcs_folder_ext & "\data\"
        If (RiskDataDirectory) = "" Then RiskDataDirectory = DefaultRiskDataDirectory & "basemodel_default\"

        'If Confirm_File(FireDatabaseName, readfile, 0) = 0 Then FireDatabaseName = UserPersonalDataFolder & gcs_folder_ext & DefaultFireDatabaseName
        'If Confirm_File(MaterialsDatabaseName, readfile, 0) = 0 Then MaterialsDatabaseName = UserPersonalDataFolder & gcs_folder_ext & DefaultMaterialsDatabaseName

        If My.Computer.FileSystem.FileExists(FireDatabaseName) = False Then
            'get a valid copy
            My.Computer.FileSystem.CopyFile(ApplicationPath & DefaultFireDatabaseName, UserPersonalDataFolder & gcs_folder_ext & DefaultFireDatabaseName, True)
        End If
        If My.Computer.FileSystem.FileExists(MaterialsDatabaseName) = False Then
            My.Computer.FileSystem.CopyFile(ApplicationPath & DefaultMaterialsDatabaseName, UserPersonalDataFolder & gcs_folder_ext & DefaultMaterialsDatabaseName, True)
        End If


        TalkToEVACNZ = False
        NumberRooms = 1
        graphroom = 0
        Error_Control = 0.1
        Error_Control_ventflow = 0.001
        Wallnodes = 15
        Ceilingnodes = 15
        Floornodes = 10
        gTimeX = 0
        autopopulate = False
        'VM2 = False
        'AlphaT = 0.0469
        'PeakHRR = 20000
        'StoreHeight = 1
        'useT2fire = True
        'usepowerlawdesignfire = False

        'Vent Description
        For i = 1 To NumberRooms + 1
            For j = 1 To NumberRooms + 1
                NumberVents(i, j) = 0
                NumberVents(j, i) = 0
                NumberCVents(i, j) = 0
                'NumberCVents(j, i) = 0
            Next j
        Next i

        Resize_Vents() 'size the vent arrays
        Resize_CVents()

        For i = 1 To NumberRooms + 1
            For j = 1 To NumberRooms + 1
                FRintegrity(i, j, 1) = 0
                FRMaxOpening(i, j, 1) = 100
                FRMaxOpeningTime(i, j, 1) = 10
                VentWidth(i, j, 1) = 0
                VentCD(i, j, 1) = 0
                WallLength1(i, j, 1) = 0
                WallLength2(i, j, 1) = 0
                VentHeight(i, j, 1) = 0
                spillplume(i, j, 1) = 0
                spillplumemodel(i, j, 1) = 0
                spillbalconyprojection(i, j, 1) = True
                VentSillHeight(i, j, 1) = 0
                VentOpenTime(i, j, 1) = 0
                VentCloseTime(i, j, 1) = 0
                Downstand(i, j, 1) = 0
                GLASSconductivity(i, j, 1) = 0.76
                GLASSemissivity(i, j, 1) = 1.0#
                GLASSexpansion(i, j, 1) = 0.0000095
                GLASSthickness(i, j, 1) = 4
                GLASSdistance(i, j, 1) = 0
                GLASSFalloutTime(i, j, 1) = 0
                AutoBreakGlass(i, j, 1) = False

                trigger_ventopenduration(i, j, 1) = 0 '2009.2
                AutoOpenVent(i, j, 1) = False '2009.2
                trigger_device(0, i, j, 1) = False '2009.2
                trigger_device(1, i, j, 1) = False '2009.2
                trigger_device(2, i, j, 1) = False '2009.2
                trigger_device(3, i, j, 1) = False
                trigger_device(4, i, j, 1) = False
                trigger_device(5, i, j, 1) = False
                trigger_device(6, i, j, 1) = False
                SDtriggerroom(i, j, 1) = 1 '2009.2
                HRR_ventopenduration(i, j, 1) = 0 '2009.2
                HRR_ventopendelay(i, j, 1) = 0 '2009.2
                HRR_threshold(i, j, 1) = 500 '2009.2
                HRR_threshold_time(i, j, 1) = 0 '2009.2
                HRRFlag(i, j, 1) = 0 '2009.2
                trigger_ventopendelay(i, j, 1) = 0 '2009.2

                'breakflag(CShort(i), CShort(j), 1) = False
                breakflag(i, j, 1) = False
                SpillPlumeBalc(i, j, 1) = True
                GlassFlameFlux(i, j, 1) = False
                GLASSshading(i, j, 1) = 15
                GLASSbreakingstress(i, j, 1) = 47
                GLASSalpha(i, j, 1) = 0.00000036
                GlassYoungsModulus(i, j, 1) = 72000
            Next j
        Next i

        'Program Options
        MonitorHeight = 2
        Activity = "Light"
        InteriorTemp = 293
        ExteriorTemp = 288
        RelativeHumidity = 0.65
        SimTime = 600
        DisplayInterval = 10
        ExcelInterval = 10
        Description = ""
        JobNumber = ""
        StartOccupied = 0
        EndOccupied = 10000
        Timestep = 1

        'Fire Parameters
        'RadiantLossFraction = 0.3
        SootFactor = 1
        'MLUnitArea = 0.011
        EmissionCoefficient = 0.8
        preCO = 0.04
        postCO = 0.4
        preSoot = 0.07
        postSoot = 0.14
        sootmode = False
        comode = False
        SootAlpha = 2.5
        SootEpsilon = 1.2

        Fuel_Thickness = 0.05 'm
        Stick_Spacing = 0.05 'm
        Cribheight = 0.8 'm
        ExcessFuelFactor = 1.3 'm
        HoC_fuel = 13 'MJ/kg
        'FuelHeatofGasification = 3
        FuelSurfaceArea = 0
        nC = 1
        nH = 2
        nO = 0.5
        nN = 0
        StoichAFratio = 6.1
        FLED = 400 'MJ/m2
        Enhance = False
        fueltype = "wood"
        'frmOptions1.txtFLED.Text = CStr(FLED)

        frmOptions1.txtFuelThickness.Text = CStr(Fuel_Thickness)
        frmOptions1.txtCribheight.Text = CStr(Cribheight)
        frmOptions1.txtStickSpacing.Text = CStr(Stick_Spacing)
        frmOptions1.txtExcessFuelFactor.Text = CStr(ExcessFuelFactor)
        'frmOptions1.txtHOCFuel.Text = CStr(HoC_fuel)
        frmOptions1.txtnC.Text = CStr(nC)
        frmOptions1.txtnH.Text = CStr(nH)
        frmOptions1.txtnO.Text = CStr(nO)
        frmOptions1.txtnN.Text = CStr(nN)
        frmOptions1.txtStoich.Text = CStr(StoichAFratio)
        frmOptions1.chkHCNcalc.CheckState = System.Windows.Forms.CheckState.Unchecked 'disabled by default
        nowallflow = True
        frmOptions1.chkWallFlowDisable.CheckState = System.Windows.Forms.CheckState.Checked 'disabled by default
        NumberObjects = 0
        fireroom = 1

        FEDPath(0, 0) = 0
        FEDPath(1, 0) = SimTime
        FEDPath(2, 0) = fireroom

        FEDPath(0, 1) = SimTime
        FEDPath(1, 1) = SimTime
        FEDPath(2, 1) = fireroom

        FEDPath(0, 2) = SimTime
        FEDPath(1, 2) = SimTime
        FEDPath(2, 2) = fireroom


        ReDim FireHeight(NumberObjects + 1)
        ReDim FireLocation(NumberObjects + 1)
        ReDim EnergyYield(NumberObjects + 1)
        'ReDim COYield(NumberObjects)
        ReDim CO2Yield(NumberObjects + 1)
        ReDim SootYield(NumberObjects + 1)
        ReDim WaterVaporYield(NumberObjects + 1)
        ReDim HCNYield(NumberObjects + 1)
        ReDim HCNuserYield(NumberObjects + 1)
        ReDim ObjCRF(NumberObjects + 1)
        ReDim ObjCRFauto(NumberObjects + 1)
        ReDim ObjFTPindexpilot(NumberObjects + 1)
        ReDim ObjFTPindexauto(NumberObjects + 1)
        ReDim ObjFTPlimitpilot(NumberObjects + 1)
        ReDim ObjFTPlimitauto(NumberObjects + 1)
        ReDim ObjLabel(NumberObjects + 1)
        ReDim ObjLength(NumberObjects + 1)
        ReDim ObjWidth(NumberObjects + 1)
        ReDim ObjHeight(NumberObjects + 1)
        ReDim ObjDimX(NumberObjects + 1)
        ReDim ObjDimY(NumberObjects + 1)
        ReDim ObjElevation(NumberObjects + 1)
        ReDim ObjectDescription(NumberObjects + 1)
        ReDim ObjectMass(NumberObjects + 1)
        ReDim ObjectItemID(NumberObjects + 1)
        ReDim ObjectIgnTime(NumberObjects + 1)
        ReDim ObjectMLUA(0 To 2, NumberObjects + 1)
        ReDim ObjectLHoG(NumberObjects + 1)
        ReDim ObjectIgnMode(NumberObjects + 1)
        ReDim ObjectRLF(NumberObjects + 1)
        ReDim ObjectWindEffect(NumberObjects + 1)
        ReDim ObjectPyrolysisOption(NumberObjects + 1)
        ReDim ObjectPoolDensity(NumberObjects + 1)
        ReDim ObjectPoolDiameter(NumberObjects + 1)
        ReDim ObjectPoolFBMLR(NumberObjects + 1)
        ReDim ObjectPoolVapTemp(NumberObjects + 1)
        ReDim ObjectPoolRamp(NumberObjects + 1)
        ReDim ObjectPoolVolume(NumberObjects + 1)
        ReDim HeatReleaseData(2, 500, NumberObjects + 1)
        ReDim MLRData(2, 500, NumberObjects + 1)
        ReDim NumberDataPoints(NumberObjects + 1)
        ReDim WallSurface(NumberRooms)
        ReDim WallSubstrate(NumberRooms)
        ReDim CeilingSurface(NumberRooms)
        ReDim CeilingSubstrate(NumberRooms)
        ReDim FloorSubstrate(NumberRooms)
        ReDim FloorSurface(NumberRooms)
        ReDim HaveCeilingSubstrate(NumberRooms)
        ReDim HaveWallSubstrate(NumberRooms)
        ReDim HaveFloorSubstrate(NumberRooms)
        ReDim HaveSD(NumberRooms)
        ReDim SDinside(NumberRooms)
        ReDim SpecifyOD(NumberRooms)
        ReDim TwoZones(0 To NumberRooms)
        ReDim FanAutoStart(NumberRooms)

        SDinside(1) = True
        EnergyYield(1) = 16
        CO2Yield(1) = 1.2
        SootYield(1) = 0.03
        WaterVaporYield(1) = 1.0#
        HCNuserYield(1) = 0
        ObjCRF(1) = 9.5
        ObjFTPindexpilot(1) = 1.0
        ObjFTPlimitpilot(1) = 481
        ObjCRFauto(1) = 22
        ObjFTPindexauto(1) = 1.0
        ObjFTPlimitauto(1) = 432
        ObjLength(1) = 1
        ObjWidth(1) = 1
        ObjHeight(1) = 1
        ObjDimX(1) = RoomLength(1) / 2 - 0.5
        ObjDimY(1) = RoomWidth(1) / 2 - 0.5
        ObjElevation(1) = 0
        ObjectDescription(1) = ""
        ObjectItemID(1) = 1
        ObjectIgnTime(1) = 0
        ObjectMLUA(0, 1) = 0 'MLUA constant A
        ObjectMLUA(1, 1) = 0 'MLUA constant B
        ObjectLHoG(1) = 3
        ObjectIgnMode(1) = ""
        ObjectRLF(1) = 0.3
        ObjectWindEffect(1) = 1
        ObjectPyrolysisOption(1) = 0
        ObjectPoolDensity(1) = 684
        ObjectPoolDiameter(1) = 0
        ObjectPoolFBMLR(1) = 0.101
        ObjectPoolRamp(1) = 0
        ObjectPoolVolume(1) = 0
        'COYield(1) = 0.002
        FireHeight(1) = 0
        FireLocation(1) = 0
        NumberDataPoints(1) = 2
        HeatReleaseData(1, 1, 1) = 0
        HeatReleaseData(2, 1, 1) = 0
        MLRData(1, 1, 1) = 0
        MLRData(2, 1, 1) = 0

        'Room Description Defaults
        For i = 1 To NumberRooms
            RoomWidth(i) = 5
            RoomLength(i) = 5
            RoomDescription(i) = ""
            RoomHeight(i) = 3
            FloorElevation(i) = 0
            RoomAbsX(i) = 0
            RoomAbsY(i) = 0
            CeilingSlope(i) = False
            TwoZones(i) = True
            MinStudHeight(i) = 3
            'WallSurface(i) = "concrete"
            WallSurface(i) = "plasterboard"
            WallSubstrate(i) = "None"
            'CeilingSurface(i) = "concrete"
            CeilingSurface(i) = "plasterboard"
            CeilingSubstrate(i) = "None"
            FloorSubstrate(i) = "None"
            FloorSurface(i) = "concrete"
            'frmDescribeRoom.lblWallSurface.Text = WallSurface(i)
            'frmDescribeRoom.lblWallSubstrate.Text = WallSubstrate(i)
            'frmDescribeRoom.lblCeilingSurface.Text = CeilingSurface(i)
            'frmDescribeRoom.lblCeilingSubstrate.Text = CeilingSubstrate(i)
            'frmDescribeRoom.lblFloorSurface.Text = FloorSurface(i)
            ThermalInertiaWall(i) = 0.515
            ThermalInertiaCeiling(i) = 0.515
            ThermalInertiaFloor(i) = 0.515
            FloorConeDataFile(i) = "null.txt"
            WallConeDataFile(i) = "null.txt"
            CeilingConeDataFile(i) = "null.txt"
            WallSootYield(i) = 0
            WallH2OYield(i) = 0
            WallHCNYield(i) = 0
            WallCO2Yield(i) = 0
            CeilingSootYield(i) = 0
            FloorSootYield(i) = 0
            CeilingH2OYield(i) = 0
            CeilingHCNYield(i) = 0
            FloorH2OYield(i) = 0
            FloorHCNYield(i) = 0
            CeilingCO2Yield(i) = 0
            FloorCO2Yield(i) = 0
            'WallThickness(i) = 100
            WallThickness(i) = 13
            WallSubThickness(i) = 0
            'CeilingThickness(i) = 100
            CeilingThickness(i) = 13
            CeilingSubThickness(i) = 0
            FloorSubThickness(i) = 0
            FloorThickness(i) = 100
            'CeilingConductivity(i) = 1.2
            CeilingConductivity(i) = 0.16
            CeilingSubConductivity(i) = 1
            'WallConductivity(i) = 1.2
            WallConductivity(i) = 0.16
            WallSubConductivity(i) = 1
            FloorSubConductivity(i) = 1
            FloorConductivity(i) = 1.2
            'WallSpecificHeat(i) = 880
            WallSpecificHeat(i) = 900
            'CeilingSpecificHeat(i) = 880
            CeilingSpecificHeat(i) = 900
            WallSubSpecificHeat(i) = 1000
            FloorSubSpecificHeat(i) = 1000
            CeilingSubSpecificHeat(i) = 1000
            FloorSpecificHeat(i) = 880
            'WallDensity(i) = 2300
            WallDensity(i) = 810
            'CeilingDensity(i) = 2300
            CeilingDensity(i) = 810
            WallSubDensity(i) = 2000
            FloorSubDensity(i) = 2000
            CeilingSubDensity(i) = 2000
            FloorDensity(i) = 2300
            'mechanical extraction
            ExtractRate(i) = 0
            NumberFans(i) = 1
            ExtractStartTime(i) = 0
            FanElevation(i) = RoomHeight(i)
            MaxPressure(i) = 50
            Extract(i) = True
            UseFanCurve(i) = True

            'call detector defaults
            Device_Defaults((i))

            'RTI(i) = 140
            'RadialDistance(i) = 4.3
            'SmokeOD(i) = 0.14
            'SDdelay(i) = 15
            'DetSensitivity(i) = 2.5
            'SDRadialDist(i) = 0
            'SDdepth(i) = 0.025
            'cfactor(i) = 1
            'ActuationTemp(i) = 68 + 273
            'WaterSprayDensity(i) = 0
            'SprinkDistance(i) = 0.025
            'DetectorType(i) = 0

        Next i
        For i = 1 To 4
            Surface_Emissivity(i, 1) = 0.5
        Next i

        'broke
        'frmSprink.cboDetectorType.SelectedIndex = 0

        'room corner
        WallHeatFlux = 45
        CeilingHeatFlux = 35
        FlameAreaConstant = 0.0065
        corner = 0
        FlameLengthPower = 1
        BurnerWidth = 0.17
        IgTempW(fireroom) = 742
        IgTempC(fireroom) = 742
        IgTempF(fireroom) = 742
        RCNone = True

        'endpoints
        VisibilityEndPoint = 10
        ConvectEndPoint = 80 + 273
        TargetEndPoint = 0.3
        TempEndPoint = 200 + 273
        FEDEndPoint = 0.3

        'Other Defaults
        NumberTimeSteps = 0
        ' NumberObjects = 0
        Resize_Objects()
        NumberDataPoints(1) = 0

        'broke
        'frmExtract.lstRoomID.SelectedIndex = 0
        'FrmSprink.lstRoomDetector.SelectedIndex = 0
        'frmSmokeDetector.lstRoomSD.SelectedIndex = 0
        'Dim location As New ComboBox
        'location = frmFire.lstObjectLocation
        'If location.Items.Count > 0 Then
        '    location.SelectedIndex = 0
        'Else
        '    location.SelectedIndex = -1
        'End If

        frmOptions1.optMcCaffrey.Checked = True

        frmOptions1.txtErrorControl.Text = Format(Error_Control, "0.000")
        frmOptions1.txtErrorControlVentFlow.Text = Format(Error_Control_ventflow, "0.000000")
        frmOptions1.optTwoLayer.Checked = True
        frmInputs.chkSaveVentFlow.CheckState = CheckState.Unchecked

        IgnCorrelation = vbFTP
        'disable menu item run if no heat release data entered
        'For i = 0 To NumberObjects
        '    If NumberDataPoints(i) <> 0 And NumberObjects > 0 Then
        '        'MDIFrmMain.mnuRun.Enabled = True
        '        MDIFrmMain.mnuRun.Visible = True
        '        Exit For
        '    Else
        '        'MDIFrmMain.mnuRun.Enabled = False
        '        MDIFrmMain.mnuRun.Visible = False
        '    End If
        'Next i

        'cjModel = cjAlpert
        MDIFrmMain.mnuExcel.Enabled = False
        MDIFrmMain.ToolStripStatusLabel1.Text = "NEW MODEL"
        MDIFrmMain.ToolStripStatusLabel2.Text = Description
        frmOptions1.optReflectiveSign.Checked = True

        Dim matname, firename As String
        matname = Path.GetFileName(MaterialsDatabaseName)
        firename = Path.GetFileName(FireDatabaseName)

        nC = 0.95
        nH = 2.4
        nO = 1
        nN = 0
        StoichAFratio = 6.1
        frmOptions1.txtnC.Text = CStr(nC)
        frmOptions1.txtnH.Text = CStr(nH)
        frmOptions1.txtnO.Text = CStr(nO)
        frmOptions1.txtnN.Text = CStr(nN)
        frmOptions1.txtStoich.Text = CStr(StoichAFratio)
        frmOptions1.cboABSCoeff.Text = fueltype
        ' frmOptions1.txtRadiantLossFraction.Text = CStr(RadiantLossFraction)
        ' frmOptions1.txtMassLossUnitArea.Text = CStr(MLUnitArea)
        frmOptions1.txtEmissionCoefficient.Text = CStr(EmissionCoefficient)
        frmOptions1.txtSootEps.Text = CStr(SootEpsilon)
        frmOptions1.optCOman.Checked = comode
        If comode = False Then frmOptions1.optCOauto.Checked = True
        frmOptions1.optSootman.Checked = sootmode
        If sootmode = False Then frmOptions1.optSootauto.Checked = True
        frmOptions1.txtpreCO.Text = CStr(preCO)
        frmOptions1.txtpostCO.Text = CStr(postCO)
        frmOptions1.txtpreSoot.Text = CStr(preSoot)
        frmOptions1.txtPostSoot.Text = CStr(postSoot)
        MDIFrmMain.NZBCVM2ToolStripMenuItem.Checked = UNCHECKED
        MDIFrmMain.RiskSimulatorToolStripMenuItem.Checked = CHECKED

        ventlog = False
        frmInputs.chkSaveVentFlow.Checked = False

        autosavepdf = False
        frmInputs.ChkAutosavePdf.Checked = False

        autosaveXL = False
        frmInputs.chkAutosaveXL.Checked = False

        useCLTmodel = False
        frmCLT.optCLTOFF.Checked = True

        'kinetic propeties for each component
        'Activation Energy
        E_array(1) = 198000.0 'J/mol cellulose
        E_array(2) = 164000.0 'J/mol hemicellulose
        E_array(3) = 152000.0 'J/mol lignin
        E_array(0) = 100000.0 'J/mol

        'Pre-exponential factor 
        A_array(1) = 351000000000000.0 '1/s
        A_array(2) = 32500000000000.0 '1/s
        A_array(3) = 84100000000000.0 '1/s
        A_array(0) = 10000000000000.0 '1/s

        'Reaction order
        n_array(1) = 1.1
        n_array(2) = 2.1
        n_array(3) = 5
        n_array(0) = 1

        'initial component mass fraction
        mf_compinit(0) = 0.1
        mf_compinit(1) = 0.44
        mf_compinit(2) = 0.37
        mf_compinit(3) = 0.09

        'char yields
        char_yield(1) = 0.13
        char_yield(2) = 0.13
        char_yield(3) = 0.13

        Call MDIFrmMain.VM2setup()

    End Sub

    Sub New_File_Start()
        '*  ================================================================
        '*      Set up initial defaults for a new model
        '*      for the first time program is started
        '*  ================================================================

        Dim i, j As Short

        ReDim NumberVents(MaxNumberRooms + 1, MaxNumberRooms + 1)
        ReDim NumberCVents(MaxNumberRooms + 1, MaxNumberRooms + 1)
        ReDim FEDPath(0 To 2, 0 To 2)

        'frmBlank.Show()

        'If (DataDirectory) = "" Then DataDirectory = UserAppDataFolder & gcs_folder_ext & "\data\"
        If (FireDatabaseName) = "" Then FireDatabaseName = UserPersonalDataFolder & gcs_folder_ext & DefaultFireDatabaseName
        If (MaterialsDatabaseName) = "" Then MaterialsDatabaseName = UserPersonalDataFolder & gcs_folder_ext & DefaultMaterialsDatabaseName

        'Vent Description
        NumberRooms = 1
        ReDim Sprinkler(3, NumberRooms)
        fireroom = 1
        graphroom = 0
        Timestep = 1
        Wallnodes = 15
        Ceilingnodes = 15
        Floornodes = 10
        Error_Control = 0.1
        Error_Control_ventflow = 0.001
        LEsolver = "LU Decomposition"
        PessimiseCombWall = True
        frmOptions1.optLUdecom.Checked = True
        frmOptions1.txtErrorControl.Text = Format(Error_Control, "0.000")
        frmOptions1.txtErrorControlVentFlow.Text = Format(Error_Control_ventflow, "0.000000")
        frmOptions1.optTwoLayer.Checked = True
        VM2 = False
        AlphaT = 0.0469
        PeakHRR = 20000
        StoreHeight = 1
        useT2fire = True
        usepowerlawdesignfire = False
        useCLTmodel = False

        For i = 1 To NumberRooms + 1
            For j = 1 To NumberRooms + 1
                NumberVents(i, j) = 0
                NumberVents(j, i) = 0
                NumberCVents(i, j) = 0
                'NumberCVents(j, i) = 0
            Next j
        Next i

        Resize_Rooms()
        Resize_Vents() 'size the vent arrays
        Resize_CVents() 'size the vent arrays

        'CVentOpenTime(1, 1) = 0
        'CVentDC(1, 1) = 0.6
        'CVentArea(1, 1) = 0

        For i = 1 To NumberRooms + 1
            For j = 1 To NumberRooms + 1
                FRintegrity(i, j, 1) = 0
                FRMaxOpening(i, j, 1) = 100
                FRMaxOpeningTime(i, j, 1) = 10
                VentWidth(i, j, 1) = 0
                VentCD(i, j, 1) = 0
                WallLength1(i, j, 1) = 0
                WallLength2(i, j, 1) = 0
                VentHeight(i, j, 1) = 0
                spillplume(i, j, 1) = 0
                spillplumemodel(i, j, 1) = 0
                spillbalconyprojection(i, j, 1) = True
                VentSillHeight(i, j, 1) = 0
                VentOpenTime(i, j, 1) = 0
                VentCloseTime(i, j, 1) = 0
                Downstand(i, j, 1) = 0.76
                GLASSconductivity(i, j, 1) = 0.76
                GLASSemissivity(i, j, 1) = 1
                GLASSexpansion(i, j, 1) = 0.0000095
                GLASSthickness(i, j, 1) = 4
                GLASSdistance(i, j, 1) = 0
                GLASSFalloutTime(i, j, 1) = 0
                AutoBreakGlass(i, j, 1) = False

                AutoOpenVent(i, j, 1) = False '2009.2
                trigger_device(0, i, j, 1) = False '2009.2
                trigger_device(1, i, j, 1) = False '2009.2
                trigger_device(2, i, j, 1) = False '2009.2
                trigger_device(3, i, j, 1) = False
                trigger_device(4, i, j, 1) = False
                trigger_device(5, i, j, 1) = False
                trigger_device(6, i, j, 1) = False
                SDtriggerroom(i, j, 1) = 1 '2009.2
                HRR_ventopenduration(i, j, 1) = 0 '2009.2
                HRR_ventopendelay(i, j, 1) = 0 '2009.2
                trigger_ventopenduration(i, j, 1) = 0 '2009.2
                HRR_threshold(i, j, 1) = 500 '2009.2
                trigger_ventopendelay(i, j, 1) = 0 '2009.2
                HRR_threshold_time(i, j, 1) = 0 '2009.2
                HRRFlag(i, j, 1) = 0 '2009.2

                'breakflag(CInt(i), CInt(j), 1) = False
                breakflag(i, j, 1) = False
                SpillPlumeBalc(i, j, 1) = True
                GlassFlameFlux(i, j, 1) = False
                GLASSshading(i, j, 1) = 15
                GLASSbreakingstress(i, j, 1) = 47
                GLASSalpha(i, j, 1) = 0.00000036
                GlassYoungsModulus(i, j, 1) = 72000
            Next j
        Next i

        'Program Options
        MonitorHeight = 2
        Activity = "Light"
        InteriorTemp = 293
        ExteriorTemp = 288
        RelativeHumidity = 0.65
        SimTime = 600
        DisplayInterval = 10
        ExcelInterval = 10
        Description = ""
        JobNumber = ""
        StartOccupied = 0
        EndOccupied = 10000

        'Fire Parameters
        'RadiantLossFraction = 0.3
        SootFactor = 1
        ' MLUnitArea = 0.011
        EmissionCoefficient = 0.8
        preCO = 0.04
        postCO = 0.4
        preSoot = 0.07
        postSoot = 0.14
        sootmode = False
        comode = False
        SootAlpha = 2.5
        SootEpsilon = 1.2

        Fuel_Thickness = 0.05
        Stick_Spacing = 0.05
        Cribheight = 0.8 'm
        ExcessFuelFactor = 1.3
        HoC_fuel = 19
        'FuelHeatofGasification = 3
        FuelSurfaceArea = 0
        nC = 1
        nH = 2
        nO = 0.5
        nN = 0
        Enhance = False
        FLED = 400
        fueltype = "wood"
        StoichAFratio = 6.1
        NumberObjects = 0
        ReDim FireHeight(NumberObjects + 1)
        ReDim FireLocation(NumberObjects + 1)
        ReDim EnergyYield(NumberObjects + 1)
        'ReDim COYield(NumberObjects)
        ReDim CO2Yield(NumberObjects + 1)
        ReDim SootYield(NumberObjects + 1)
        ReDim WaterVaporYield(NumberObjects + 1)
        ReDim HCNYield(NumberObjects + 1)
        ReDim HCNuserYield(NumberObjects + 1)
        ReDim ObjCRF(NumberObjects + 1)
        ReDim ObjFTPindexpilot(NumberObjects + 1)
        ReDim ObjFTPlimitpilot(NumberObjects + 1)
        ReDim ObjCRFauto(NumberObjects + 1)
        ReDim ObjFTPindexauto(NumberObjects + 1)
        ReDim ObjFTPlimitauto(NumberObjects + 1)
        ReDim ObjLabel(NumberObjects + 1)
        ReDim ObjLength(NumberObjects + 1)
        ReDim ObjWidth(NumberObjects + 1)
        ReDim ObjHeight(NumberObjects + 1)
        ReDim ObjDimX(NumberObjects + 1)
        ReDim ObjDimY(NumberObjects + 1)
        ReDim ObjElevation(NumberObjects + 1)
        ReDim ObjectDescription(NumberObjects + 1)
        ReDim ObjectMass(NumberObjects + 1)
        ReDim ObjectItemID(NumberObjects + 1)
        ReDim ObjectIgnTime(NumberObjects + 1)
        ReDim ObjectMLUA(0 To 2, NumberObjects + 1)
        ReDim ObjectIgnMode(NumberObjects + 1)
        ReDim ObjectRLF(NumberObjects + 1)
        ReDim ObjectWindEffect(NumberObjects + 1)
        ReDim ObjectPyrolysisOption(NumberObjects + 1)
        ReDim ObjectPoolDensity(NumberObjects + 1)
        ReDim ObjectPoolDiameter(NumberObjects + 1)
        ReDim ObjectPoolFBMLR(NumberObjects + 1)
        ReDim ObjectPoolRamp(NumberObjects + 1)
        ReDim ObjectPoolVolume(NumberObjects + 1)
        ReDim ObjectPoolVapTemp(NumberObjects + 1)
        ReDim ObjectLHoG(NumberObjects + 1)
        ReDim HeatReleaseData(2, 500, NumberObjects + 1)
        ReDim MLRData(2, 500, NumberObjects + 1)
        ReDim NumberDataPoints(NumberObjects + 1)
        ReDim WallSurface(NumberRooms)
        ReDim WallSubstrate(NumberRooms)
        ReDim CeilingSurface(NumberRooms)
        ReDim CeilingSubstrate(NumberRooms)
        ReDim FloorSubstrate(NumberRooms)
        ReDim FloorSurface(NumberRooms)
        ReDim HaveWallSubstrate(NumberRooms)
        ReDim HaveFloorSubstrate(NumberRooms)
        ReDim HaveSD(NumberRooms)
        ReDim SDinside(NumberRooms)
        ReDim SpecifyOD(NumberRooms)
        ReDim HaveCeilingSubstrate(NumberRooms)

        SDinside(1) = True
        EnergyYield(1) = 16
        CO2Yield(1) = 1.2
        SootYield(1) = 0.03
        WaterVaporYield(1) = 1.0#
        HCNuserYield(1) = 0

        ObjHeight(1) = 1
        ObjDimX(1) = RoomLength(1) / 2 - 0.5
        ObjWidth(1) = RoomWidth(1) / 2 - 0.5
        ObjLength(1) = 1
        ObjWidth(1) = 1
        ObjElevation(1) = 0
        ObjCRF(1) = 9.5
        ObjFTPindexpilot(1) = 1
        ObjFTPlimitpilot(1) = 481
        ObjCRFauto(1) = 22
        ObjFTPindexauto(1) = 1
        ObjFTPlimitauto(1) = 432

        HCNYield(1) = 0
        ObjectDescription(1) = ""
        ObjectItemID(1) = 1
        ObjectIgnTime(1) = 0
        ObjectMLUA(0, 1) = 0
        ObjectMLUA(1, 1) = 0
        ObjectIgnMode(1) = ""
        ObjectRLF(1) = 0.3
        ObjectLHoG(1) = 3
        ObjectWindEffect(1) = 1
        ObjectPyrolysisOption(1) = 0
        ObjectPoolDensity(1) = 684
        ObjectPoolDiameter(1) = 0
        ObjectPoolFBMLR(1) = 0.101
        ObjectPoolRamp(1) = 0
        ObjectPoolVolume(1) = 0
        'COYield(1) = 0.005
        FireHeight(1) = 0
        FireLocation(1) = 0
        NumberDataPoints(1) = 2
        HeatReleaseData(1, 1, 1) = 0
        HeatReleaseData(2, 1, 1) = 0
        MLRData(1, 1, 1) = 0
        MLRData(2, 1, 1) = 0


        'Room Description Defaults
        For i = 1 To NumberRooms
            'mechanical extraction
            ExtractRate(i) = 0
            NumberFans(i) = 1
            ExtractStartTime(i) = 0
            FanElevation(i) = RoomHeight(i)
            MaxPressure(i) = 50
            Extract(i) = True
            UseFanCurve(i) = True

            RoomWidth(i) = 5
            RoomLength(i) = 5
            RoomDescription(i) = ""
            RoomHeight(i) = 3
            FloorElevation(i) = 0
            RoomAbsX(i) = 0
            RoomAbsY(i) = 0
            CeilingSlope(i) = False
            TwoZones(i) = True
            MinStudHeight(i) = 3
            ' WallSurface(i) = "concrete"
            WallSurface(i) = "plasterboard"
            WallSubstrate(i) = "None"
            'CeilingSurface(i) = "concrete"
            CeilingSurface(i) = "plasterboard"
            CeilingSubstrate(i) = "None"
            FloorSubstrate(i) = "None"
            FloorSurface(i) = "concrete"

            'generates error
            'frmDescribeRoom.lblWallSurface.Text = ""
            'frmDescribeRoom.lblWallSurface.Text = WallSurface(i)
            'frmDescribeRoom.lblWallSubstrate.Text = WallSubstrate(i)
            'frmDescribeRoom.lblCeilingSurface.Text = CeilingSurface(i)
            'frmDescribeRoom.lblCeilingSubstrate.Text = CeilingSubstrate(i)
            'frmDescribeRoom.lblFloorSurface.Text = FloorSurface(i)

            ThermalInertiaWall(i) = 0.515
            ThermalInertiaCeiling(i) = 0.515
            ThermalInertiaFloor(i) = 0.515
            FloorConeDataFile(i) = "null.txt"
            WallConeDataFile(i) = "null.txt"
            CeilingConeDataFile(i) = "null.txt"
            WallSootYield(i) = 0
            WallH2OYield(i) = 0
            WallHCNYield(i) = 0
            WallCO2Yield(i) = 0
            CeilingSootYield(i) = 0
            FloorSootYield(i) = 0
            CeilingH2OYield(i) = 0
            CeilingHCNYield(i) = 0
            FloorH2OYield(i) = 0
            FloorHCNYield(i) = 0
            CeilingCO2Yield(i) = 0
            FloorCO2Yield(i) = 0
            'WallThickness(i) = 100
            WallThickness(i) = 13
            WallSubThickness(i) = 0
            ' CeilingThickness(i) = 100
            CeilingThickness(i) = 13
            CeilingSubThickness(i) = 0
            FloorSubThickness(i) = 0
            FloorThickness(i) = 100
            'CeilingConductivity(i) = 1.2
            CeilingConductivity(i) = 0.85
            CeilingSubConductivity(i) = 1
            'WallConductivity(i) = 1.2
            WallConductivity(i) = 0.85
            WallSubConductivity(i) = 1
            FloorSubConductivity(i) = 1
            FloorConductivity(i) = 1.2
            'WallSpecificHeat(i) = 880
            WallSpecificHeat(i) = 900
            'CeilingSpecificHeat(i) = 880
            CeilingSpecificHeat(i) = 900
            WallSubSpecificHeat(i) = 1000
            FloorSubSpecificHeat(i) = 1000
            CeilingSubSpecificHeat(i) = 1000
            FloorSpecificHeat(i) = 880
            'WallDensity(i) = 2300
            WallDensity(i) = 810
            'CeilingDensity(i) = 2300
            CeilingDensity(i) = 810
            WallSubDensity(i) = 2000
            FloorSubDensity(i) = 2000
            CeilingSubDensity(i) = 2000
            FloorDensity(i) = 2300
            'mechanical extraction
            ExtractRate(i) = 0
            NumberFans(i) = 1
            ExtractStartTime(i) = 0
            FanElevation(i) = RoomHeight(i)
            MaxPressure(i) = 50
            Extract(i) = True
            UseFanCurve(i) = True

            'call detector defaults
            Device_Defaults(i)

        Next i

        For i = 1 To 4
            Surface_Emissivity(i, 1) = 0.5
        Next i

        'frmSprink.cboDetectorType.SelectedIndex = 0

        'room corner
        WallHeatFlux = 45
        CeilingHeatFlux = 35
        'FlameAreaConstant = 0.067
        'FlameLengthPower = 2 / 3
        FlameAreaConstant = 0.0065
        FlameLengthPower = 1
        BurnerWidth = 0.17
        IgTempW(fireroom) = 742
        IgTempC(fireroom) = 742
        IgTempF(fireroom) = 742
        RCNone = True

        'endpoints
        VisibilityEndPoint = 10
        ConvectEndPoint = 80 + 273
        TargetEndPoint = 0.3
        TempEndPoint = 200 + 273
        FEDEndPoint = 0.3

        'Other Defaults
        NumberTimeSteps = 0
        Resize_Objects()
        NumberDataPoints(1) = 0
        NumberDataPoints(0) = 0

        MDIFrmMain.ToolStripStatusLabel1.Text = "NEW MODEL"
        MDIFrmMain.ToolStripStatusLabel2.Text = Description

        Dim matname, firename As String
        matname = Path.GetFileName(MaterialsDatabaseName)
        firename = Path.GetFileName(FireDatabaseName)
        Dim room_0 As New ComboBox
        Dim room_1 As New ComboBox
        Dim room_2 As New ComboBox
        Dim room_3 As New ComboBox

        'room_0 = frmDescribeRoom.ComboBox_RoomID
        'room_1 = frmDescribeRoom._lstRoomID_1
        'room_2 = frmDescribeRoom._lstRoomID_2
        'room_3 = frmDescribeRoom._lstRoomID_3

        Dim fireobj As New ComboBox
        Dim val As Integer = 0
        fireobj = frmFire.lstObjectLocation

        If fireobj.Items.Count > 0 Then
            fireobj.SelectedIndex = val
        End If


        IgnCorrelation = vbFTP

        frmOptions1.txtnC.Text = CStr(nC)
        frmOptions1.txtnH.Text = CStr(nH)
        frmOptions1.txtnO.Text = CStr(nO)
        frmOptions1.txtnN.Text = CStr(nN)
        frmOptions1.txtStoich.Text = CStr(StoichAFratio)
        frmOptions1.cboABSCoeff.Text = fueltype
        frmOptions1.txtEmissionCoefficient.Text = CStr(EmissionCoefficient)
        frmOptions1.txtSootEps.Text = CStr(SootEpsilon)
        frmOptions1.chkHCNcalc.CheckState = CheckState.Unchecked
        nowallflow = True
        frmOptions1.chkWallFlowDisable.CheckState = CheckState.Checked
        frmOptions1.optCOman.Checked = comode
        If comode = False Then frmOptions1.optCOauto.Checked = True
        frmOptions1.optSootman.Checked = sootmode
        If sootmode = False Then frmOptions1.optSootauto.Checked = True
        frmOptions1.txtpreCO.Text = CStr(preCO)
        frmOptions1.txtpostCO.Text = CStr(postCO)
        frmOptions1.txtpreSoot.Text = CStr(preSoot)
        frmOptions1.txtPostSoot.Text = CStr(postSoot)
        MDIFrmMain.NZBCVM2ToolStripMenuItem.Checked = UNCHECKED
        MDIFrmMain.RiskSimulatorToolStripMenuItem.Checked = CHECKED

        Call MDIFrmMain.VM2setup()

    End Sub

    Sub Normalise_Data(ByVal id As Integer, ByVal room As Integer)
        '*  ==================================================================
        '*  Normalise the hrr cone data by dividing through by the peak hrr
        '*
        '*  5 April 1998 Colleen Wade
        '*  ==================================================================

        Dim i, curve As Short
        curve = 1

        'UPGRADE_WARNING: Couldn't resolve default property of object id. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If id = 1 Or id = 4 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object room. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ReDim NormalHRR_W(1, ConeNumber_W(curve, room))
            'UPGRADE_WARNING: Couldn't resolve default property of object room. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For i = 1 To ConeNumber_W(curve, room)
                'UPGRADE_WARNING: Couldn't resolve default property of object room. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If PeakWallHRR(1, room) > 0 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object room. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    NormalHRR_W(curve, i) = ConeXW(room, curve, i) / PeakWallHRR(1, room)
                Else
                    NormalHRR_W(curve, i) = 0
                End If
            Next
            'UPGRADE_WARNING: Couldn't resolve default property of object id. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf id = 2 Or id = 5 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object room. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ReDim NormalHRR_C(1, ConeNumber_C(curve, room))
            'UPGRADE_WARNING: Couldn't resolve default property of object room. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For i = 1 To ConeNumber_C(curve, room)
                'UPGRADE_WARNING: Couldn't resolve default property of object room. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If PeakCeilingHRR(1, room) > 0 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object room. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    NormalHRR_C(curve, i) = ConeXC(room, curve, i) / PeakCeilingHRR(1, room)
                Else
                    NormalHRR_C(curve, i) = 0
                End If
            Next
            'UPGRADE_WARNING: Couldn't resolve default property of object id. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf id = 3 Or id = 6 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object room. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ReDim NormalHRR_F(1, ConeNumber_F(curve, room))
            'UPGRADE_WARNING: Couldn't resolve default property of object room. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For i = 1 To ConeNumber_F(curve, room)
                'UPGRADE_WARNING: Couldn't resolve default property of object room. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If PeakFloorHRR(1, room) > 0 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object room. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    NormalHRR_F(curve, i) = ConeXF(room, curve, i) / PeakFloorHRR(1, room)
                Else
                    NormalHRR_F(curve, i) = 0
                End If
            Next
        End If
    End Sub

    Sub Number_TimeSteps()
        '*  ==================================================================
        '*  Determine the number of time steps required
        '*  ==================================================================

        'find the integer part of the folllowing expresion
        NumberTimeSteps = Int(SimTime / Timestep)

        'If NumberTimeSteps > 30000 Then NumberTimeSteps = 30000

    End Sub

    Function O2Mass_Rate(ByVal Q As Double) As Double
        '*  ===================================================================
        '*  This function returns the value of the sum of the fuel mass loss rates
        '*  multiplied by the O2 depletion rate for all burning objects (kg-O2/sec)
        '*  13100 kJ energy assumed released per kg O2 consumed
        '*
        '*  Q = heat release at time t (kW) already adjusted for ventilation limit
        '*  ====================================================================

        O2Mass_Rate = -Q / 13100

    End Function



    Sub RegressionL(ByRef X() As Double, ByRef Y() As Double, ByRef n As Short, ByRef A As Double, ByRef b As Double, ByRef R As Double)
        '*************************************************************************************
        '*  Carry out a linear regression
        '*
        '*  Revised 23 February 1997 Colleen wade
        '*************************************************************************************

        Dim Sy2, sY, SX, Sx2, Sxy As Double
        Dim Numdata, PMarraybase, i As Short
        Dim Yy, Xx, temp As Double

        On Error GoTo Regresshandler

        PMarraybase = 1

        SX = 0.0# : sY = 0.0# : Sxy = 0.0# : Sx2 = 0.0# : Sy2 = 0.0#
        Numdata = n + 1 - PMarraybase
        For i = PMarraybase To n
            Xx = X(i) : Yy = Y(i)
            SX = SX + Xx
            sY = sY + Yy
            Sxy = Sxy + Xx * Yy
            Sx2 = Sx2 + Xx * Xx
            Sy2 = Sy2 + Yy * Yy
        Next i
        b = (Numdata * Sxy - SX * sY) / (Numdata * Sx2 - SX * SX)
        A = (sY - b * SX) / Numdata
        temp = (Numdata * Sx2 - SX * SX) * (Numdata * Sy2 - sY * sY)
        R = (Numdata * Sxy - SX * sY) / Sqrt(temp)
        Exit Sub

Regresshandler:
        flagstop = 1
        Exit Sub

    End Sub

    Sub Resize_Objects()
        '*  =================================================
        '*      This procedure resizes the object arrays to
        '*      optimize their size.
        '*
        '*      Global variables used:
        '*          NumberObjects, COYield, CO2Yield
        '*          SootYield, WaterVaporYield, EnergyYield
        '*
        '*  Revised by Colleen Wade 29 January 1997
        '*  =================================================

        If NumberObjects > 0 Then
            'ReDim Preserve FireHeight(NumberObjects + 1)
            'ReDim Preserve FireLocation(NumberObjects + 1)
            'ReDim Preserve EnergyYield(NumberObjects + 1)
            'ReDim Preserve CO2Yield(NumberObjects + 1)
            'ReDim Preserve SootYield(NumberObjects + 1)
            'ReDim Preserve WaterVaporYield(NumberObjects + 1)
            'ReDim Preserve HCNYield(NumberObjects + 1)
            'ReDim Preserve HCNuserYield(NumberObjects + 1)

            ReDim Preserve FireHeight(0 To NumberObjects)
            ReDim Preserve FireLocation(0 To NumberObjects)
            ReDim Preserve EnergyYield(0 To NumberObjects)
            ReDim Preserve CO2Yield(0 To NumberObjects)
            ReDim Preserve SootYield(0 To NumberObjects)
            ReDim Preserve WaterVaporYield(0 To NumberObjects)
            ReDim Preserve HCNYield(0 To NumberObjects)
            ReDim Preserve HCNuserYield(0 To NumberObjects)

            ReDim Preserve ObjElevation(0 To NumberObjects)
            ReDim Preserve ObjLabel(0 To NumberObjects)
            ReDim Preserve ObjLength(0 To NumberObjects)
            ReDim Preserve ObjWidth(0 To NumberObjects)
            ReDim Preserve ObjHeight(0 To NumberObjects)
            ReDim Preserve ObjDimX(0 To NumberObjects)
            ReDim Preserve ObjDimY(0 To NumberObjects)
            ReDim Preserve ObjCRF(0 To NumberObjects)
            ReDim Preserve ObjFTPindexpilot(0 To NumberObjects)
            ReDim Preserve ObjFTPlimitpilot(0 To NumberObjects)
            ReDim Preserve ObjCRFauto(0 To NumberObjects)
            ReDim Preserve ObjFTPindexauto(0 To NumberObjects)
            ReDim Preserve ObjFTPlimitauto(0 To NumberObjects)
            ReDim Preserve ObjectDescription(0 To NumberObjects)
            ReDim Preserve ObjectMass(0 To NumberObjects)
            ReDim Preserve ObjectItemID(0 To NumberObjects)
            ReDim Preserve ObjectIgnTime(0 To NumberObjects)
            ReDim Preserve ObjectMLUA(0 To 2, NumberObjects)
            ReDim Preserve ObjectLHoG(0 To NumberObjects)
            ReDim Preserve ObjectIgnMode(0 To NumberObjects)
            ReDim Preserve ObjectRLF(0 To NumberObjects)
            ReDim Preserve ObjectWindEffect(0 To NumberObjects)
            ReDim Preserve ObjectPyrolysisOption(0 To NumberObjects)
            ReDim Preserve ObjectPoolDensity(0 To NumberObjects)
            ReDim Preserve ObjectPoolDiameter(0 To NumberObjects)
            ReDim Preserve ObjectPoolFBMLR(0 To NumberObjects)
            ReDim Preserve ObjectPoolRamp(0 To NumberObjects)
            ReDim Preserve ObjectPoolVapTemp(0 To NumberObjects)
            ReDim Preserve ObjectPoolVolume(0 To NumberObjects)
            ReDim Preserve HeatReleaseData(2, 500, NumberObjects + 1)
            ReDim Preserve MLRData(2, 500, NumberObjects + 1)
            ' ReDim Preserve NumberDataPoints(NumberObjects + 1)
            ReDim Preserve NumberDataPoints(0 To NumberObjects)
            ReDim Preserve item_location(0 To 8, 0 To NumberObjects)
            ReDim Preserve NewRadiantLossFraction(0 To NumberObjects)
        End If

    End Sub

    Sub Resize_Vents()
        '*  =================================================
        '*      This procedure resizes the vent arrays to
        '*      optimise their size.
        '*
        '*  =================================================

        Dim i, j As Short

        MaxNumberVents = 1
        For i = 1 To NumberRooms
            If i > MaxNumberRooms Then Exit For
            For j = 2 To NumberRooms + 1
                If j > MaxNumberRooms + 1 Then Exit For
                If NumberVents(i, j) > MaxNumberVents Then
                    MaxNumberVents = NumberVents(i, j)
                End If
            Next j
        Next i

        If MaxNumberVents > 0 Then
            ReDim Preserve VentCD(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve VentProb(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)

            ReDim Preserve soffitheight(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve VentHeight(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve VentFace(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve HOReliability(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve HOactive(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve VentOffset(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve VentSillHeight(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve VentWidth(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve WallLength1(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve WallLength2(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve VentOpenTime(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve VentCloseTime(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve FRintegrity(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve FRMaxOpening(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve FRMaxOpeningTime(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve FRgastemp(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve FRcriteria(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)

            ReDim Preserve AutoOpenVent(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents) '2009.2
            ReDim Preserve trigger_device(0 To 6, MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents) '2009.2
            ReDim Preserve SDtriggerroom(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents) '2009.2
            ReDim Preserve HRR_ventopenduration(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents) '2009.2
            ReDim Preserve HRR_ventopendelay(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents) '2009.2
            ReDim Preserve trigger_ventopenduration(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents) '2009.2
            ReDim Preserve HRR_threshold(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents) '2009.2
            ReDim Preserve trigger_ventopendelay(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents) '2009.2
            ReDim Preserve HRR_threshold_time(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents) '2009.2
            ReDim Preserve HRRFlag(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents) '2009.2
            ReDim Preserve VentBreakTime(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents) '2009.24
            ReDim Preserve VentBreakTimeClose(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents) '2009.24

            ReDim Preserve SpillPlumeBalc(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve Downstand(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve spillplume(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve spillplumemodel(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve spillbalconyprojection(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)

            ReDim Preserve AutoBreakGlass(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve GLASSconductivity(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve GLASSemissivity(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve GLASSexpansion(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve GlassFlameFlux(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve GLASSthickness(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve GLASSdistance(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve GLASSFalloutTime(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve GLASSshading(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve GLASSbreakingstress(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve GLASSalpha(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve GlassYoungsModulus(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve breakflag(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            ReDim Preserve Highest_Vent(MaxNumberRooms + 1)
            ReDim Preserve Lowest_Vent(MaxNumberRooms + 1)
            ReDim Preserve WAfloor_flag(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
        End If
    End Sub

    Sub Room_FloorArea()
        '*  =================================================================
        '*  Calculate the floor area of the room.
        '*  =================================================================
        Dim i As Short
        For i = 1 To NumberRooms
            RoomFloorArea(i) = RoomWidth(i) * RoomLength(i)
        Next i
    End Sub

    Sub Room_Volume()
        '*  =================================================================
        '*  Calculate the volume of the room.
        '*
        '*  Revised 14 October 1996 Colleen Wade
        '*  =================================================================
        Dim i As Short

        'If frmRoomDimensions.optFlatCeiling.Value = True Or RoomHeight = MinStudHeight Then
        For i = 1 To NumberRooms
            If CeilingSlope(i) = False Or RoomHeight(i) = MinStudHeight(i) Then
                'for a flat ceiling
                RoomVolume(i) = RoomWidth(i) * RoomLength(i) * RoomHeight(i)
                'ElseIf frmRoomDimensions.optSlopeCeiling.Value = True Then
            ElseIf CeilingSlope(i) = True Then
                'sloping ceiling
                RoomVolume(i) = RoomWidth(i) * RoomLength(i) * (MinStudHeight(i) + RoomHeight(i)) / 2
            End If
        Next i
    End Sub




    Function Runge_Kutta_O2(ByRef mplume As Double, ByRef MT As Double, ByRef MO2 As Double, ByRef UT As Double, ByVal uvol As Double, ByRef Y As Double) As Double
        '*  =====================================================================
        '*  This function returns the value of the derivative function for the
        '*  mass fraction of O2 in the upper layer
        '*
        '*  MPlume = mass flow in the plume (kg/s)
        '*  UT = upper layer temperature (K)
        '*  UVol = upper layer volume (m3)
        '*  Y = O2 mass fraction
        '*  MT = total mass loss rate (kg/s)
        '*  MO2 = oxygen depletion rate (kg O2/sec)
        '*
        '*  Revised 29 November 1996 Colleen Wade
        '*  =====================================================================

        Dim con As Double

        If uvol <= 0.001 * RoomVolume(1) Then uvol = 0.001 * RoomVolume(1)

        'some convenient constants for the calculation
        con = ReferenceDensity * ReferenceTemp * uvol

        'for all layer heights below the base of the fire
        'air entrained in the plume is zero

        Runge_Kutta_O2 = UT / (con) * (mplume * (O2MassFraction(1, 1, 1) - Y) + (MO2 - MT * Y))

    End Function

    Function Runge_Kutta_Soot(ByRef mplume As Double, ByRef MT As Double, ByRef MSoot As Double, ByRef UT As Double, ByVal uvol As Double, ByRef Y As Double) As Double
        '*  =====================================================================
        '*  This function returns the value of the derivative function for the
        '*  mass fraction of Soot in the upper layer
        '*
        '*  MPlume = mass flow in the plume (kg/s)
        '*  UT = upper layer temperature (K)
        '*  Z = layer height above fire (m)
        '*  Y = Soot mass fraction
        '*  MT = total fuel mass loss rate (kg/s)
        '*  MSoot = Soot mass production rate (kg/s)
        '*
        '*  Revised 30 September 1996 Colleen Wade
        '*  ======================================================================

        Dim con As Double
        On Error GoTo RKsoothandler

        If uvol <= 0.001 * RoomVolume(1) Then uvol = 0.001 * RoomVolume(1)

        'some convenient constants for the calculation
        con = ReferenceDensity * ReferenceTemp * uvol

        'for all other layer heights above the base of the fire
        'for all layer heights below the base of the fire
        'air entrained in the plume is zero
        Runge_Kutta_Soot = UT / (con) * (mplume * (SootMassFraction(1, 1, 1) - Y) + (MSoot - MT * Y))
        Exit Function

RKsoothandler:
        MsgBox(ErrorToString(Err.Number) & " in Runge_Kutta_Soot")
        ERRNO = Err.Number
        'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'System.Windows.Forms.Cursor.Current = Default_Renamed
        Exit Function

    End Function

    Function Runge_Kutta_UT(ByRef mplume As Double, ByRef Q As Double, ByRef UT As Double, ByVal uvol As Double, ByRef qloss As Double, ByRef LT As Double) As Double
        ''*  =====================================================================
        ''*  Find the value of the derivative for the temperature of the upper gas
        ''*  layer
        ''*
        ''*  MPlume = mass flow in the plume (kg/s)
        ''*  Q = total heat release from fire (kW)
        ''*  UT = upper layer temp (K)
        ''*  UVol = upper layer volume (m3)
        ''*  QLoss = losses from upper layer to room surfaces (kW)
        ''*
        ''*  Revised 17/6/97 Colleen Wade
        ''*  =====================================================================

        'Dim con As Double

        'If uvol <= 0.01 * RoomVolume(1) Then
        '    uvol = 0.01 * RoomVolume(1)
        'ElseIf uvol > 0.99 * RoomVolume(1) Then
        '    uvol = 0.99 * RoomVolume(1)
        'End If

        ''a convenient constant for the calculation
        'con = ReferenceTemp * SpecificHeat_air * ReferenceDensity * uvol

        ''Runge_Kutta_UT = -(SpecificHeat_air * mplume * (UT - LT) - (1 - RadiantLossFraction) * Q + qloss) * UT / con
        'Runge_Kutta_UT = -(SpecificHeat_air * mplume * (UT - LT) - (1 - NewRadiantLossFraction) * Q + qloss) * UT / con

    End Function


    Sub Save_Results(ByRef OpenResultsFile As Object)
        '*  ======================================================
        '*  This procedure saves the results to an ascii file so
        '*  they can be accessed and imported by spreadsheets etc
        '*  ======================================================

        Dim s As String
        Dim j As Short

        'define a format string
        s = "Scientific"

        'UPGRADE_WARNING: Couldn't resolve default property of object OpenResultsFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileOpen(1, OpenResultsFile, OpenMode.Output)

        PrintLine(1, "Time (sec)", "Layer (m)", "Upper Temp (C)", "HRR (kW)", "Mass (kg/s)", "Plume (kg/s)", "Vent (kg/s)", "CO2 (%)", "CO (%)", "O2 (%)", "FED gas(inc)", "Wall Temp (C)", "Ceiling Temp (C)", "Rad on Floor (kW/m2)", "Lower Temp (C)")

        For j = 1 To NumberTimeSteps
            'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            If tim(j, 1) Mod DisplayInterval = 0 Then
                'PrintLine(1, VB6.Format(tim(j, 1), s), VB6.Format(layerheight(1, j), s), VB6.Format(uppertemp(1, j) - 273, s), VB6.Format(HeatRelease(fireroom, j, 2), s), VB6.Format(FuelMassLossRate(j, fireroom), s), VB6.Format(massplumeflow(j, fireroom), s), VB6.Format(VentMassOut(j, 1), s), VB6.Format(CO2VolumeFraction(1, j, 1) * 100, s), VB6.Format(COVolumeFraction(1, j, 1) * 100, s), VB6.Format(O2VolumeFraction(1, j, 1) * 100, s), VB6.Format(FEDSum(1, j), s), VB6.Format(Upperwalltemp(1, j) - 273, s), VB6.Format(CeilingTemp(fireroom, j) - 273, s), VB6.Format(Target(1, j), s), VB6.Format(lowertemp(1, j) - 273, s))
            End If
        Next j
        FileClose(1)

    End Sub

    Sub Size_Arrays()
        '*  ==================================================================
        '*  Redimension the size of arrays used to store simulation results.
        '*
        '*  24/11/97 C A Wade
        '*  ==================================================================

        'Dim i As Short
        Dim i As Integer

        i = NumberTimeSteps + 1

        ReDim tim(i, 1)
        ReDim HeatRelease(NumberRooms + 1, i, 2)
        ReDim FuelMassLossRate(i, NumberRooms)
        ReDim WoodBurningRate(i)
        ReDim FuelBurningRate(0 To 5, 0 To NumberRooms, 0 To i)
        'FuelBurningRate(0, 0 To NumberRooms, 0 To i) - user supplied free-burning mass loss rate
        'FuelBurningRate(1, 0 To NumberRooms, 0 To i) - ventilation effect mass loss rate (including O2 vitiated)
        'FuelBurningRate(2, 0 To NumberRooms, 0 To i) - thermal effect burning rate 
        'FuelBurningRate(3, 0 To NumberRooms, 0 To i) - total mass loss rate 
        'FuelBurningRate(4, 0 To NumberRooms, 0 To i) - burning rate (contributing to energy release in the enclosure)
        'FuelBurningRate(5, 0 To NumberRooms, 0 To i) - fuel area shrinkage ratio AFB/AF

        ReDim SootMassGen(i)
        ReDim FuelHoC(i)

        If NumSprinklers > 0 Then
            ReDim CJetTemp(i, 2, NumSprinklers - 1)
        Else
            ReDim CJetTemp(i, 2, 0)
        End If

        ReDim TotalFuel(i)
        ReDim massplumeflow(i, NumberRooms)
        ReDim qloss(i)

        ReDim layerheight(NumberRooms, i)
        ReDim UpperVolume(NumberRooms, i)
        ReDim uppertemp(NumberRooms, i)
        ReDim lowertemp(NumberRooms, i)
        ReDim upperemissivity(NumberRooms, i)
        ReDim RoomPressure(NumberRooms, i)
        ReDim energy(NumberRooms, i)
        ReDim VentMassOut(i, 4)
        ReDim UFlowToOutside(NumberRooms, i)
        ReDim FlowGER(NumberRooms, i)
        ReDim SPR(NumberRooms, i)
        ReDim FlowToUpper(NumberRooms, i)
        ReDim FlowToLower(NumberRooms, i)
        ReDim WallFlowtoUpper(NumberRooms, i)
        ReDim WallFlowtoLower(NumberRooms, i)
        ReDim ventflow(i, NumberRooms + 1, 2) 'cw 18/3/2010
        ReDim LinkTemp(NumberRooms, i)
        ReDim SprinkTemp(NumSprinklers, i)
        ReDim SmokeConcentration(NumberRooms, i)
        ReDim SmokeConcentrationSD(NumSmokeDetectors, i)
        ReDim OD_outside(NumberRooms, i)
        ReDim OD_outsideSD(NumSmokeDetectors, i)
        ReDim OD_insideSD(NumSmokeDetectors, i)
        ReDim OD_inside(NumberRooms, i)
        ReDim ventfire(NumberRooms + 1, i)
        ReDim TUHC(NumberRooms + 1, i, 2)
        ReDim CO2MassFraction(NumberRooms + 1, i, 2)
        ReDim CO2VolumeFraction(NumberRooms + 1, i, 2)
        ReDim SootMassFraction(NumberRooms + 1, i, 2)
        ReDim H2OMassFraction(NumberRooms + 1, i, 2)
        ReDim HCNMassFraction(NumberRooms + 1, i, 2)
        ReDim COMassFraction(NumberRooms + 1, i, 2)
        ReDim GlobalER(i)
        ReDim absorb(NumberRooms, i, 2)
        ReDim COVolumeFraction(NumberRooms + 1, i, 2)
        ReDim H2OVolumeFraction(NumberRooms + 1, i, 2)
        ReDim HCNVolumeFraction(NumberRooms + 1, i, 2)
        ReDim OD_upper(NumberRooms, i + 1)
        ReDim OD_lower(NumberRooms, i + 1)
        ReDim Visibility(NumberRooms, i + 1)
        ReDim FEDSum(NumberRooms, i + 1)
        ReDim FEDRadSum(NumberRooms, i + 1)
        ReDim SurfaceRad(NumberRooms, i + 1)
        ReDim Target(NumberRooms, i)
        ReDim O2MassFraction(NumberRooms + 1, i, 2)
        ReDim O2VolumeFraction(NumberRooms + 1, i, 2)
        ReDim IgFloorTemp(NumberRooms, i)
        ReDim IgWallTemp(NumberRooms, i)
        ReDim IgCeilingTemp(NumberRooms, i)
        ReDim CeilingIgniteObject(NumberRooms)
        ReDim WallIgniteObject(NumberRooms)
        ReDim WallIgniteFlag(NumberRooms)
        ReDim FloorIgniteFlag(NumberRooms)
        ReDim CeilingIgniteFlag(NumberRooms)
        ReDim FloorIgniteTime(NumberRooms)
        ReDim CeilingIgniteTime(NumberRooms)
        ReDim CeilingIgniteStep(NumberRooms)
        ReDim FloorIgniteStep(NumberRooms)
        ReDim WallIgniteTime(NumberRooms)
        ReDim WallIgniteStep(NumberRooms)
        ReDim Upperwalltemp(NumberRooms, i)
        ReDim UnexposedUpperwalltemp(NumberRooms, i)
        ReDim UnexposedLowerwalltemp(NumberRooms, i)
        ReDim UnexposedCeilingtemp(NumberRooms, i)
        ReDim UnexposedFloortemp(NumberRooms, i)
        ReDim LowerWallTemp(NumberRooms, i)
        ReDim CeilingTemp(NumberRooms, i)
        ReDim FloorTemp(NumberRooms, i)
        ReDim QCeiling(NumberRooms, i)
        ReDim QUpperWall(NumberRooms, i)
        ReDim QLowerWall(NumberRooms, i)
        ReDim QFloor(NumberRooms, i)
        ReDim QCeilingAST(NumberRooms, 3, i)
        ReDim QUpperWallAST(NumberRooms, 3, i)
        ReDim QLowerWallAST(NumberRooms, 3, i)
        ReDim QFloorAST(NumberRooms, 3, i)
        ReDim TotalLosses(NumberRooms, i)
        ReDim NHL(0 To 6, NumberRooms, i)
        '0=q/root k rho c area weighted for comp
        '1=NHL area weighted for comp
        '2=time in harmathy furnace
        '3=ceiling NHL
        '4=upper wall NHL
        '5=lower wall NHL
        '6=floor NHL
        ReDim Y_pyrolysis(NumberRooms, i + 1)
        ReDim Yf_pyrolysis(NumberRooms, i + 1)
        ReDim Y_burnout(NumberRooms, i + 1)
        ReDim Yf_burnout(NumberRooms, i + 1)
        ReDim X_pyrolysis(NumberRooms, i + 1)
        ReDim Xf_pyrolysis(NumberRooms, i + 1)
        ReDim Z_pyrolysis(NumberRooms, i + 1)
        ReDim FlameVelocity(NumberRooms, 3, i + 1) '1=upward velocity, 2=lateral velocity 3=upward burnout
        ReDim FloorVelocity(NumberRooms, 3, i + 1) '1=upward velocity, 2=lateral velocity 3=upward burnout
        ReDim pyrolarea(NumberRooms, 3, i) '1=wall, 2=ceiling areas, 3=floor area
        ReDim delta_area(NumberRooms, 3, i) '1=wall, 2=ceiling areas, 3=floor area

    End Sub

    Sub Soffit_Height()
        '*  ==================================================================
        '*  Calculate the height of the soffit with respect to the floor
        '*  for each vent.
        '*  ==================================================================

        Dim j, i, k As Short

        For j = 1 To NumberRooms
            For k = 1 To NumberRooms + 1
                If NumberVents(j, k) > 0 Then
                    For i = 1 To NumberVents(j, k)
                        soffitheight(j, k, i) = VentSillHeight(j, k, i) + VentHeight(j, k, i)
                    Next i
                End If
            Next k
        Next j

    End Sub

    Function SootMass_Rate(ByVal T As Double, ByVal step_Renamed As Integer) As Double
        '*  ==================================================================
        '*  This function returns the value of the sum of the fuel mass loss rates
        '*  multiplied by the soot generation rate for all burning objects.
        '*  called once at each timestep for each room
        '*  5 September 2008 - C Wade
        '*  ==================================================================

        Dim dummy, mlo As Double
        Dim QCeiling, total, QWall, QFloor As Double
        Dim i As Integer

        For i = 1 To NumberObjects

            If frmOptions1.optRCNone.Checked = True Or i > 1 Then

                If sootmode = True Then 'manual entry of pre post flashover soot yields

                    If Flashover = True Then
                        'in this case Qburner, QFloor, QWall, QCeiling are all zero
                        dummy = MassLoss_Object(i, T, Qburner, QFloor, QWall, QCeiling) * postSoot
                    Else

                        If VentilationLimitFlag = True And VM2 = True Then
                            'case 5 VM2 rules
                            'in this case Qburner, QFloor, QWall, QCeiling are all zero
                            dummy = MassLoss_Object(i, T, Qburner, QFloor, QWall, QCeiling) * postSoot

                        Else
                            'in this case Qburner, QFloor, QWall, QCeiling are all zero
                            dummy = MassLoss_Object(i, T, Qburner, QFloor, QWall, QCeiling) * preSoot

                        End If

                    End If
                Else

                    'in this case Qburner, QFloor, QWall, QCeiling are all zero
                    dummy = MassLoss_Object(i, T, Qburner, QFloor, QWall, QCeiling) * SootYield(i)

                End If

            Else
                'room corner simulation with flame spread
                'may need checking, does this work properly with manual input of soot yields?
                mlo = MassLoss_Object(i, T, Qburner, QFloor, QWall, QCeiling)
                If WallEffectiveHeatofCombustion(fireroom) > 0 And CeilingEffectiveHeatofCombustion(fireroom) > 0 Then

                    If sootmode = True Then 'manual entry of pre post flashover soot yields
                        'If GlobalER(step_Renamed) > 1 Then
                        If Flashover = True Then

                            dummy = Qburner / (NewEnergyYield(i) * 1000) * postSoot + QWall / (WallEffectiveHeatofCombustion(fireroom) * 1000) * WallSootYield(fireroom) + QCeiling / (CeilingEffectiveHeatofCombustion(fireroom) * 1000) * CeilingSootYield(fireroom)
                        Else
                            dummy = Qburner / (NewEnergyYield(i) * 1000) * preSoot + QWall / (WallEffectiveHeatofCombustion(fireroom) * 1000) * WallSootYield(fireroom) + QCeiling / (CeilingEffectiveHeatofCombustion(fireroom) * 1000) * CeilingSootYield(fireroom)
                        End If
                    Else
                        dummy = Qburner / (NewEnergyYield(i) * 1000) * SootYield(i) + QWall / (WallEffectiveHeatofCombustion(fireroom) * 1000) * WallSootYield(fireroom) + QCeiling / (CeilingEffectiveHeatofCombustion(fireroom) * 1000) * CeilingSootYield(fireroom)
                    End If

                ElseIf CeilingEffectiveHeatofCombustion(fireroom) <= 0 And WallEffectiveHeatofCombustion(fireroom) > 0 Then

                    If sootmode = True Then 'manual entry of pre post flashover soot yields
                        'If GlobalER(step_Renamed) > 1 Then
                        If Flashover = True Then
                            dummy = Qburner / (NewEnergyYield(i) * 1000) * postSoot + QWall / (WallEffectiveHeatofCombustion(fireroom) * 1000) * WallSootYield(fireroom)
                        Else
                            dummy = Qburner / (NewEnergyYield(i) * 1000) * preSoot + QWall / (WallEffectiveHeatofCombustion(fireroom) * 1000) * WallSootYield(fireroom)
                        End If
                    Else
                        dummy = Qburner / (NewEnergyYield(i) * 1000) * SootYield(i) + QWall / (WallEffectiveHeatofCombustion(fireroom) * 1000) * WallSootYield(fireroom)
                    End If

                ElseIf WallEffectiveHeatofCombustion(fireroom) <= 0 And CeilingEffectiveHeatofCombustion(fireroom) > 0 Then
                    If sootmode = True Then 'manual entry of pre post flashover soot yields
                        'If GlobalER(step_Renamed) > 1 Then
                        If Flashover = True Then
                            dummy = Qburner / (NewEnergyYield(i) * 1000) * postSoot + QCeiling / (CeilingEffectiveHeatofCombustion(fireroom) * 1000) * CeilingSootYield(fireroom)
                        Else
                            dummy = Qburner / (NewEnergyYield(i) * 1000) * preSoot + QCeiling / (CeilingEffectiveHeatofCombustion(fireroom) * 1000) * CeilingSootYield(fireroom)
                        End If
                    Else
                        dummy = Qburner / (NewEnergyYield(i) * 1000) * SootYield(i) + QCeiling / (CeilingEffectiveHeatofCombustion(fireroom) * 1000) * CeilingSootYield(fireroom)
                    End If

                End If
            End If
            total = total + dummy
        Next i

        SootMass_Rate = total

    End Function
    Function SootMass_Rate_withfuelresponse(ByVal T As Double, ByVal step_Renamed As Integer) As Double
        '*  ==================================================================
        '*  This function returns the value of the sum of the fuel mass loss rates
        '*  multiplied by the soot generation rate for all burning objects.
        '*  called once at each timestep for each room
        '*  5 September 2008 - C Wade
        '*  ==================================================================

        Dim dummy, mlo As Double
        Dim QCeiling, total, QWall, QFloor As Double
        Dim i As Integer
        Dim o2lower = O2MassFraction(fireroom, stepcount, 2)
        Dim incidentflux = Target(fireroom, stepcount)
        Dim mplume = massplumeflow(stepcount, fireroom)
        Dim ltemp = lowertemp(fireroom, stepcount)

        For i = 1 To NumberObjects

            If frmOptions1.optRCNone.Checked = True Or i > 1 Then

                If sootmode = True Then 'manual entry of pre post flashover soot yields

                    If Flashover = True Then
                        'in this case Qburner, QFloor, QWall, QCeiling are all zero
                        dummy = MassLoss_ObjectwithFuelResponse(i, T, Qburner, o2lower, ltemp, mplume, incidentflux, 0) * postSoot
                    Else

                        If VentilationLimitFlag = True And VM2 = True Then
                            'case 5 VM2 rules
                            'in this case Qburner, QFloor, QWall, QCeiling are all zero
                            dummy = MassLoss_ObjectwithFuelResponse(i, T, Qburner, o2lower, ltemp, mplume, incidentflux, 0) * postSoot

                        Else
                            'in this case Qburner, QFloor, QWall, QCeiling are all zero
                            dummy = MassLoss_ObjectwithFuelResponse(i, T, Qburner, o2lower, ltemp, mplume, incidentflux, 0) * preSoot

                        End If

                    End If
                Else

                    'in this case Qburner, QFloor, QWall, QCeiling are all zero
                    dummy = MassLoss_ObjectwithFuelResponse(i, T, Qburner, o2lower, ltemp, mplume, incidentflux, 0) * SootYield(i)

                End If

            Else
                'room corner simulation with flame spread
                'may need checking, does this work properly with manual input of soot yields?
                mlo = MassLoss_ObjectwithFuelResponse(i, T, Qburner, o2lower, mplume, ltemp, incidentflux, 0)
                If WallEffectiveHeatofCombustion(fireroom) > 0 And CeilingEffectiveHeatofCombustion(fireroom) > 0 Then

                    If sootmode = True Then 'manual entry of pre post flashover soot yields
                        'If GlobalER(step_Renamed) > 1 Then
                        If Flashover = True Then

                            dummy = Qburner / (NewEnergyYield(i) * 1000) * postSoot + QWall / (WallEffectiveHeatofCombustion(fireroom) * 1000) * WallSootYield(fireroom) + QCeiling / (CeilingEffectiveHeatofCombustion(fireroom) * 1000) * CeilingSootYield(fireroom)
                        Else
                            dummy = Qburner / (NewEnergyYield(i) * 1000) * preSoot + QWall / (WallEffectiveHeatofCombustion(fireroom) * 1000) * WallSootYield(fireroom) + QCeiling / (CeilingEffectiveHeatofCombustion(fireroom) * 1000) * CeilingSootYield(fireroom)
                        End If
                    Else
                        dummy = Qburner / (NewEnergyYield(i) * 1000) * SootYield(i) + QWall / (WallEffectiveHeatofCombustion(fireroom) * 1000) * WallSootYield(fireroom) + QCeiling / (CeilingEffectiveHeatofCombustion(fireroom) * 1000) * CeilingSootYield(fireroom)
                    End If

                ElseIf CeilingEffectiveHeatofCombustion(fireroom) <= 0 And WallEffectiveHeatofCombustion(fireroom) > 0 Then

                    If sootmode = True Then 'manual entry of pre post flashover soot yields
                        'If GlobalER(step_Renamed) > 1 Then
                        If Flashover = True Then
                            dummy = Qburner / (NewEnergyYield(i) * 1000) * postSoot + QWall / (WallEffectiveHeatofCombustion(fireroom) * 1000) * WallSootYield(fireroom)
                        Else
                            dummy = Qburner / (NewEnergyYield(i) * 1000) * preSoot + QWall / (WallEffectiveHeatofCombustion(fireroom) * 1000) * WallSootYield(fireroom)
                        End If
                    Else
                        dummy = Qburner / (NewEnergyYield(i) * 1000) * SootYield(i) + QWall / (WallEffectiveHeatofCombustion(fireroom) * 1000) * WallSootYield(fireroom)
                    End If

                ElseIf WallEffectiveHeatofCombustion(fireroom) <= 0 And CeilingEffectiveHeatofCombustion(fireroom) > 0 Then
                    If sootmode = True Then 'manual entry of pre post flashover soot yields
                        'If GlobalER(step_Renamed) > 1 Then
                        If Flashover = True Then
                            dummy = Qburner / (NewEnergyYield(i) * 1000) * postSoot + QCeiling / (CeilingEffectiveHeatofCombustion(fireroom) * 1000) * CeilingSootYield(fireroom)
                        Else
                            dummy = Qburner / (NewEnergyYield(i) * 1000) * preSoot + QCeiling / (CeilingEffectiveHeatofCombustion(fireroom) * 1000) * CeilingSootYield(fireroom)
                        End If
                    Else
                        dummy = Qburner / (NewEnergyYield(i) * 1000) * SootYield(i) + QCeiling / (CeilingEffectiveHeatofCombustion(fireroom) * 1000) * CeilingSootYield(fireroom)
                    End If

                End If
            End If
            total = total + dummy
        Next i

        SootMass_Rate_withfuelresponse = total

    End Function

    Sub Update_EndPointCaptions()
        '*  ======================================================
        '*      This event checks for empty captions for the endpoint
        '*      condition, and substitutes and appropriate message
        '*      if found to be blank.
        '*
        '*  Revised 19 October 1996 Colleen Wade
        '*  ======================================================

        If frmEndPoints.lblTarget.Text = "" Then
            frmEndPoints.lblTarget.Text = "FED Thermal (incap) of " & VB6.Format(TargetEndPoint) & "  Not Reached."
        End If

        If frmEndPoints.lblTemp.Text = "" Then
            frmEndPoints.lblTemp.Text = "An Upper Layer Temperature of " & VB6.Format(TempEndPoint - 273) & " deg C Not Reached."
        End If

        If frmEndPoints.lblVisibility.Text = "" Then
            frmEndPoints.lblVisibility.Text = "Visibility at " & VB6.Format(MonitorHeight, "0.0") & " m Above Floor Did Not Reduce to " & VB6.Format(VisibilityEndPoint) & " m."
        End If

        If frmEndPoints.lblConvect.Text = "" Then
            frmEndPoints.lblConvect.Text = "Temperature at " & VB6.Format(MonitorHeight, "0.0") & " m Above Floor Did Not Reach " & VB6.Format(ConvectEndPoint - 273) & " deg C."
        End If

        If frmEndPoints.lblSprinkler.Text = "" Then
            frmEndPoints.lblSprinkler.Text = "The Sprinkler/Detector Did Not Actuate."
        End If

        If frmEndPoints.lblFED.Text = "" Then
            frmEndPoints.lblFED.Text = "The FED narcotic gases did not reach an incapacitating dose of " & VB6.Format(FEDEndPoint)
        End If

    End Sub

    Function Virtual_Source(ByVal Q As Single, ByVal diameter As Double) As Double
        '*  ====================================================================
        '*  This function returns the value of the virtual origin, for use in
        '*  plume entrainment models.
        '*
        '*  Arguments passed to the function are:
        '*      Q           =   actual heat release rate (kW)
        '*      Diameter    =   diameter of the fire (m)
        '*
        '*  Revised 5 December 1995 Colleen Wade
        '*  ====================================================================

        If Q < 0 Then Q = 0

        'Heskestad's relation for virtual source of the plume
        Virtual_Source = -1.02 * diameter + 0.083 * Q ^ (2 / 5)

    End Function

    Sub Volume_Fractions()
        '*  ====================================================================
        '*  This subprocedure calculates species volume fraction from mass fractions
        '*  ====================================================================

        Dim i, j As Integer
        Dim mw_upper, mw_lower As Double

        For j = 1 To NumberRooms
            For i = 1 To NumberTimeSteps + 1
                mw_upper = MolecularWeightHCN * HCNMassFraction(j, i, 1) + MolecularWeightCO * COMassFraction(j, i, 1) + MolecularWeightCO2 * CO2MassFraction(j, i, 1) + MolecularWeightH2O * H2OMassFraction(j, i, 1) + MolecularWeightO2 * O2MassFraction(j, i, 1) + MolecularWeightN2 * DMax1(0, 1 - O2MassFraction(j, i, 1) - HCNMassFraction(j, i, 1) - H2OMassFraction(j, i, 1) - COMassFraction(j, i, 1) - CO2MassFraction(j, i, 1))
                mw_lower = MolecularWeightHCN * HCNMassFraction(j, i, 2) + MolecularWeightCO * COMassFraction(j, i, 2) + MolecularWeightCO2 * CO2MassFraction(j, i, 2) + MolecularWeightH2O * H2OMassFraction(j, i, 2) + MolecularWeightO2 * O2MassFraction(j, i, 2) + MolecularWeightN2 * DMax1(0, 1 - O2MassFraction(j, i, 2) - HCNMassFraction(j, i, 2) - H2OMassFraction(j, i, 2) - COMassFraction(j, i, 2) - CO2MassFraction(j, i, 2))
                CO2VolumeFraction(j, i, 1) = CO2MassFraction(j, i, 1) * mw_upper / MolecularWeightCO2
                H2OVolumeFraction(j, i, 1) = H2OMassFraction(j, i, 1) * mw_upper / MolecularWeightH2O
                HCNVolumeFraction(j, i, 1) = HCNMassFraction(j, i, 1) * mw_upper / MolecularWeightHCN
                COVolumeFraction(j, i, 1) = COMassFraction(j, i, 1) * mw_upper / MolecularWeightCO
                O2VolumeFraction(j, i, 1) = O2MassFraction(j, i, 1) * mw_upper / MolecularWeightO2
                CO2VolumeFraction(j, i, 2) = CO2MassFraction(j, i, 2) * mw_lower / MolecularWeightCO2
                H2OVolumeFraction(j, i, 2) = H2OMassFraction(j, i, 2) * mw_lower / MolecularWeightH2O
                HCNVolumeFraction(j, i, 2) = HCNMassFraction(j, i, 2) * mw_lower / MolecularWeightHCN
                COVolumeFraction(j, i, 2) = COMassFraction(j, i, 2) * mw_lower / MolecularWeightCO
                O2VolumeFraction(j, i, 2) = O2MassFraction(j, i, 2) * mw_lower / MolecularWeightO2

                If COVolumeFraction(j, i, 1) < 0 Then COVolumeFraction(j, i, 1) = 0
                If COVolumeFraction(j, i, 2) < 0 Then COVolumeFraction(j, i, 2) = 0
            Next i
        Next j
    End Sub
    Public Sub Graph_Cone_Data(ByVal room As Integer, ByRef timedata(,,) As Double, ByVal YAxisTitle As String, ByRef datatobeplotted(,,) As Double, ByVal DataShift As Single, ByRef multiplier As Single, ByVal Style As Short, ByVal ymax As Single, ByVal NumberCurves(,) As Integer)

        'Dim graph As System.Windows.Forms.Control
        Dim points As Short
        Dim i, j, curves, curve As Short
        Dim roomid As Integer
        Dim flux(,) As Double

        If frmOptions1.optRCNone.Checked = True Then
            Exit Sub
        End If

        Try

            'if no data exists
            '  If NumberCurves(< 1 Then Exit Sub
            frmPlot.Chart1.Series.Clear()
            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            'frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right

            If YAxisTitle <> "Ceiling Cone HRR (kW/m2)" Then
                If YAxisTitle = "Wall Cone HRR (kW/m2)" Then
                    'If CurveNumber_W(room) = 0 Then
                    For roomid = 1 To NumberRooms
                        If WallConeDataFile(room) = "" Then WallConeDataFile(room) = "null.txt"
                        Call Flame_Spread_Props(room, 1, wallhigh, WallConeDataFile(room), Surface_Emissivity(2, room), ThermalInertiaWall(room), IgTempW(room), WallEffectiveHeatofCombustion(room), WallHeatofGasification(room), AreaUnderWallCurve(room), WallFTP(room), WallQCrit(room), Walln(room))
                    Next roomid
                    'End If
                    points = UBound(ConeXW, 3)
                    datatobeplotted = ConeXW
                    flux = ExtFlux_W
                    curves = CurveNumber_W(room)
                    'Do While CurveNumber_W(room) * points / factor > 3800
                    ' factor = factor * 2
                    ' Loop
                Else
                    For roomid = 1 To NumberRooms
                        If FloorConeDataFile(room) = "" Then FloorConeDataFile(room) = "null.txt"
                        Call Flame_Spread_Props(room, 3, floorhigh, FloorConeDataFile(room), Surface_Emissivity(4, room), ThermalInertiaFloor(room), IgTempF(room), FloorEffectiveHeatofCombustion(room), FloorHeatofGasification(room), AreaUnderFloorCurve(room), FloorFTP(room), FloorQCrit(room), Floorn(room))
                    Next roomid
                    'End If
                    points = UBound(ConeXF, 3)
                    datatobeplotted = ConeXF
                    flux = ExtFlux_F
                    curves = CurveNumber_F(room)
                    'Do While CurveNumber_F(room) * points / factor > 3800
                    ' factor = factor * 2
                    ' Loop
                End If
            Else

                For roomid = 1 To NumberRooms
                    If CeilingConeDataFile(room) = "" Then CeilingConeDataFile(room) = "null.txt"
                    Call Flame_Spread_Props(room, 2, ceilinghigh, CeilingConeDataFile(room), Surface_Emissivity(1, room), ThermalInertiaCeiling(room), IgTempC(room), CeilingEffectiveHeatofCombustion(room), CeilingHeatofGasification(room), AreaUnderCeilingCurve(room), CeilingFTP(room), CeilingQCrit(room), Ceilingn(room))
                Next roomid

                points = UBound(ConeXC, 3)
                datatobeplotted = ConeXC
                flux = ExtFlux_C
                curves = CurveNumber_C(room)
                'Do While CurveNumber_C(room) * points / factor > 3800
                ' factor = factor * 2
                ' Loop
            End If



            'With frmGraph.AxMSChart1
            '.chartType = MSChart20Lib.VtChChartType.VtChChartType2dXY
            '.Visible = True
            '.TitleText = Left(Description, 73) & " Room " & CStr(room)
            '.Plot.UniformAxis = False

            '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).AxisTitle.Text = "Time from Ignition (sec)"
            '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).AxisTitle.Text = YAxisTitle

            '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).ValueScale.Auto = True
            '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).ValueScale.Minimum = 0
            '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).ValueScale.Maximum = ymax
            '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).ValueScale.MajorDivision = 10

            '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).ValueScale.Auto = True
            '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).ValueScale.Minimum = 0
            '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).ValueScale.Maximum = Math.Ceiling(xdata.Max)
            '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).ValueScale.MajorDivision = CInt(0.1 * xdata.Max)

            '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).AxisTitle.TextLayout.Orientation = MSChart20Lib.VtOrientation.VtOrientationHorizontal

            '.Plot.SeriesCollection(1).SeriesMarker.Auto = True
            '.Plot.SeriesCollection(1).DataPoints(-1).Marker.Style = MSChart20Lib.VtMarkerStyle.VtMarkerStyleFilledCircle
            '.Plot.SeriesCollection(1).DataPoints(-1).Marker.Size = 150
            '.Plot.SeriesCollection(1).DataPoints(-1).Marker.Visible = True

            '.Plot.DataSeriesInRow = False
            '.Plot.SeriesCollection(1).ShowLine = True
            ' .Plot.SeriesCollection(1).StatLine.Flag = MSChart20Lib.VtChStats.VtChStatsRegression
            ' .Legend.Location.Visible = True

            '		On Error GoTo graphhandler

            Dim chdata(0 To points - 1, 0 To 2 * curves - 1) As Object
            curve = 1
            For i = 0 To 2 * curves - 1 Step 2

                chdata(0, i) = CStr(flux(curve, room)) & " kW/m2"

                For j = 1 To points - 1
                    chdata(j, i) = timedata(room, curve, j)
                    chdata(j, i + 1) = datatobeplotted(room, curve, j) * multiplier + DataShift 'data to be plotted
                    ' frmPlot.Chart1.Series("Room " & room).Points.AddXY(chdata(j, i), chdata(j, i + 1))
                Next

                curve = curve + 1
            Next i

            Dim ydata(0 To points) As Double

            frmPlot.Chart1.Series.Clear()
            'room = 1
            'For i = 0 To 2 * curves - 1 Step 2
            For i = 1 To curves
                'frmPlot.Chart1.Series.Add("1")

                frmPlot.Chart1.Series.Add(i - 1)

                'frmPlot.Chart1.Series("1").ChartType = SeriesChartType.FastLine
                frmPlot.Chart1.Series(i - 1).ChartType = SeriesChartType.FastLine
                frmPlot.Chart1.Series(i - 1).Name = CStr(flux(i, room)) & " kW/m2"

                For j = 1 To points - 1
                    ydata(j) = datatobeplotted(room, i, j) * multiplier + DataShift 'data to be plotted
                    'frmPlot.Chart1.Series("1").Points.AddXY(timedata(room, curve, j), ydata(j))
                    frmPlot.Chart1.Series(i - 1).Points.AddXY(timedata(room, i, j), ydata(j))
                Next

            Next i

            frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()


here:
            If YAxisTitle <> "Ceiling Cone HRR (kW/m2)" Then
                If YAxisTitle <> "Floor Cone HRR (kW/m2)" Then
                    For curve = 1 To CurveNumber_W(room)
                        If points = 0 Then Exit Sub

                    Next curve
                Else
                    For curve = 1 To CurveNumber_F(room)

                    Next curve
                End If
            Else
                For curve = 1 To CurveNumber_C(room)

                Next curve
            End If

            Exit Sub

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in Graph_Cone_Data")
        End Try

    End Sub

    Public Sub RunTime_ArrayTL(ByRef comboGraphIndex As Short, ByRef Graphdata(,) As Double, ByRef Shift As Single, ByRef multiplier As Single)
        '===================================================================
        '   returns the data array to be used in the top left run-time graph
        '===================================================================

        Static i As Integer
        Dim room As Short
        Dim j As Integer
        'ReDim Preserve Graphdata(1 To 10, 1 To stepcount) As Double
        ReDim Preserve Graphdata(MaxNumberRooms, stepcount)

        If stepcount < 2 Then Exit Sub
        If stepcount = 2 Then
            'room = frmOptions1.lstGraphRoom.ListIndex + 1
            i = 1
        End If
        j = i
        For room = 1 To NumberRooms
            i = j
            Select Case comboGraphIndex
                Case 1
                    While i <= stepcount
                        Graphdata(room, i) = layerheight(room, i)
                        Shift = 0
                        multiplier = 1
                        i = i + 1
                    End While
                Case 2
                    While i <= stepcount
                        Graphdata(room, i) = uppertemp(room, i)
                        Shift = -273
                        multiplier = 1
                        i = i + 1
                    End While
                Case 3
                    While i <= stepcount
                        Graphdata(room, i) = lowertemp(room, i)
                        Shift = -273
                        multiplier = 1
                        i = i + 1
                    End While
                Case 4
                    While i <= stepcount
                        Graphdata(room, i) = HeatRelease(room, i, 2)
                        Shift = 0
                        multiplier = 1
                        i = i + 1
                    End While
                Case 5
                    While i <= stepcount
                        'Graphdata(room, i) = massplumeflow(i, fireroom)
                        Graphdata(room, i) = massplumeflow(i, room)
                        Shift = 0
                        multiplier = 1
                        i = i + 1
                    End While
                Case 6
                    While i <= stepcount
                        Graphdata(room, i) = FlowToUpper(room, i)
                        Shift = 0
                        multiplier = 1
                        i = i + 1
                    End While
                Case 7
                    While i <= stepcount
                        Graphdata(room, i) = FlowToLower(room, i)
                        Shift = 0
                        multiplier = 1
                        i = i + 1
                    End While
                Case 8
                    While i <= stepcount
                        Graphdata(room, i) = UFlowToOutside(room, i)
                        Shift = 0
                        multiplier = 1
                        i = i + 1
                    End While
                Case 9
                    While i <= stepcount
                        Graphdata(room, i) = O2MassFraction(room, i, 1)
                        Shift = 0
                        multiplier = 100 * MolecularWeightAir / MolecularWeightO2
                        i = i + 1
                    End While
                Case 10
                    While i <= stepcount
                        Graphdata(room, i) = O2MassFraction(room, i, 2)
                        Shift = 0
                        multiplier = 100 * MolecularWeightAir / MolecularWeightO2
                        i = i + 1
                    End While
                Case 11
                    While i <= stepcount
                        Graphdata(room, i) = CO2MassFraction(room, i, 1)
                        Shift = 0
                        multiplier = 100 * MolecularWeightAir / MolecularWeightCO2
                        i = i + 1
                    End While
                Case 12
                    While i <= stepcount
                        Graphdata(room, i) = CO2MassFraction(room, i, 2)
                        Shift = 0
                        multiplier = 100 * MolecularWeightAir / MolecularWeightCO2
                        i = i + 1
                    End While
                Case 13
                    While i <= stepcount
                        Graphdata(room, i) = COMassFraction(room, i, 1)
                        Shift = 0
                        multiplier = 100 * MolecularWeightAir / MolecularWeightCO
                        i = i + 1
                    End While
                Case 14
                    While i <= stepcount
                        Graphdata(room, i) = COMassFraction(room, i, 2)
                        Shift = 0
                        multiplier = 100 * MolecularWeightAir / MolecularWeightCO
                        i = i + 1
                    End While
                Case 15
                    While i <= stepcount
                        Graphdata(room, i) = OD_upper(room, i)
                        Shift = 0
                        multiplier = 1
                        i = i + 1
                    End While
                Case 16
                    While i <= stepcount
                        Graphdata(room, i) = OD_lower(room, i)
                        Shift = 0
                        multiplier = 1
                        i = i + 1
                    End While
                Case 17
                    While i <= stepcount
                        Graphdata(room, i) = Visibility(room, i)
                        Shift = 0
                        multiplier = 1
                        i = i + 1
                    End While
                Case 18
                    While i <= stepcount
                        Graphdata(room, i) = RoomPressure(room, i)
                        Shift = 0
                        multiplier = 1
                        i = i + 1
                    End While
                Case 19
                    While i <= stepcount
                        Graphdata(room, i) = Yf_pyrolysis(room, i)
                        Shift = 0
                        multiplier = 1
                        i = i + 1
                    End While
                Case 20
                    While i <= stepcount
                        Graphdata(room, i) = Y_pyrolysis(room, i)
                        Shift = 0
                        multiplier = 1
                        i = i + 1
                    End While

            End Select
        Next room
    End Sub

    Public Function Get_Convection_Coefficient(ByRef temp As Double, ByRef ambient As Double) As Double
        '=====================================================================
        '   return a value for the convective heat transfer coefficient
        '   revised 28/12/97 Colleen Wade
        '=====================================================================

        Dim h As Double

        h = 5 + 0.45 * (temp - ambient)
        If h > 50 Then h = 50
        'h = 10

        Get_Convection_Coefficient = h

    End Function

    Public Function Get_Convection_Coefficient2(ByVal Area As Double, ByVal gastemp As Double, ByVal surfacetemp As Double, ByRef surface As String) As Double
        '=====================================================================
        '   return a value for the convective heat transfer coefficient
        '   revised 5/1/97 Colleen Wade
        '=====================================================================
        'Get_Convection_Coefficient2 = 0
        'Exit Function

        Dim k, Pr, h, GR, V, L As Double
        Dim constant As Single
        On Error GoTo errorhandler

        If CShort(surface) = VERTICAL Then
            constant = 0.13
        Else
            If gastemp > surfacetemp Then
                constant = 0.21
            Else
                constant = 0.012
            End If
        End If

        Pr = 0.72 'prandtl no

        V = 7.18 * 10 ^ (-10) * ((gastemp + surfacetemp) / 2) ^ (7 / 4)

        L = (Area) ^ 0.5

        GR = G * L ^ 3 * Abs(gastemp - surfacetemp) / (V ^ 2 * gastemp)

        k = 2.72 * 10 ^ (-4) * ((gastemp + surfacetemp) / 2) ^ 0.8

        If L <> 0 Then
            h = k / L * constant * (GR * Pr) ^ (1 / 3)
        Else
            h = 0
        End If

        Get_Convection_Coefficient2 = h
        Exit Function

errorhandler:
        Get_Convection_Coefficient2 = 10

    End Function

    Public Sub Graph_Data_Species(ByRef layer As Short, ByVal Title As String, ByVal datatobeplotted(,,) As Double, ByVal DataShift As Single, ByVal datamultiplier As Single, ByVal timeunit As Integer)
        '*  ====================================================================
        '*  This function takes data for a variable from a two-dimensional array
        '*  and displays it in a graph
        '*  ====================================================================
        Dim j As Integer
        Dim room As Integer
        Dim ydata(0 To NumberTimeSteps) As Double

        'if no data exists
        If NumberTimeSteps < 1 Then
            MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Sub
        End If
        Try
            If timeunit = 0 Then timeunit = 1
            'timeunit = 60 'show minutes
            'timeunit = 1 'show seconds

            frmPlot.Chart1.Series.Clear()
            room = 1
            For room = 1 To NumberRooms

                frmPlot.Chart1.Series.Add("Room " & room)

                frmPlot.Chart1.Series("Room " & room).ChartType = SeriesChartType.FastLine

                If Not roomcolor(room - 1).IsEmpty Then
                    frmPlot.Chart1.Series("Room " & room).Color = roomcolor(room - 1) '=line color
                End If

                For j = 1 To NumberTimeSteps
                    ydata(j) = (datamultiplier * datatobeplotted(room, j, layer) + DataShift) 'data to be plotted
                    frmPlot.Chart1.Series("Room " & room).Points.AddXY(tim(j, 1) / timeunit, ydata(j))
                Next

            Next room
            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.00"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            If timeunit = 60 Then
                frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (min)"
            Else
                frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            End If
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            'frmPlot.Chart1.Titles("Title1").Text = Title
            ' Disable X axis margin

            frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Graph_Data_Species")
        End Try
    End Sub
    Public Sub Graph_Data_fuelburningrate(ByRef layer As Short, ByVal Title As String, ByVal datatobeplotted(,,) As Double, ByVal DataShift As Single, ByVal datamultiplier As Single, ByVal timeunit As Integer)
        '*  ====================================================================
        '*  This function takes data for a variable from a three-dimensional array
        '*  and displays it in a graph
        '*  ====================================================================
        Dim j As Integer
        Dim room As Integer
        Dim ydata(0 To NumberTimeSteps) As Double

        'if no data exists
        If NumberTimeSteps < 1 Then
            MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Sub
        End If
        Try
            If timeunit = 0 Then timeunit = 1

            frmPlot.Chart1.Series.Clear()
            room = fireroom
            'For room = 1 To NumberRooms

            frmPlot.Chart1.Series.Add("Total mass loss rate ")
            frmPlot.Chart1.Series("Total mass loss rate ").ChartType = SeriesChartType.FastLine

            'If Not roomcolor(room - 1).IsEmpty Then
            'frmPlot.Chart1.Series("Total fuel burning rate ").Color = roomcolor(room - 1) '=line color
            frmPlot.Chart1.Series("Total mass loss rate ").Color = Color.Black '=line color
            frmPlot.Chart1.Series("Total mass loss rate ").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Series("Total mass loss rate ").BorderWidth = 2
            'End If

            For j = 1 To NumberTimeSteps
                ydata(j) = (datamultiplier * datatobeplotted(3, room, j) + DataShift) 'data to be plotted
                frmPlot.Chart1.Series("Total mass loss rate ").Points.AddXY(tim(j, 1) / timeunit, ydata(j))
            Next

            frmPlot.Chart1.Series.Add("free burn mass loss rate ")
            frmPlot.Chart1.Series("free burn mass loss rate ").ChartType = SeriesChartType.FastLine

            'If Not roomcolor(room - 1).IsEmpty Then
            'frmPlot.Chart1.Series("Total fuel burning rate ").Color = roomcolor(room - 1) '=line color
            frmPlot.Chart1.Series("free burn mass loss rate ").Color = Color.Blue '=line color
            frmPlot.Chart1.Series("free burn mass loss rate ").BorderDashStyle = ChartDashStyle.DashDotDot
            frmPlot.Chart1.Series("free burn mass loss rate ").BorderWidth = 1
            'End If

            For j = 1 To NumberTimeSteps
                ydata(j) = (datamultiplier * datatobeplotted(0, room, j) + DataShift) 'data to be plotted
                frmPlot.Chart1.Series("free burn mass loss rate ").Points.AddXY(tim(j, 1) / timeunit, ydata(j))
            Next

            frmPlot.Chart1.Series.Add("Ventilation effect ")
            frmPlot.Chart1.Series("Ventilation effect ").ChartType = SeriesChartType.FastLine

            'If Not roomcolor(room - 1).IsEmpty Then
            frmPlot.Chart1.Series("Ventilation effect ").Color = Color.Green '=line color
            frmPlot.Chart1.Series("Ventilation effect ").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Series("Ventilation effect ").BorderWidth = 1
            'End If

            For j = 1 To NumberTimeSteps
                ydata(j) = (datamultiplier * datatobeplotted(1, room, j) + DataShift) 'data to be plotted
                frmPlot.Chart1.Series("Ventilation effect ").Points.AddXY(tim(j, 1) / timeunit, ydata(j))
            Next

            frmPlot.Chart1.Series.Add("Thermal effect ")
            frmPlot.Chart1.Series("Thermal effect ").ChartType = SeriesChartType.FastLine

            'If Not roomcolor(room - 1).IsEmpty Then
            frmPlot.Chart1.Series("Thermal effect ").Color = Color.Red '=line color
            frmPlot.Chart1.Series("Thermal effect ").BorderDashStyle = ChartDashStyle.Dash
            frmPlot.Chart1.Series("Thermal effect ").BorderWidth = 1
            'End If

            For j = 1 To NumberTimeSteps
                ydata(j) = (datamultiplier * datatobeplotted(2, room, j) + DataShift) 'data to be plotted
                frmPlot.Chart1.Series("Thermal effect ").Points.AddXY(tim(j, 1) / timeunit, ydata(j))
            Next

            'Next room
            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1

            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.00"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            If timeunit = 60 Then
                frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (min)"
            Else
                frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            End If
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            'frmPlot.Chart1.Titles("Title1").Text = Title
            ' Disable X axis margin

            frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Graph_Data_fuelburningrate")
        End Try
    End Sub
    Public Sub Graph_Data_fuelburningrate2(ByRef layer As Short, ByVal Title As String, ByVal datatobeplotted(,,) As Double, ByVal DataShift As Single, ByVal datamultiplier As Single, ByVal timeunit As Integer)
        '*  ====================================================================
        '*  This function takes data for a variable from a three-dimensional array
        '*  and displays it in a graph
        '*  ====================================================================
        Dim j As Integer
        Dim room As Integer
        Dim ydata(0 To NumberTimeSteps) As Double

        'if no data exists
        If NumberTimeSteps < 1 Then
            MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Sub
        End If
        Try
            If timeunit = 0 Then timeunit = 1

            frmPlot.Chart1.Series.Clear()
            room = fireroom
            'For room = 1 To NumberRooms

            frmPlot.Chart1.Series.Add("Total mass loss rate ")
            frmPlot.Chart1.Series("Total mass loss rate ").ChartType = SeriesChartType.FastLine

            'If Not roomcolor(room - 1).IsEmpty Then
            'frmPlot.Chart1.Series("Total fuel burning rate ").Color = roomcolor(room - 1) '=line color
            frmPlot.Chart1.Series("Total mass loss rate ").Color = Color.Black '=line color
            frmPlot.Chart1.Series("Total mass loss rate ").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Series("Total mass loss rate ").BorderWidth = 2
            'End If

            For j = 1 To NumberTimeSteps
                ydata(j) = (datamultiplier * datatobeplotted(3, room, j) + DataShift) 'data to be plotted
                frmPlot.Chart1.Series("Total mass loss rate ").Points.AddXY(tim(j, 1) / timeunit, ydata(j))
            Next

            frmPlot.Chart1.Series.Add("Burning rate ")
            frmPlot.Chart1.Series("Burning rate ").ChartType = SeriesChartType.FastLine

            'If Not roomcolor(room - 1).IsEmpty Then
            'frmPlot.Chart1.Series("Total fuel burning rate ").Color = roomcolor(room - 1) '=line color
            frmPlot.Chart1.Series("Burning rate ").Color = Color.Blue '=line color
            frmPlot.Chart1.Series("Burning rate ").BorderDashStyle = ChartDashStyle.DashDotDot
            frmPlot.Chart1.Series("Burning rate ").BorderWidth = 1
            'End If

            For j = 1 To NumberTimeSteps
                ydata(j) = (datamultiplier * datatobeplotted(4, room, j) + DataShift) 'data to be plotted
                frmPlot.Chart1.Series("Burning rate ").Points.AddXY(tim(j, 1) / timeunit, ydata(j))
            Next

            'Next room
            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1

            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.00"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            If timeunit = 60 Then
                frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (min)"
            Else
                frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            End If
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            'frmPlot.Chart1.Titles("Title1").Text = Title
            ' Disable X axis margin

            frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Graph_Data_fuelburningrate")
        End Try
    End Sub
    Public Sub Graph_Data_areashrinkage(ByRef layer As Short, ByVal Title As String, ByVal datatobeplotted(,,) As Double, ByVal DataShift As Single, ByVal datamultiplier As Single, ByVal timeunit As Integer)
        '*  ====================================================================
        '*  This function takes data for a variable from a three-dimensional array
        '*  and displays it in a graph
        '*  ====================================================================
        Dim j As Integer
        Dim room As Integer
        Dim ydata(0 To NumberTimeSteps) As Double

        'if no data exists
        If NumberTimeSteps < 1 Then
            MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Sub
        End If
        Try
            If timeunit = 0 Then timeunit = 1

            frmPlot.Chart1.Series.Clear()
            room = fireroom
            'For room = 1 To NumberRooms

            frmPlot.Chart1.Series.Add("Area shrinkage ratio ")
            frmPlot.Chart1.Series("Area shrinkage ratio ").ChartType = SeriesChartType.FastLine

            'If Not roomcolor(room - 1).IsEmpty Then
            'frmPlot.Chart1.Series("Total fuel burning rate ").Color = roomcolor(room - 1) '=line color
            frmPlot.Chart1.Series("Area shrinkage ratio ").Color = Color.Black '=line color
            frmPlot.Chart1.Series("Area shrinkage ratio ").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Series("Area shrinkage ratio ").BorderWidth = 2
            'End If

            For j = 1 To NumberTimeSteps
                ydata(j) = (datamultiplier * datatobeplotted(5, room, j) + DataShift) 'data to be plotted
                frmPlot.Chart1.Series("Area shrinkage ratio ").Points.AddXY(tim(j, 1) / timeunit, ydata(j))
            Next


            'Next room
            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1

            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.00"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            If timeunit = 60 Then
                frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (min)"
            Else
                frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            End If
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            'frmPlot.Chart1.Titles("Title1").Text = Title
            ' Disable X axis margin

            frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Graph_Data_areashrinkage")
        End Try
    End Sub

    Public Sub Graph_Data_Runtime(ByRef layer As Short, ByVal Title As String, ByVal datatobeplotted(,,) As Double, ByVal DataShift As Single, ByVal datamultiplier As Single)
        '*  ====================================================================
        '*  This function takes data for a variable from a three-dimensional array
        '*  and displays it in a graph
        '*  ====================================================================

        Dim room As Integer
        Dim ydata(0 To NumberTimeSteps) As Double

        'if no data exists
        If NumberTimeSteps < 1 Then
            'MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Sub
        End If
        Try

            If stepcount = 1 Then
                MDIFrmMain.ChartRuntime1.Dock = DockStyle.None
                MDIFrmMain.ChartRuntime1.BackColor = Color.AliceBlue
                MDIFrmMain.ChartRuntime1.ChartAreas("ChartArea1").BorderWidth = 1
                MDIFrmMain.ChartRuntime1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
                MDIFrmMain.ChartRuntime1.ChartAreas("ChartArea1").AxisY.Title = Title
                MDIFrmMain.ChartRuntime1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
                MDIFrmMain.ChartRuntime1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
                MDIFrmMain.ChartRuntime1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
                MDIFrmMain.ChartRuntime1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
                MDIFrmMain.ChartRuntime1.Legends("Legend1").BorderWidth = 1
                MDIFrmMain.ChartRuntime1.Legends("Legend1").BackColor = Color.White
                MDIFrmMain.ChartRuntime1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
                MDIFrmMain.ChartRuntime1.Legends("Legend1").Docking = Docking.Right

                MDIFrmMain.ChartRuntime1.Visible = True
                MDIFrmMain.ChartRuntime1.Series.Clear()

                For room = 1 To NumberRooms

                    MDIFrmMain.ChartRuntime1.Series.Add("Room " & room)

                    MDIFrmMain.ChartRuntime1.Series("Room " & room).ChartType = SeriesChartType.FastLine

                Next room

            End If

            For room = 1 To NumberRooms

                ydata(stepcount - 1) = (datamultiplier * datatobeplotted(room, stepcount - 1, layer) + DataShift) 'data to be plotted
                MDIFrmMain.ChartRuntime1.Series("Room " & room).Points.AddXY(tim(stepcount - 1, 1), ydata(stepcount - 1))

            Next room

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Graph_Data_Runtime")
        End Try
    End Sub
    Public Sub Graph_Data_Runtime2(ByRef chart As Object, ByVal Title As String, ByVal datatobeplotted(,) As Double, ByVal DataShift As Single, ByVal datamultiplier As Single)
        '*  ====================================================================
        '*  This function takes data for a variable from a three-dimensional array
        '*  and displays it in a graph
        '*  ====================================================================

        Dim room As Integer
        Dim ydata(0 To NumberTimeSteps) As Double

        'if no data exists
        If NumberTimeSteps < 1 Then
            'MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Sub
        End If
        Try

            If stepcount = 1 Then
                chart.Dock = DockStyle.None
                chart.BackColor = Color.AliceBlue
                chart.ChartAreas("ChartArea1").BorderWidth = 1
                chart.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
                chart.ChartAreas("ChartArea1").AxisY.Title = Title
                chart.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
                chart.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
                chart.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
                chart.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
                chart.Legends("Legend1").BorderWidth = 1
                chart.Legends("Legend1").BackColor = Color.White
                chart.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
                chart.Legends("Legend1").Docking = Docking.Right

                chart.Visible = True
                chart.Series.Clear()

                For room = 1 To NumberRooms

                    chart.Series.Add("Room " & room)

                    chart.Series("Room " & room).ChartType = SeriesChartType.FastLine

                Next room

            End If

            If stepcount <= NumberTimeSteps Then
                For room = 1 To NumberRooms

                    ydata(stepcount) = (datamultiplier * datatobeplotted(room, stepcount) + DataShift) 'data to be plotted
                    chart.Series("Room " & room).Points.AddXY(tim(stepcount, 1), ydata(stepcount))

                Next room
            End If
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Graph_Data_Runtime2")
        End Try
    End Sub
    Public Sub Open_File(ByRef opendatafile As String, ByVal batch As Boolean)
        '*  ===================================================================
        '*  Open an existing BRANZFIRE model.
        '*  ===================================================================

        Dim myfile As String
        Dim k, i, Action, plume As Short
        Dim SuppressCeiling, HCNcalc As Boolean
        Dim dummy As String = ""
        Dim room As Short
        ' Dim illumination, post As Boolean
        Dim post As Boolean
        Dim ver As String = ""
        Dim usefan As Boolean

        Dim matname, firename As String

        Try


            If opendatafile <> "" Then

                'returns filename without path if it exists
                myfile = Path.GetFileName(opendatafile)

                If myfile <> "" Then

                    FileOpen(1, opendatafile, OpenMode.Input)
                    New_File()

                    Input(1, dummy)
                    Input(1, ver)
                    Input(1, Description)
                    Input(1, dummy)
                    Input(1, corner)
                    Input(1, dummy)
                    Input(1, NumberRooms)
                    Resize_Rooms()
                    For room = 1 To NumberRooms
                        Input(1, dummy)
                        Input(1, room)
                        Input(1, dummy)
                        Input(1, RoomWidth(room))
                        Input(1, dummy)
                        Input(1, RoomLength(room))
                        Input(1, dummy)
                        Input(1, RoomDescription(room))
                        Input(1, dummy)
                        Input(1, RoomHeight(room))
                        Input(1, dummy)
                        Input(1, MinStudHeight(room))
                        Input(1, dummy)
                        Input(1, FloorElevation(room))
                        Input(1, dummy)
                        Input(1, WallSurface(room))
                        Input(1, dummy)
                        Input(1, WallSubstrate(room))
                        Input(1, dummy)
                        Input(1, CeilingSurface(room))
                        Input(1, dummy)
                        Input(1, CeilingSubstrate(room))
                        If CSng(ver) >= 2001.5 Then Input(1, dummy)
                        Input(1, FloorSubstrate(room))
                        Input(1, dummy)
                        Input(1, WallThickness(room))
                        Input(1, dummy)
                        Input(1, CeilingThickness(room))
                        Input(1, dummy)
                        Input(1, FloorThickness(room))
                        Input(1, dummy)
                        Input(1, WallConductivity(room))
                        Input(1, dummy)
                        Input(1, CeilingConductivity(room))
                        Input(1, dummy)
                        Input(1, FloorConductivity(room))
                        Input(1, dummy)
                        Input(1, FloorSurface(room))
                        Input(1, dummy)
                        Input(1, WallSpecificHeat(room))
                        Input(1, dummy)
                        Input(1, WallDensity(room))
                        Input(1, dummy)
                        Input(1, WallSubThickness(room))
                        Input(1, dummy)
                        Input(1, CeilingSubThickness(room))
                        If CSng(ver) >= 2001.5 Then Input(1, dummy)
                        Input(1, FloorSubThickness(room))
                        Input(1, dummy)
                        Input(1, WallSubConductivity(room))
                        If CSng(ver) >= 2001.5 Then Input(1, dummy)
                        Input(1, FloorSubConductivity(room))
                        Input(1, dummy)
                        Input(1, WallSubSpecificHeat(room))
                        If CSng(ver) >= 2001.5 Then Input(1, dummy)
                        Input(1, FloorSubSpecificHeat(room))
                        Input(1, dummy)
                        Input(1, WallSubDensity(room))
                        If CSng(ver) >= 2001.5 Then Input(1, dummy)
                        Input(1, FloorSubDensity(room))
                        Input(1, dummy)
                        Input(1, FloorSpecificHeat(room))
                        Input(1, dummy)
                        Input(1, FloorDensity(room))
                        Input(1, dummy)
                        Input(1, CeilingSpecificHeat(room))
                        Input(1, dummy)
                        Input(1, CeilingDensity(room))
                        Input(1, dummy)
                        Input(1, CeilingSubConductivity(room))
                        Input(1, dummy)
                        Input(1, CeilingSubSpecificHeat(room))
                        Input(1, dummy)
                        Input(1, CeilingSubDensity(room))
                        Input(1, dummy)
                        Input(1, HaveCeilingSubstrate(room))
                        Input(1, dummy)
                        Input(1, HaveWallSubstrate(room))
                        If CSng(ver) >= 2001.5 Then Input(1, dummy)
                        Input(1, HaveFloorSubstrate(room))
                        'If frmDescribeRoom.optFlatCeiling.Value = True Then CeilingSlope(room) = False
                        Input(1, dummy)
                        Input(1, CeilingSlope(room))
                        Input(1, dummy)
                        Input(1, Surface_Emissivity(1, room))
                        Input(1, dummy)
                        Input(1, Surface_Emissivity(2, room))
                        Input(1, dummy)
                        Input(1, Surface_Emissivity(3, room))
                        Input(1, dummy)
                        Input(1, Surface_Emissivity(4, room))
                    Next room

                    Input(1, dummy)
                    Input(1, InteriorTemp)
                    Input(1, dummy)
                    Input(1, ExteriorTemp)
                    Input(1, dummy)
                    Input(1, RelativeHumidity)
                    Input(1, dummy)
                    Input(1, MonitorHeight)
                    Input(1, dummy)
                    Input(1, Activity)
                    Input(1, dummy)
                    Input(1, dummy)
                    Input(1, dummy)
                    Input(1, dummy)
                    'Input(1, MLUnitArea)
                    Input(1, dummy)
                    Input(1, EmissionCoefficient)
                    Input(1, dummy)
                    Input(1, SimTime)
                    Input(1, dummy)
                    Input(1, DisplayInterval)
                    Input(1, dummy)
                    Input(1, plume)
                    Input(1, dummy)
                    Input(1, SuppressCeiling)
                    Input(1, dummy)
                    Input(1, FlameAreaConstant)
                    Input(1, dummy)
                    Input(1, FlameLengthPower)
                    Input(1, dummy)
                    Input(1, BurnerWidth)
                    Input(1, dummy)
                    Input(1, WallHeatFlux)
                    Input(1, dummy)
                    Input(1, CeilingHeatFlux)
                    For room = 1 To NumberRooms
                        For i = 2 To NumberRooms + 1

                            If room < i Then
                                Input(1, dummy)
                                Input(1, NumberVents(room, i))
                            End If

                            NumberVents(i, room) = NumberVents(room, i)

                        Next i
                    Next room
                    Resize_Vents()

                    For room = 1 To NumberRooms
                        For i = 2 To NumberRooms + 1
                            If room < i Then
                                For k = 1 To NumberVents(room, i)
                                    If NumberVents(room, i) > 0 Then
                                        Input(1, dummy)
                                        Input(1, dummy)
                                        Input(1, dummy)
                                        Input(1, dummy)
                                        Input(1, dummy)
                                        Input(1, dummy)
                                        Input(1, dummy)
                                        Input(1, VentHeight(room, i, k))
                                        Input(1, dummy)
                                        Input(1, VentWidth(room, i, k))
                                        Input(1, dummy)
                                        Input(1, VentSillHeight(room, i, k))
                                        Input(1, dummy)
                                        Input(1, VentOpenTime(room, i, k))
                                        Input(1, dummy)
                                        Input(1, VentCloseTime(room, i, k))
                                        If CSng(ver) >= 2001.5 Then
                                            Input(1, dummy)
                                            Input(1, GLASSconductivity(room, i, k))
                                            Input(1, dummy)
                                            Input(1, GLASSemissivity(room, i, k))
                                            Input(1, dummy)
                                            Input(1, GLASSexpansion(room, i, k))
                                            Input(1, dummy)
                                            Input(1, GLASSthickness(room, i, k))
                                            Input(1, dummy)
                                            Input(1, GLASSshading(room, i, k))
                                            Input(1, dummy)
                                            Input(1, GLASSbreakingstress(room, i, k))
                                            Input(1, dummy)
                                            Input(1, GLASSalpha(room, i, k))
                                            Input(1, dummy)
                                            Input(1, GlassYoungsModulus(room, i, k))
                                            Input(1, dummy)
                                            Input(1, AutoBreakGlass(room, i, k))
                                        End If
                                        If CSng(ver) >= 2001.7 Then
                                            Input(1, dummy)
                                            Input(1, GLASSFalloutTime(room, i, k))
                                            Input(1, dummy)
                                            Input(1, GLASSdistance(room, i, k))
                                            Input(1, dummy)
                                            Input(1, GlassFlameFlux(room, i, k))
                                        End If
                                        If CSng(ver) >= 2004.2 Then
                                            Input(1, dummy)
                                            Input(1, Downstand(room, i, k))
                                            Input(1, dummy)
                                            Input(1, SpillPlumeBalc(room, i, k))
                                        End If
                                        If CSng(ver) >= 2002.89 Then
                                            Input(1, dummy)
                                            Input(1, spillplume(room, i, k))
                                        End If
                                        If CSng(ver) >= 2005.1 Then
                                            Input(1, dummy)
                                            Input(1, spillplumemodel(room, i, k))
                                        End If
                                        If CSng(ver) >= CSng(2005.11) Then
                                            Input(1, dummy)
                                            'Input(1, spillbalconyprojection(room, i, k))
                                            Input(1, dummy)
                                        End If
                                        If i <> NumberRooms + 1 Then
                                            Input(1, dummy)
                                            Input(1, WallLength1(room, i, k))
                                            Input(1, dummy)
                                            Input(1, WallLength2(room, i, k))
                                        End If

                                        spillplume(i, room, k) = spillplume(room, i, k)
                                        spillplumemodel(i, room, k) = spillplumemodel(room, i, k)
                                        spillbalconyprojection(i, room, k) = spillbalconyprojection(room, i, k)
                                        VentHeight(i, room, k) = VentHeight(room, i, k)
                                        VentWidth(i, room, k) = VentWidth(room, i, k)
                                        If i <= NumberRooms Then VentSillHeight(i, room, k) = VentSillHeight(room, i, k) + FloorElevation(room) - FloorElevation(i)
                                        VentOpenTime(i, room, k) = VentOpenTime(room, i, k)
                                        VentCloseTime(i, room, k) = VentCloseTime(room, i, k)
                                        Downstand(i, room, k) = Downstand(room, i, k)
                                        GLASSconductivity(i, room, k) = GLASSconductivity(room, i, k)
                                        GLASSemissivity(i, room, k) = GLASSemissivity(room, i, k)
                                        GLASSexpansion(i, room, k) = GLASSexpansion(room, i, k)
                                        AutoBreakGlass(i, room, k) = AutoBreakGlass(room, i, k)
                                        GLASSthickness(i, room, k) = GLASSthickness(room, i, k)
                                        GLASSFalloutTime(i, room, k) = GLASSFalloutTime(room, i, k)
                                        GLASSshading(i, room, k) = GLASSshading(room, i, k)
                                        GLASSbreakingstress(i, room, k) = GLASSbreakingstress(room, i, k)
                                        GLASSalpha(i, room, k) = GLASSalpha(room, i, k)
                                        GlassYoungsModulus(i, room, k) = GlassYoungsModulus(room, i, k)

                                        '2009.20 include dbh  vent opening options - not in this version yet
                                        If CSng(ver) >= CSng(2009.2) Then
                                            Input(1, dummy)
                                            Input(1, AutoOpenVent(room, i, k))
                                            Input(1, dummy)
                                            Input(1, HRR_threshold(room, i, k))
                                            Input(1, dummy)
                                            Input(1, HRR_threshold_time(room, i, k))
                                            Input(1, dummy)
                                            Input(1, HRRFlag(room, i, k))
                                            Input(1, dummy)
                                            Input(1, trigger_ventopendelay(room, i, k))
                                            Input(1, dummy)
                                            Input(1, trigger_ventopenduration(room, i, k))
                                            Input(1, dummy)
                                            Input(1, HRR_ventopendelay(room, i, k))
                                            Input(1, dummy)
                                            Input(1, HRR_ventopenduration(room, i, k))
                                            Input(1, dummy)
                                            Input(1, SDtriggerroom(room, i, k))
                                            Input(1, dummy)
                                            Input(1, trigger_device(0, room, i, k))
                                            Input(1, dummy)
                                            Input(1, trigger_device(1, room, i, k))
                                            Input(1, dummy)
                                            Input(1, trigger_device(2, room, i, k))
                                            If CSng(ver) > CSng(2012.19) Then
                                                Input(1, dummy)
                                                Input(1, trigger_device(3, room, i, k))
                                            End If
                                            If CSng(ver) > CSng(2012.2) Then
                                                Input(1, dummy)
                                                Input(1, trigger_device(4, room, i, k))
                                            End If
                                            If CSng(ver) > CSng(2012.38) Then
                                                Input(1, dummy)
                                                Input(1, trigger_device(5, room, i, k))
                                            End If
                                            If CSng(ver) > CSng(2013.08) Then
                                                Input(1, dummy)
                                                Input(1, trigger_device(6, room, i, k))
                                            End If
                                            AutoOpenVent(i, room, k) = AutoOpenVent(room, i, k)
                                            HRR_threshold(i, room, k) = HRR_threshold(room, i, k)
                                            HRR_threshold_time(i, room, k) = HRR_threshold_time(room, i, k)
                                            HRRFlag(i, room, k) = HRRFlag(room, i, k)
                                            trigger_ventopendelay(i, room, k) = trigger_ventopendelay(room, i, k)
                                            trigger_ventopenduration(i, room, k) = trigger_ventopenduration(room, i, k)
                                            HRR_ventopendelay(i, room, k) = HRR_ventopendelay(room, i, k)
                                            HRR_ventopenduration(i, room, k) = HRR_ventopenduration(room, i, k)
                                            SDtriggerroom(i, room, k) = SDtriggerroom(room, i, k)
                                            trigger_device(0, i, room, k) = trigger_device(0, room, i, k)
                                            trigger_device(1, i, room, k) = trigger_device(1, room, i, k)
                                            trigger_device(2, i, room, k) = trigger_device(2, room, i, k)
                                            trigger_device(3, i, room, k) = trigger_device(3, room, i, k)
                                            trigger_device(4, i, room, k) = trigger_device(4, room, i, k)
                                            trigger_device(5, i, room, k) = trigger_device(5, room, i, k)
                                            trigger_device(6, i, room, k) = trigger_device(6, room, i, k)

                                        End If
                                    End If
                                Next k
                            End If
                        Next i
                    Next room

                    Input(1, dummy)

                    Input(1, NumberObjects)

                    Resize_Objects()

                    For k = 1 To NumberObjects
                        Input(1, dummy)
                        Input(1, NumberDataPoints(k))
                        Input(1, dummy)
                        Input(1, EnergyYield(k))
                        Input(1, dummy)
                        Input(1, COYield)
                        Input(1, dummy)
                        Input(1, CO2Yield(k))
                        Input(1, dummy)
                        Input(1, SootYield(k))
                        Input(1, dummy)
                        Input(1, WaterVaporYield(k))
                        Input(1, dummy)
                        Input(1, FireHeight(k))
                        Input(1, dummy)
                        Input(1, FireLocation(k))

                        'If CSng(ver) >= 2009.21 Then
                        '    Input(1, HCNuserYield(k))
                        '    Input(1, ObjCRF(k))
                        '    Input(1, ObjLength(k))
                        '    Input(1, ObjWidth(k))
                        '    Input(1, ObjHeight(k))
                        '    Input(1, ObjElevation(k))
                        'End If

                        Input(1, dummy)
                        For i = 1 To NumberDataPoints(k)
                            If i < 501 Then Input(1, HeatReleaseData(1, i, k))
                            Input(1, HeatReleaseData(2, i, k))
                        Next i
                    Next k
                    Input(1, dummy)
                    Input(1, dummy)
                    Input(1, dummy)
                    Input(1, dummy)
                    Input(1, dummy)
                    Input(1, dummy)
                    Input(1, dummy)
                    Input(1, dummy)
                    Input(1, dummy)
                    Input(1, dummy)
                    Input(1, dummy)
                    Input(1, dummy)
                    Input(1, dummy)
                    Input(1, Sprinkler(1, fireroom))
                    Input(1, Sprinkler(2, fireroom))
                    Input(1, Sprinkler(3, fireroom))
                    Input(1, dummy)
                    Input(1, TargetEndPoint)
                    Input(1, dummy)
                    Input(1, TempEndPoint)
                    Input(1, dummy)
                    Input(1, VisibilityEndPoint)
                    Input(1, dummy)
                    Input(1, FEDEndPoint)
                    Input(1, dummy)
                    Input(1, ConvectEndPoint)
                    For room = 1 To NumberRooms
                        Input(1, WallConeDataFile(room))
                        If CSng(ver) >= 2001.5 Then Input(1, FloorConeDataFile(room))
                        Input(1, CeilingConeDataFile(room))
                        Input(1, dummy)
                        Input(1, WallTSMin(room))
                        Input(1, dummy)
                        Input(1, WallFlameSpreadParameter(room))
                        Input(1, dummy)
                        Input(1, WallEffectiveHeatofCombustion(room))
                        Input(1, dummy)
                        Input(1, CeilingEffectiveHeatofCombustion(room))
                        If CSng(ver) >= 2001.5 Then Input(1, dummy)
                        Input(1, FloorEffectiveHeatofCombustion(room))
                        Input(1, dummy)
                        Input(1, ExtractRate(room))
                        Input(1, dummy)
                        Input(1, ExtractStartTime(room))
                        Input(1, dummy)
                        Input(1, fanon(room))
                        Input(1, dummy)
                        Input(1, MaxPressure(room))
                        Input(1, dummy)
                        Input(1, Extract(room))
                        If CSng(ver) >= 2003.1 Then Input(1, dummy)
                        Input(1, NumberFans(room))
                        Input(1, dummy)
                        Input(1, WallSootYield(room))
                        Input(1, dummy)
                        Input(1, CeilingSootYield(room))
                        If CSng(ver) >= 2001.5 Then Input(1, dummy)
                        Input(1, FloorSootYield(room))
                        Input(1, dummy)
                        Input(1, WallCO2Yield(room))
                        Input(1, dummy)
                        Input(1, CeilingCO2Yield(room))
                        If CSng(ver) >= 2001.5 Then Input(1, dummy)
                        Input(1, FloorCO2Yield(room))
                        Input(1, dummy)
                        Input(1, WallH2OYield(room))
                        Input(1, dummy)
                        Input(1, CeilingH2OYield(room))
                        If CSng(ver) >= 2001.5 Then Input(1, dummy)
                        Input(1, FloorH2OYield(room))
                        If CSng(ver) >= 2002.4 Then
                            Input(1, dummy)
                            Input(1, FloorTSMin(room))
                            Input(1, dummy)
                            Input(1, FloorFlameSpreadParameter(room))
                        End If
                        If CSng(ver) >= 2003.1 Then
                            'Input #1, dummy, FloorHCNYield(room)
                            'Input #1, dummy, WallHCNYield(room)
                            'Input #1, dummy, CeilingHCNYield(room)
                        End If
                    Next room
                    Input(1, dummy)
                    Input(1, fireroom)
                    DetectorType(fireroom) = DetectorType(1)

                    'cfactor(fireroom) = cfactor(1)
                    'RadialDistance(fireroom) = RadialDistance(1)
                    'ActuationTemp(fireroom) = ActuationTemp(1)
                    'WaterSprayDensity(fireroom) = WaterSprayDensity(1)
                    Sprinkler(1, fireroom) = Sprinkler(1, 1)
                    Sprinkler(2, fireroom) = Sprinkler(2, 1)
                    Sprinkler(3, fireroom) = Sprinkler(3, 1)
                    'frmSprinklerList.optControl.Checked = Sprinkler(1, fireroom)
                    'frmSprinklerList.optSprinkOff.Checked = Sprinkler(2, fireroom)
                    'frmSprinklerList.optSuppression.Checked = Sprinkler(3, fireroom)
                    'frmSprink.optControl.Checked = Sprinkler(1, fireroom)
                    'frmSprink.optSprinkOff.Checked = Sprinkler(2, fireroom)
                    'frmSprink.optSuppression.Checked = Sprinkler(3, fireroom)
                    If NumberRooms > 1 Then
                        Input(1, dummy)
                        Input(1, dummy)
                        dummy = Strings.Replace(dummy, "#", "")
                        IgniteNextRoom = Convert.ToBoolean(dummy)
                    End If

                    Input(1, dummy)
                    Input(1, StartOccupied)
                    Input(1, dummy)
                    Input(1, EndOccupied)
                    Input(1, dummy)
                    Input(1, illumination)
                    If CSng(ver) >= 2000.01 Then
                        For room = 1 To NumberRooms + 1
                            For i = 1 To NumberRooms + 1
                                If room <> i Then
                                    Input(1, dummy)
                                    Input(1, dummy)
                                    NumberCVents(room, i) = CShort(dummy)
                                End If
                            Next i
                        Next room
                        Resize_CVents()

                        For room = 1 To NumberRooms + 1
                            For i = 1 To NumberRooms + 1
                                If room <> i Then
                                    For k = 1 To NumberCVents(room, i)
                                        If NumberCVents(room, i) > 0 Then
                                            Input(1, dummy)
                                            Input(1, dummy)
                                            Input(1, dummy)
                                            Input(1, dummy)
                                            Input(1, dummy)
                                            Input(1, dummy)
                                            Input(1, dummy)
                                            Input(1, CVentArea(room, i, k))
                                            Input(1, dummy)
                                            Input(1, CVentOpenTime(room, i, k))
                                            Input(1, dummy)
                                            Input(1, CVentCloseTime(room, i, k))
                                            Input(1, dummy)
                                            Input(1, CVentAuto(room, i, k))
                                        End If
                                    Next k
                                End If
                            Next i
                        Next room
                        For room = 1 To NumberRooms
                            Input(1, dummy)
                            Input(1, dummy)
                            dummy = Strings.Replace(dummy, "#", "")
                            usefan = Convert.ToBoolean(dummy)

                            UseFanCurve(room) = usefan
                            Input(1, dummy)
                            Input(1, FanElevation(room))
                        Next room
                    End If
                    If CSng(ver) >= 2000.02 Then
                        Input(1, dummy)
                        Input(1, Ceilingnodes)
                        Input(1, dummy)
                        Input(1, Wallnodes)
                        Input(1, dummy)
                        Input(1, Floornodes)
                        Input(1, dummy)
                        Input(1, LEsolver)
                        Input(1, dummy)
                        Input(1, Enhance)
                        Input(1, dummy)
                        Input(1, JobNumber)
                        Input(1, dummy)
                        Input(1, ExcelInterval)
                        For room = 1 To NumberRooms
                            Input(1, dummy)
                            Input(1, TwoZones(room))
                        Next room
                        Input(1, dummy)
                        Input(1, Timestep)

                        Input(1, dummy)
                        Input(1, Error_Control)

                        Input(1, dummy)
                        Input(1, FireDatabaseName)
                        Input(1, dummy)
                        Input(1, MaterialsDatabaseName)

                        For room = 1 To NumberRooms
                            Input(1, dummy)
                            Input(1, HaveSD(room))
                            Input(1, dummy)
                            Input(1, SmokeOD(room))
                            Input(1, dummy)
                            Input(1, SDdelay(room))
                            If CSng(ver) >= 2005.1 Then
                                Input(1, dummy)
                                Input(1, DetSensitivity(room))
                            End If
                            If CSng(ver) >= 2003.21 Then
                                Input(1, dummy)
                                Input(1, SDRadialDist(room))
                                Input(1, dummy)
                                Input(1, SDdepth(room))
                                Input(1, dummy)
                                Input(1, SDinside(room))
                            End If
                            Input(1, dummy)
                            Input(1, FanAutoStart(room))
                            Input(1, dummy)
                            Input(1, SpecifyOD(room))
                        Next room
                        Input(1, dummy)
                        Input(1, cjModel)
                    End If
                    If CSng(ver) >= 2000.03 Then
                        Input(1, dummy)
                        Input(1, UseOneCurve)
                        Input(1, dummy)
                        Input(1, IgnCorrelation)
                    End If
                    If CSng(ver) >= 2000.04 Then
                        Input(1, dummy)
                        Input(1, SprinkDistance(fireroom))
                    End If
                    If CSng(ver) >= 2002.1 Then
                        Input(1, dummy)
                        Input(1, ventlog)
                        Input(1, dummy)
                        Input(1, SootFactor)
                    End If
                    If CSng(ver) >= 2002.2 Then
                        Input(1, dummy)
                        Input(1, post)
                        Input(1, dummy)
                        Input(1, FLED)
                        Input(1, dummy)
                        Input(1, dummy)
                        Input(1, dummy)
                        Input(1, Fuel_Thickness)
                        Input(1, dummy)
                        Input(1, HoC_fuel)
                    End If
                    If CSng(ver) > 2002.7 Then
                        Input(1, dummy)
                        Input(1, Stick_Spacing)
                        Input(1, dummy)
                        Input(1, SootAlpha)
                        Input(1, dummy)
                        Input(1, SootEpsilon)
                    End If
                    If CSng(ver) >= 2003.1 Then
                        Input(1, dummy)
                        Input(1, nC)
                        Input(1, dummy)
                        Input(1, nH)
                        Input(1, dummy)
                        Input(1, nO)
                        Input(1, dummy)
                        Input(1, nN)
                        Input(1, dummy)
                        Input(1, fueltype)
                        Input(1, dummy)
                        Input(1, nowallflow)
                        Input(1, dummy)
                        Input(1, HCNcalc)
                    End If
                    If CSng(ver) >= 2007.2 Then
                        Input(1, dummy)
                        Input(1, preCO)
                        Input(1, dummy)
                        Input(1, postCO)
                        Input(1, dummy)
                        Input(1, preSoot)
                        Input(1, dummy)
                        Input(1, postSoot)
                        Input(1, dummy)
                        Input(1, comode)
                        Input(1, dummy)
                        Input(1, sootmode)
                    End If
                    'If CSng(ver) >= 2010.3 Then
                    '    Input(1, dummy)
                    '    Input(1, MDIFrmMain.NZBCVM2ToolStripMenuItem.Checked)
                    'End If
                    FileClose(1)

                    If post = True Then
                        frmOptions1.optPostFlashover.Checked = True
                    Else
                        frmOptions1.OptPreFlashover.Checked = True
                    End If
                    'frmOptions1.txtFLED.Text = CStr(FLED)

                    frmOptions1.txtFuelThickness.Text = CStr(Fuel_Thickness)
                    frmOptions1.txtStickSpacing.Text = CStr(Stick_Spacing)

                    'frmOptions1.txtHOCFuel.Text = CStr(HoC_fuel)
                    frmOptions1.txtnC.Text = CStr(nC)
                    frmOptions1.txtnH.Text = CStr(nH)
                    frmOptions1.txtnO.Text = CStr(nO)
                    frmOptions1.txtnN.Text = CStr(nN)
                    frmOptions1.txtStoich.Text = CStr(StoichAFratio)
                    frmOptions1.cboABSCoeff.Text = fueltype
                    If nowallflow = True Then
                        frmOptions1.chkWallFlowDisable.CheckState = System.Windows.Forms.CheckState.Checked
                    Else
                        frmOptions1.chkWallFlowDisable.CheckState = System.Windows.Forms.CheckState.Unchecked
                    End If
                    If HCNcalc = True Then
                        frmOptions1.chkHCNcalc.CheckState = System.Windows.Forms.CheckState.Checked
                    Else
                        frmOptions1.chkHCNcalc.CheckState = System.Windows.Forms.CheckState.Unchecked
                    End If
                    If ventlog = True Then
                        frmInputs.chkSaveVentFlow.CheckState = System.Windows.Forms.CheckState.Checked
                    Else
                        frmInputs.chkSaveVentFlow.CheckState = System.Windows.Forms.CheckState.Unchecked
                    End If
                    frmQuintiere.optUseOneCurve.Checked = UseOneCurve

                    If Confirm_File(FireDatabaseName, readfile, 1) = 0 Then FireDatabaseName = UserPersonalDataFolder & gcs_folder_ext & DefaultFireDatabaseName
                    If Confirm_File(MaterialsDatabaseName, readfile, 1) = 0 Then MaterialsDatabaseName = UserPersonalDataFolder & gcs_folder_ext & DefaultMaterialsDatabaseName

                    frmOptions1.txtErrorControl.Text = CStr(Error_Control)
                    frmOptions1.txtErrorControlVentFlow.Text = CStr(Error_Control_ventflow)

                    'disable menu item run if no heat release data entered
                    For i = 1 To NumberObjects
                        If NumberDataPoints(i) <> 0 And NumberObjects > 0 Then
                            'MDIFrmMain.mnuRun.Enabled = True
                            frmInputs.StartToolStripLabel1.Visible = True
                            Exit For
                        Else
                            'MDIFrmMain.mnuRun.Enabled = False
                            frmInputs.StartToolStripLabel1.Visible = False
                        End If
                    Next i

                    If LEsolver = "LU Decomposition" Or LEsolver = "" Then
                        frmOptions1.optLUdecom.Checked = True
                    Else
                        frmOptions1.optGaussJor.Checked = True
                    End If
                    If Enhance = True Then
                        frmOptions1.optEnhanceOn.Checked = True
                    Else
                        frmOptions1.optEnhanceOff.Checked = True
                    End If

                    If illumination = True Then
                        frmOptions1.optIlluminatedSign.Checked = True
                    Else
                        frmOptions1.optReflectiveSign.Checked = True
                    End If

                    If corner = 0 Then
                        MDIFrmMain.mnuQuintiereGraph.Visible = False
                        MDIFrmMain.mnuConeHRR.Visible = False
                    Else
                        MDIFrmMain.mnuQuintiereGraph.Visible = True
                        MDIFrmMain.mnuConeHRR.Visible = True
                    End If

                    If plume = 1 Then
                        frmOptions1.optStrongPlume.Checked = True
                    Else
                        frmOptions1.optMcCaffrey.Checked = True
                    End If

                    'If CeilingSlope(1) = True Then
                    '    frmDescribeRoom.optFlatCeiling.Checked = False
                    'Else
                    '    frmDescribeRoom.optFlatCeiling.Checked = True
                    'End If

                    'If HaveWallSubstrate(1) = True Then
                    '    frmDescribeRoom.chkAddWallSubstrate.Checked = True
                    'Else
                    '    frmDescribeRoom.chkAddWallSubstrate.Checked = False
                    'End If
                    'If HaveFloorSubstrate(1) = True Then
                    '    frmDescribeRoom.chkAddFloorSubstrate.Checked = True
                    'Else
                    '    frmDescribeRoom.chkAddFloorSubstrate.Checked = False
                    'End If
                    'If HaveCeilingSubstrate(1) = True Then
                    '    frmDescribeRoom.chkAddSubstrate.Checked = True
                    'Else
                    '    frmDescribeRoom.chkAddSubstrate.Checked = False
                    'End If

                    frmOptions1.txtNodes(0).Text = CStr(Ceilingnodes)
                    frmOptions1.txtNodes(1).Text = CStr(Wallnodes)
                    frmOptions1.txtNodes(2).Text = CStr(Floornodes)

                    If frmFire.lstObjectLocation.Items.Count > 0 Then
                        frmFire.lstObjectLocation.SelectedIndex = FireLocation(1)
                    Else
                        frmFire.lstObjectLocation.SelectedIndex = -1
                    End If
                    frmOptions1.txtExteriorTemp.Text = CStr(ExteriorTemp - 273)
                    frmOptions1.txtInteriorTemp.Text = CStr(InteriorTemp - 273)
                    frmOptions1.txtFlameLengthPower.Text = CStr(FlameLengthPower)
                    frmOptions1.txtBurnerWidth.Text = CStr(BurnerWidth)
                    frmOptions1.txtFlameAreaConstant.Text = CStr(FlameAreaConstant)
                    frmOptions1.txtCeilingHeatFlux.Text = CStr(CeilingHeatFlux)
                    frmOptions1.txtWallHeatFlux.Text = CStr(WallHeatFlux)
                    'frmOptions1.txtRadiantLossFraction.Text = CStr(RadiantLossFraction)
                    'frmOptions1.txtMassLossUnitArea.Text = CStr(MLUnitArea)
                    frmOptions1.txtEmissionCoefficient.Text = CStr(EmissionCoefficient)
                    frmOptions1.txtpreCO.Text = CStr(preCO)
                    frmOptions1.txtpostCO.Text = CStr(postCO)
                    frmOptions1.txtpreSoot.Text = CStr(preSoot)
                    frmOptions1.txtPostSoot.Text = CStr(postSoot)
                    frmOptions1.optCOman.Checked = comode
                    frmOptions1.optSootman.Checked = sootmode
                    frmOptions1.txtSootAlpha.Text = CStr(SootAlpha)
                    frmOptions1.txtSootEps.Text = CStr(SootEpsilon)
                    'frmDescribeRoom.lblCeilingSubstrate.Text = CeilingSubstrate(1)
                    'frmDescribeRoom.lblFloorSubstrate.Text = FloorSubstrate(1)
                    'frmDescribeRoom.lblCeilingSurface.Text = CeilingSurface(1)
                    'frmDescribeRoom.lblWallSubstrate.Text = WallSubstrate(1)
                    'frmDescribeRoom.lblWallSurface.Text = WallSurface(1)
                    'frmDescribeRoom.lblFloorSurface.Text = FloorSurface(1)

                    If corner = 1 Then
                        frmOptions1.optKarlsson.Checked = True
                        'MDIFrmMain.StatusBar1.Panels(3).Text = "Room-Corner Simulation ON"
                    ElseIf corner = 2 Then
                        frmOptions1.optQuintiere.Checked = True
                        'MDIFrmMain.StatusBar1.Panels(3).Text = "Room-Corner Simulation ON"
                    Else
                        frmOptions1.optRCNone.Checked = True
                        'MDIFrmMain.StatusBar1.Panels(3).Text = "Room-Corner Simulation OFF"
                    End If

                    'frmSprink.cboDetectorType.SelectedIndex = DetectorType(fireroom)
                    'frmOptions1.lstGraphRoom.SelectedIndex = 0

                    If IgniteNextRoom = True Then frmQuintiere.chkSpreadAdjacentRoom.CheckState = System.Windows.Forms.CheckState.Checked

                    If NumberRooms > 1 Then
                        'frmVentGeom.shpWall(0).Height = VB6.TwipsToPixelsY(1000 * WallLength1(1, 2, 1))
                        'frmVentGeom.shpWall(1).Height = VB6.TwipsToPixelsY(1000 * WallLength2(1, 2, 1))
                    End If

                    'matname = Dir(MaterialsDatabaseName)
                    matname = Path.GetFileName(MaterialsDatabaseName)

                    firename = Path.GetFileName(FireDatabaseName)
                    If opendatafile = My.Application.Info.DirectoryPath & "\ISO9705.tem" Then
                        DataFile = UserPersonalDataFolder & gcs_folder_ext & "\data\ISO9705.mod"
                        MsgBox("ISO9705 template file Successfully Loaded", MB_ICONINFORMATION, ProgramTitle)

                    Else
                        If batch = False Then MsgBox(DataFile & " Successfully Loaded", MB_ICONINFORMATION, ProgramTitle)
                    End If
                End If
            Else
                'no file selected
                MsgBox("Select File to Open", MB_ICONINFORMATION, ProgramTitle)
            End If

            frmInputs.rtb_log.Clear()

            'Dim bname As String = Path.GetFileName(DataFile)
            'bname = bname.Replace(",mod", "")
            'frmInputs.txtBaseName.Text = bname

            Exit Sub


        Catch ex As Exception
            MsgBox(Err.Description & " Line " & Err.Erl, MsgBoxStyle.Exclamation, "Exception in Read_OutputFile_xml ")
            'If Err.Number = 62 Then Resume Next 'end of file
            Action = File_Errors(Err.Number)
            Select Case Action
                Case 1
                    'Stop
                    'Resume Next
                Case Else
                    FileClose(1)
                    MsgBox("There was an error reading the data file.")
                    Exit Sub
            End Select
        End Try

    End Sub
    Public Sub Read_File_xml(ByRef opendatafile As String, ByVal batch As Boolean)
        'read input.xml

        Dim myfile As String
        Dim k, i, id, plume As Short
        Dim SuppressCeiling, HCNcalc As Boolean
        Dim dummy As String = ""
        Dim room As Short
        ' Dim illumination, post As Boolean
        Dim post As Boolean
        Dim ver As String = ""
        Dim LEsolver As String = ""
        Dim matname, firename, time_tmp, hrr_tmp As String
        Dim kmax(MaxNumberRooms + 1, MaxNumberRooms + 1) As Integer
        Dim kvmax(MaxNumberRooms + 1, MaxNumberRooms + 1) As Integer

        Try


            If opendatafile <> "" Then

                'returns filename without path if it exists
                'myfile = Dir(opendatafile)
                myfile = Path.GetFileName(opendatafile)

                If myfile <> "" Then

                    If myfile = "sprinklers.xml" Then
                        MsgBox("This XML file does not appear to be a valid input file.", MsgBoxStyle.Critical)
                        Exit Sub
                    End If
                    'New_File()

                    Dim myXmlSettings As New XmlReaderSettings With {
                        .IgnoreComments = True,
                        .IgnoreWhitespace = True
                    }

                    Using DFR As XmlReader = XmlReader.Create(opendatafile, myXmlSettings)
                        DFR.Read()

                        DFR.ReadStartElement("simulation")
                        DFR.ReadStartElement("general_settings")
                        ver = DFR.ReadElementString()
                        dummy = DFR.ReadElementString()

                        If dummy <> "input" Then
                            MsgBox("This XML file does not appear to be a BRANZFIRE input file.", MsgBoxStyle.Critical)
                            Exit Sub
                        End If

                        Description = DFR.ReadElementString()
                        If ver >= 2010.3 Then VM2 = DFR.ReadElementString()
                        InteriorTemp = DFR.ReadElementString()
                        ExteriorTemp = DFR.ReadElementString()
                        RelativeHumidity = DFR.ReadElementString()
                        SimTime = DFR.ReadElementString()
                        DisplayInterval = DFR.ReadElementString()

                        Ceilingnodes = DFR.ReadElementString()
                        Wallnodes = DFR.ReadElementString()
                        Floornodes = DFR.ReadElementString()
                        Enhance = DFR.ReadElementString()
                        JobNumber = DFR.ReadElementString()
                        ExcelInterval = DFR.ReadElementString()
                        Timestep = DFR.ReadElementString()
                        Error_Control = DFR.ReadElementString()
                        If CDbl(ver) >= 2012.47 Then Error_Control_ventflow = DFR.ReadElementString()

                        FireDatabaseName = DFR.ReadElementString()
                        MaterialsDatabaseName = DFR.ReadElementString()
                        cjModel = DFR.ReadElementString()
                        ventlog = DFR.ReadElementString()
                        LEsolver = DFR.ReadElementString()
                        nowallflow = DFR.ReadElementString()

                        DFR.ReadEndElement() 'clear </general_settings>

                        NumberRooms = CInt(DFR.Item("number_rooms"))
                        DFR.ReadStartElement("rooms")

                        Resize_Rooms()

                        For i = 1 To NumberRooms
                            room = CInt(DFR.Item("id"))
                            CeilingSlope(room) = CBool(DFR.Item("ceilingslope"))
                            DFR.ReadStartElement("room")
                            RoomWidth(room) = DFR.ReadElementString()
                            RoomLength(room) = DFR.ReadElementString()
                            RoomHeight(room) = DFR.ReadElementString()
                            RoomDescription(room) = DFR.ReadElementString()
                            MinStudHeight(room) = DFR.ReadElementString()
                            FloorElevation(room) = DFR.ReadElementString()
                            TwoZones(room) = DFR.ReadElementString()

                            If CSng(ver) > 2010 Then
                                RoomAbsX(room) = DFR.ReadElementString() 'abs X position
                                RoomAbsY(room) = DFR.ReadElementString() 'abs y position
                            End If

                            DFR.ReadStartElement("wall_lining")
                            WallSurface(room) = DFR.ReadElementString()
                            WallThickness(room) = DFR.ReadElementString()
                            WallConductivity(room) = DFR.ReadElementString()
                            WallSpecificHeat(room) = DFR.ReadElementString()
                            WallDensity(room) = DFR.ReadElementString()
                            Surface_Emissivity(2, room) = DFR.ReadElementString()
                            Surface_Emissivity(3, room) = Surface_Emissivity(2, room)
                            WallConeDataFile(room) = DFR.ReadElementString()
                            WallTSMin(room) = DFR.ReadElementString()
                            WallFlameSpreadParameter(room) = DFR.ReadElementString()
                            WallEffectiveHeatofCombustion(room) = DFR.ReadElementString()
                            WallSootYield(room) = DFR.ReadElementString()
                            WallCO2Yield(room) = DFR.ReadElementString()
                            WallH2OYield(room) = DFR.ReadElementString()
                            WallHCNYield(room) = DFR.ReadElementString()
                            If CSng(ver) > 2015.02 Then PessimiseCombWall = DFR.ReadElementString()
                            DFR.ReadEndElement() 'clear </wall_lining>

                            HaveWallSubstrate(room) = CBool(DFR.Item("present"))
                            DFR.ReadStartElement("wall_substrate")
                            If HaveWallSubstrate(room) = True Then
                                WallSubstrate(room) = DFR.ReadElementString()
                                WallSubThickness(room) = DFR.ReadElementString()
                                WallSubConductivity(room) = DFR.ReadElementString()
                                WallSubSpecificHeat(room) = DFR.ReadElementString()
                                WallSubDensity(room) = DFR.ReadElementString()
                                DFR.ReadEndElement() 'clear </wall_substrate>
                            End If

                            DFR.ReadStartElement("ceiling_lining")
                            CeilingSurface(room) = DFR.ReadElementString()
                            CeilingThickness(room) = DFR.ReadElementString()
                            CeilingConductivity(room) = DFR.ReadElementString()
                            CeilingSpecificHeat(room) = DFR.ReadElementString()
                            CeilingDensity(room) = DFR.ReadElementString()
                            Surface_Emissivity(1, room) = DFR.ReadElementString()
                            CeilingConeDataFile(room) = DFR.ReadElementString()
                            CeilingEffectiveHeatofCombustion(room) = DFR.ReadElementString()
                            CeilingSootYield(room) = DFR.ReadElementString()
                            CeilingCO2Yield(room) = DFR.ReadElementString()
                            CeilingH2OYield(room) = DFR.ReadElementString()
                            CeilingHCNYield(room) = DFR.ReadElementString()
                            DFR.ReadEndElement() 'clear </ceiling_lining>

                            HaveCeilingSubstrate(room) = CBool(DFR.Item("present"))
                            DFR.ReadStartElement("ceiling_substrate")
                            If HaveCeilingSubstrate(room) = True Then
                                CeilingSubstrate(room) = DFR.ReadElementString()
                                CeilingSubThickness(room) = DFR.ReadElementString()
                                CeilingSubConductivity(room) = DFR.ReadElementString()
                                CeilingSubSpecificHeat(room) = DFR.ReadElementString()
                                CeilingSubDensity(room) = DFR.ReadElementString()
                                DFR.ReadEndElement() 'clear </ceiling_substrate>
                            End If


                            DFR.ReadStartElement("floor_lining")
                            FloorSurface(room) = DFR.ReadElementString()
                            FloorThickness(room) = DFR.ReadElementString()
                            FloorConductivity(room) = DFR.ReadElementString()
                            FloorSpecificHeat(room) = DFR.ReadElementString()
                            FloorDensity(room) = DFR.ReadElementString()
                            Surface_Emissivity(4, room) = DFR.ReadElementString()
                            FloorConeDataFile(room) = DFR.ReadElementString()
                            FloorTSMin(room) = DFR.ReadElementString()
                            FloorFlameSpreadParameter(room) = DFR.ReadElementString()
                            FloorEffectiveHeatofCombustion(room) = DFR.ReadElementString()
                            FloorSootYield(room) = DFR.ReadElementString()
                            FloorCO2Yield(room) = DFR.ReadElementString()
                            FloorH2OYield(room) = DFR.ReadElementString()
                            FloorHCNYield(room) = DFR.ReadElementString()
                            DFR.ReadEndElement() 'clear </floor_lining>

                            HaveFloorSubstrate(room) = CBool(DFR.Item("present"))
                            DFR.ReadStartElement("floor_substrate")
                            If HaveFloorSubstrate(room) = True Then
                                FloorSubstrate(room) = DFR.ReadElementString()
                                FloorSubThickness(room) = DFR.ReadElementString()
                                FloorSubConductivity(room) = DFR.ReadElementString()
                                FloorSubSpecificHeat(room) = DFR.ReadElementString()
                                FloorSubDensity(room) = DFR.ReadElementString()
                                DFR.ReadEndElement() 'clear </floor_substrate>
                            End If

                            DFR.ReadEndElement() 'clear </room>
                        Next i

                        DFR.ReadEndElement() 'clear </rooms>

                        corner = CShort(DFR.Item("algorithm"))
                        DFR.ReadStartElement("flamespread")
                        'corner = DFR.ReadElementString()
                        If corner > 0 Then
                            SuppressCeiling = DFR.ReadElementString()
                            FlameAreaConstant = DFR.ReadElementString()
                            FlameLengthPower = DFR.ReadElementString()
                            BurnerWidth = DFR.ReadElementString()
                            WallHeatFlux = DFR.ReadElementString()
                            CeilingHeatFlux = DFR.ReadElementString()
                            IgniteNextRoom = DFR.ReadElementString()
                            UseOneCurve = DFR.ReadElementString()
                            IgnCorrelation = DFR.ReadElementString()
                            DFR.ReadEndElement() 'clear </flamespread>
                        End If

                        DFR.ReadStartElement("tenability")
                        MonitorHeight = DFR.ReadElementString()
                        Activity = DFR.ReadElementString()
                        TargetEndPoint = DFR.ReadElementString()
                        TempEndPoint = DFR.ReadElementString() - 273
                        VisibilityEndPoint = DFR.ReadElementString()
                        FEDEndPoint = DFR.ReadElementString()
                        ConvectEndPoint = DFR.ReadElementString()
                        StartOccupied = DFR.ReadElementString()
                        EndOccupied = DFR.ReadElementString()
                        illumination = DFR.ReadElementString()
                        DFR.ReadEndElement() 'clear </tenability>

                        post = CBool(DFR.Item("post"))
                        g_post = post
                        DFR.ReadStartElement("postflashover")
                        If post = True Then
                            FLED = DFR.ReadElementString()

                            Fuel_Thickness = DFR.ReadElementString()
                            HoC_fuel = DFR.ReadElementString()
                            Stick_Spacing = DFR.ReadElementString()
                            If CSng(ver) > 2014.17 Then
                                Cribheight = DFR.ReadElementString()
                            End If
                            If CSng(ver) > 2014.15 Then
                                ExcessFuelFactor = DFR.ReadElementString()
                            End If
                            DFR.ReadEndElement() 'clear </postflashover>
                        End If

                        DFR.ReadStartElement("chemistry")
                        nC = DFR.ReadElementString()
                        nH = DFR.ReadElementString()
                        nO = DFR.ReadElementString()
                        nN = DFR.ReadElementString()
                        If CDbl(ver) > 2016.03 Then
                            StoichAFratio = DFR.ReadElementString()
                        End If
                        fueltype = DFR.ReadElementString()
                        HCNcalc = DFR.ReadElementString()
                        SootAlpha = DFR.ReadElementString()
                        SootEpsilon = DFR.ReadElementString()
                        EmissionCoefficient = DFR.ReadElementString()
                        preCO = DFR.ReadElementString()
                        postCO = DFR.ReadElementString()
                        preSoot = DFR.ReadElementString()
                        postSoot = DFR.ReadElementString()
                        comode = DFR.ReadElementString()
                        sootmode = DFR.ReadElementString()
                        DFR.ReadEndElement() 'clear </chemistry>

                        DFR.ReadStartElement("fires")
                        fireroom = DFR.ReadElementString()

                        If CSng(ver) < 2012.24 Then
                            dummy = DFR.ReadElementString()
                            dummy = DFR.ReadElementString()
                        End If


                        plume = DFR.ReadElementString()
                        NumberObjects = DFR.ReadElementString()

                        Resize_Objects()

                        For i = 1 To NumberObjects
                            'id = DFR.Item("id")
                            id = i
                            ObjectItemID(id) = DFR.Item("id")

                            If CSng(ver) > 2011.16 Then
                                ObjectDescription(id) = DFR.Item("description")
                            End If
                            If CSng(ver) > 2012.04 Then
                                ObjLabel(id) = DFR.Item("userlabel")
                            End If
                            DFR.ReadStartElement("fire")

                            EnergyYield(id) = DFR.ReadElementString()
                            CO2Yield(id) = DFR.ReadElementString()
                            SootYield(id) = DFR.ReadElementString()
                            'WaterVaporYield(id) = DFR.ReadElementString()
                            HCNuserYield(id) = DFR.ReadElementString()
                            FireHeight(id) = DFR.ReadElementString()
                            FireLocation(id) = DFR.ReadElementString()
                            NumberDataPoints(id) = DFR.ReadElementString()
                            ObjCRF(id) = DFR.ReadElementString()
                            ObjFTPindexpilot(id) = DFR.ReadElementString()
                            ObjFTPlimitpilot(id) = DFR.ReadElementString()
                            ObjCRFauto(id) = DFR.ReadElementString()
                            ObjFTPindexauto(id) = DFR.ReadElementString()
                            ObjFTPlimitauto(id) = DFR.ReadElementString()
                            ObjLength(id) = DFR.ReadElementString()
                            ObjWidth(id) = DFR.ReadElementString()
                            ObjHeight(id) = DFR.ReadElementString()
                            ObjDimX(id) = DFR.ReadElementString()
                            ObjDimY(id) = DFR.ReadElementString()
                            ObjElevation(id) = DFR.ReadElementString()
                            ObjectIgnTime(id) = DFR.ReadElementString()
                            ObjectRLF(id) = DFR.ReadElementString()
                            If DFR.Name = "windeffect" Then ObjectWindEffect(id) = DFR.ReadElementString()
                            If DFR.Name = "pyrolysisoption" Then ObjectPyrolysisOption(id) = DFR.ReadElementString()
                            If DFR.Name = "pooldensity" Then ObjectPoolDensity(id) = DFR.ReadElementString()
                            If DFR.Name = "pooldiameter" Then ObjectPoolDiameter(id) = DFR.ReadElementString()
                            If DFR.Name = "poolFBMLR" Then ObjectPoolFBMLR(id) = DFR.ReadElementString()
                            If DFR.Name = "poolramp" Then ObjectPoolRamp(id) = DFR.ReadElementString()
                            If DFR.Name = "poolvolume" Then ObjectPoolVolume(id) = DFR.ReadElementString()

                            If CSng(ver) > 2012.04 Then
                                ObjectMLUA(0, id) = DFR.ReadElementString()
                                ObjectMLUA(1, id) = DFR.ReadElementString()
                                If CSng(ver) > 2012.26 Then
                                    ObjectMLUA(2, id) = DFR.ReadElementString()
                                End If
                                ObjectLHoG(id) = DFR.ReadElementString()
                            End If

                            dummy = DFR.ReadElementString() 'coordinates of the corners of the object 
                            dummy = DFR.ReadElementString()
                            dummy = DFR.ReadElementString()
                            dummy = DFR.ReadElementString()
                            time_tmp = DFR.ReadElementString()
                            hrr_tmp = DFR.ReadElementString()
                            'need to parse these strings and copy into HeatReleaseData array
                            Dim tempvar1 As String() = time_tmp.Split(CChar(","))
                            Dim tempvar2 As String() = hrr_tmp.Split(CChar(","))
                            For k = 1 To NumberDataPoints(id)
                                HeatReleaseData(1, k, id) = tempvar1(k - 1)
                                HeatReleaseData(2, k, id) = tempvar2(k - 1)
                            Next
                            AlphaT = DFR.ReadElementString()
                            PeakHRR = DFR.ReadElementString()
                            DFR.ReadEndElement() 'clear </fire>
                        Next
                        DFR.ReadEndElement() 'clear </fires>

                        DFR.ReadStartElement("hvents")
                        Do While (DFR.Name = "hvent")
                            DFR.ReadStartElement("hvent")
                            room = DFR.ReadElementString() ' room 1
                            i = DFR.ReadElementString() ' room 2
                            k = DFR.ReadElementString() ' vent id

                            If k > kmax(room, i) Then
                                kmax(room, i) = k
                                NumberVents(room, i) = kmax(room, i)
                                NumberVents(i, room) = kmax(room, i)
                                Resize_Vents()
                            End If

                            VentHeight(room, i, k) = DFR.ReadElementString()
                            VentWidth(room, i, k) = DFR.ReadElementString()
                            VentSillHeight(room, i, k) = DFR.ReadElementString()
                            VentOpenTime(room, i, k) = DFR.ReadElementString()
                            VentCloseTime(room, i, k) = DFR.ReadElementString()

                            If CSng(ver) > 2012.23 And CSng(ver) < 2014.18 Then VentCD(room, i, k) = DFR.ReadElementString()

                            WallLength1(room, i, k) = DFR.ReadElementString()
                            WallLength2(room, i, k) = DFR.ReadElementString()
                            If CSng(ver) > 2012.42 Then VentOffset(room, i, k) = DFR.ReadElementString()
                            If CSng(ver) > 2010.1 Then VentFace(room, i, k) = DFR.ReadElementString()
                            If CSng(ver) > 2012.4 Then HOReliability(room, i, k) = DFR.ReadElementString()
                            If CSng(ver) > 2012.45 Then
                                VentCD(room, i, k) = DFR.ReadElementString()
                                dummy = DFR.ReadElementString()
                                HOactive(room, i, k) = DFR.ReadElementString()
                            End If

                            AutoBreakGlass(room, i, k) = CBool(DFR.Item("autobreak"))
                            DFR.ReadStartElement("glassbreak")
                            If AutoBreakGlass(room, i, k) = True Then
                                GLASSconductivity(room, i, k) = DFR.ReadElementString()
                                GLASSemissivity(room, i, k) = DFR.ReadElementString()
                                GLASSexpansion(room, i, k) = DFR.ReadElementString()
                                GLASSthickness(room, i, k) = DFR.ReadElementString()
                                GLASSshading(room, i, k) = DFR.ReadElementString()
                                GLASSbreakingstress(room, i, k) = DFR.ReadElementString()
                                GLASSalpha(room, i, k) = DFR.ReadElementString()
                                GlassYoungsModulus(room, i, k) = DFR.ReadElementString()
                                GLASSFalloutTime(room, i, k) = DFR.ReadElementString()
                                GLASSdistance(room, i, k) = DFR.ReadElementString()
                                GlassFlameFlux(room, i, k) = DFR.ReadElementString()
                                DFR.ReadEndElement() 'clear </glassbreak>
                            End If

                            spillplume(room, i, k) = CShort(DFR.Item("use_spillplume"))
                            spillplume(i, room, k) = spillplume(room, i, k)

                            DFR.ReadStartElement("spillplume")

                            If spillplume(room, i, k) <> 0 Then
                                Downstand(room, i, k) = DFR.ReadElementString()
                                SpillPlumeBalc(room, i, k) = DFR.ReadElementString()
                                spillplumemodel(room, i, k) = DFR.ReadElementString()
                                spillbalconyprojection(room, i, k) = DFR.ReadElementString()

                                spillplumemodel(i, room, k) = spillplumemodel(room, i, k)
                                spillbalconyprojection(i, room, k) = spillbalconyprojection(room, i, k)

                                DFR.ReadEndElement() 'clear </spillplume>
                            End If
                            DFR.ReadEndElement() 'clear </hvent>

                            VentHeight(i, room, k) = VentHeight(room, i, k)
                            VentWidth(i, room, k) = VentWidth(room, i, k)
                            If i <= NumberRooms Then VentSillHeight(i, room, k) = VentSillHeight(room, i, k) + FloorElevation(room) - FloorElevation(i)
                            VentOpenTime(i, room, k) = VentOpenTime(room, i, k)
                            VentCloseTime(i, room, k) = VentCloseTime(room, i, k)
                            WallLength1(i, room, k) = WallLength1(room, i, k)
                            WallLength2(i, room, k) = WallLength2(room, i, k)
                            Downstand(i, room, k) = Downstand(room, i, k)
                            GLASSconductivity(i, room, k) = GLASSconductivity(room, i, k)
                            GLASSemissivity(i, room, k) = GLASSemissivity(room, i, k)
                            GLASSexpansion(i, room, k) = GLASSexpansion(room, i, k)
                            AutoBreakGlass(i, room, k) = AutoBreakGlass(room, i, k)
                            GLASSthickness(i, room, k) = GLASSthickness(room, i, k)
                            GLASSFalloutTime(i, room, k) = GLASSFalloutTime(room, i, k)
                            GLASSshading(i, room, k) = GLASSshading(room, i, k)
                            GLASSbreakingstress(i, room, k) = GLASSbreakingstress(room, i, k)
                            GLASSalpha(i, room, k) = GLASSalpha(room, i, k)
                            GlassYoungsModulus(i, room, k) = GlassYoungsModulus(room, i, k)

                        Loop
                        If DFR.Name = "hvents" Then DFR.ReadEndElement() 'clear </hvents>

                        DFR.ReadStartElement("vvents")
                        Do While (DFR.Name = "vvent")
                            DFR.ReadStartElement("vvent")
                            room = DFR.ReadElementString() ' room 1 upper
                            i = DFR.ReadElementString() ' room 2 lower
                            k = DFR.ReadElementString() ' vent id

                            If k > kvmax(room, i) Then
                                kvmax(room, i) = k
                                NumberCVents(room, i) = kvmax(room, i)
                                'NumberCVents(i, room) = kvmax(room, i)
                                Resize_CVents()
                            End If

                            CVentAuto(room, i, k) = DFR.ReadElementString()
                            CVentArea(room, i, k) = DFR.ReadElementString()
                            CVentOpenTime(room, i, k) = DFR.ReadElementString()
                            CVentCloseTime(room, i, k) = DFR.ReadElementString()
                            DFR.ReadEndElement() 'clear </vvent>
                        Loop
                        If DFR.Name = "vvents" Then DFR.ReadEndElement() 'clear </vvents>

                        'this is not needed if we save in separate xml file
                        sprink_mode = DFR.Item("id")


                        DFR.ReadStartElement("sprinklers")
                        Do While (DFR.Name = "sprinkler")
                            id = DFR.Item("id")
                            DFR.ReadStartElement("sprinkler")

                            dummy = DFR.ReadElementString()
                            dummy = DFR.ReadElementString()
                            dummy = DFR.ReadElementString()
                            dummy = DFR.ReadElementString()
                            dummy = DFR.ReadElementString()
                            dummy = DFR.ReadElementString()
                            dummy = DFR.ReadElementString()
                            dummy = DFR.ReadElementString()
                            DFR.ReadEndElement() 'clear </sprinkler>
                        Loop
                        If DFR.Name = "sprinklers" Then DFR.ReadEndElement() 'clear </sprinklers>

                        If CSng(ver) < 2012.45 Then
                            SDReliability = DFR.Item("sd_reliability")
                        Else
                            SDReliability = DFR.Item("sys_reliability")
                        End If
                        If CSng(ver) > 2012.45 Then
                            sd_mode = DFR.Item("operational_status")
                        End If
                        DFR.ReadStartElement("smoke_detectors")
                        Do While (DFR.Name = "smoke_detector")
                            id = DFR.Item("id")
                            room = DFR.Item("room")
                            If room > NumberRooms Then
                                room = NumberRooms
                                MsgBox("Smoke detector " & id & " in room that doesn't exist", MsgBoxStyle.Exclamation, )
                            End If

                            DFR.ReadStartElement("smoke_detector")
                            HaveSD(room) = True
                            SmokeOD(room) = DFR.ReadElementString()
                            SDRadialDist(room) = DFR.ReadElementString()
                            SDdepth(room) = DFR.ReadElementString()
                            dummy = DFR.ReadElementString() 'sdx
                            dummy = DFR.ReadElementString() 'sdy
                            SDdelay(room) = DFR.ReadElementString()

                            DFR.ReadEndElement() 'clear </smoke_detector>
                        Loop
                        If DFR.Name = "smoke_detectors" Then DFR.ReadEndElement() 'clear </smoke_detectors>

                        If CSng(ver) < 2012.45 Then
                            FanReliability = DFR.Item("fan_reliability")
                        Else
                            FanReliability = DFR.Item("sys_reliability")
                        End If
                        If CSng(ver) > 2012.45 Then
                            mv_mode = DFR.Item("operational_status")
                        End If
                        DFR.ReadStartElement("fans")
                        Do While (DFR.Name = "fan")
                            id = DFR.Item("id")
                            room = DFR.Item("room")
                            DFR.ReadStartElement("fan")
                            dummy = DFR.ReadElementString() 'flow_rate
                            dummy = DFR.ReadElementString() 'start_time
                            dummy = DFR.ReadElementString() 'fan_reliability
                            dummy = DFR.ReadElementString() 'pressure_limit
                            dummy = DFR.ReadElementString() 'elevation
                            dummy = DFR.ReadElementString() 'extract
                            dummy = DFR.ReadElementString() 'manual
                            dummy = DFR.ReadElementString() 'fan_curve
                            dummy = DFR.ReadElementString() 'operational status

                            DFR.ReadEndElement() 'clear fan
                        Loop
                        If DFR.Name = "fans" Then DFR.ReadEndElement() 'clear fans

                        'DFR.ReadStartElement("fans")
                        'Do While (DFR.Name = "fan")
                        '    DFR.ReadStartElement("fan")
                        '    room = DFR.ReadElementString()
                        '    fanon(room) = DFR.ReadElementString()
                        '    ExtractRate(room) = DFR.ReadElementString()
                        '    ExtractStartTime(room) = DFR.ReadElementString()
                        '    MaxPressure(room) = DFR.ReadElementString()
                        '    NumberFans(room) = DFR.ReadElementString()
                        '    Extract(room) = DFR.ReadElementString()
                        '    UseFanCurve(room) = DFR.ReadElementString()
                        '    FanElevation(room) = DFR.ReadElementString()
                        '    FanAutoStart(room) = DFR.ReadElementString()
                        '    DFR.ReadEndElement() 'clear </fan>
                        'Loop
                        'If DFR.Name = "fans" Then DFR.ReadEndElement() 'clear </fans>

                        DFR.ReadEndElement() 'clear </simulation>
                    End Using

                    'frmSprink.optControl.Checked = Sprinkler(1, fireroom)
                    'frmSprink.optSprinkOff.Checked = Sprinkler(2, fireroom)
                    'frmSprink.optSuppression.Checked = Sprinkler(3, fireroom)
                    'frmSprinklerList.optControl.Checked = Sprinkler(1, fireroom)
                    'frmSprinklerList.optSprinkOff.Checked = Sprinkler(2, fireroom)
                    'frmSprinklerList.optSuppression.Checked = Sprinkler(3, fireroom)

                    frmQuintiere.optUseOneCurve.Checked = UseOneCurve
                    'frmOptions1.txtFLED.Text = CStr(FLED)

                    frmOptions1.txtFuelThickness.Text = CStr(Fuel_Thickness)
                    frmOptions1.txtStickSpacing.Text = CStr(Stick_Spacing)
                    frmOptions1.txtCribheight.Text = CStr(Cribheight)
                    frmOptions1.txtExcessFuelFactor.Text = CStr(ExcessFuelFactor)
                    ' frmOptions1.txtHOCFuel.Text = CStr(HoC_fuel)
                    frmOptions1.txtnC.Text = CStr(nC)
                    frmOptions1.txtnH.Text = CStr(nH)
                    frmOptions1.txtnO.Text = CStr(nO)
                    frmOptions1.txtnN.Text = CStr(nN)
                    frmOptions1.txtStoich.Text = CStr(StoichAFratio)
                    frmOptions1.cboABSCoeff.Text = fueltype
                    frmOptions1.txtNodes(0).Text = CStr(Ceilingnodes)
                    frmOptions1.txtNodes(1).Text = CStr(Wallnodes)
                    frmOptions1.txtNodes(2).Text = CStr(Floornodes)
                    frmOptions1.txtErrorControl.Text = CStr(Error_Control)
                    frmOptions1.txtErrorControlVentFlow.Text = CStr(Error_Control_ventflow)
                    frmOptions1.txtExteriorTemp.Text = CStr(ExteriorTemp - 273)
                    frmOptions1.txtInteriorTemp.Text = CStr(InteriorTemp - 273)
                    frmOptions1.txtFlameLengthPower.Text = CStr(FlameLengthPower)
                    frmOptions1.txtBurnerWidth.Text = CStr(BurnerWidth)
                    frmOptions1.txtFlameAreaConstant.Text = CStr(FlameAreaConstant)
                    frmOptions1.txtCeilingHeatFlux.Text = CStr(CeilingHeatFlux)
                    frmOptions1.txtWallHeatFlux.Text = CStr(WallHeatFlux)
                    'frmOptions1.txtRadiantLossFraction.Text = CStr(RadiantLossFraction)
                    ' frmOptions1.txtMassLossUnitArea.Text = CStr(MLUnitArea)
                    frmOptions1.txtEmissionCoefficient.Text = CStr(EmissionCoefficient)
                    frmOptions1.txtpreCO.Text = CStr(preCO)
                    frmOptions1.txtpostCO.Text = CStr(postCO)
                    frmOptions1.txtpreSoot.Text = CStr(preSoot)
                    frmOptions1.txtPostSoot.Text = CStr(postSoot)
                    frmOptions1.optCOman.Checked = comode
                    frmOptions1.optSootman.Checked = sootmode
                    frmOptions1.txtSootAlpha.Text = CStr(SootAlpha)
                    frmOptions1.txtSootEps.Text = CStr(SootEpsilon)
                    'frmDescribeRoom.lblCeilingSubstrate.Text = CeilingSubstrate(1)
                    'frmDescribeRoom.lblFloorSubstrate.Text = FloorSubstrate(1)
                    'frmDescribeRoom.lblCeilingSurface.Text = CeilingSurface(1)
                    'frmDescribeRoom.lblWallSubstrate.Text = WallSubstrate(1)
                    'frmDescribeRoom.lblWallSurface.Text = WallSurface(1)
                    'frmDescribeRoom.lblFloorSurface.Text = FloorSurface(1)

                    If Confirm_File(FireDatabaseName, readfile, 1) = 0 Then FireDatabaseName = UserPersonalDataFolder & gcs_folder_ext & DefaultFireDatabaseName
                    If Confirm_File(MaterialsDatabaseName, readfile, 1) = 0 Then MaterialsDatabaseName = UserPersonalDataFolder & gcs_folder_ext & DefaultMaterialsDatabaseName

                    If post = True Then
                        frmOptions1.optPostFlashover.Checked = True
                    Else
                        frmOptions1.OptPreFlashover.Checked = True
                    End If

                    If nowallflow = True Then
                        frmOptions1.chkWallFlowDisable.CheckState = System.Windows.Forms.CheckState.Checked
                    Else
                        frmOptions1.chkWallFlowDisable.CheckState = System.Windows.Forms.CheckState.Unchecked
                    End If

                    If HCNcalc = True Then
                        frmOptions1.chkHCNcalc.CheckState = System.Windows.Forms.CheckState.Checked
                    Else
                        frmOptions1.chkHCNcalc.CheckState = System.Windows.Forms.CheckState.Unchecked
                    End If

                    If ventlog = True Then
                        frmInputs.chkSaveVentFlow.CheckState = System.Windows.Forms.CheckState.Checked
                    Else
                        frmInputs.chkSaveVentFlow.CheckState = System.Windows.Forms.CheckState.Unchecked
                    End If

                    'disable menu item run if no heat release data entered
                    For i = 1 To NumberObjects
                        If NumberObjects > 0 Then
                            'MDIFrmMain.mnuRun.Enabled = True
                            frmInputs.StartToolStripLabel1.Visible = True
                            Exit For
                        Else
                            'MDIFrmMain.mnuRun.Enabled = False
                            frmInputs.StartToolStripLabel1.Visible = False
                        End If
                    Next i

                    If LEsolver = "LU decomposition" Or LEsolver = "" Then
                        frmOptions1.optLUdecom.Checked = True
                    Else
                        frmOptions1.optGaussJor.Checked = True
                    End If

                    If Enhance = True Then
                        frmOptions1.optEnhanceOn.Checked = True
                    Else
                        frmOptions1.optEnhanceOff.Checked = True
                    End If

                    If illumination = True Then
                        frmOptions1.optIlluminatedSign.Checked = True
                    Else
                        frmOptions1.optReflectiveSign.Checked = True
                    End If

                    If corner = 0 Then
                        MDIFrmMain.mnuQuintiereGraph.Visible = False
                        MDIFrmMain.mnuConeHRR.Visible = False
                    Else
                        MDIFrmMain.mnuQuintiereGraph.Visible = True
                        MDIFrmMain.mnuConeHRR.Visible = True
                    End If

                    If plume = 1 Then
                        frmOptions1.optStrongPlume.Checked = True
                    Else
                        frmOptions1.optMcCaffrey.Checked = True
                    End If

                    'If CeilingSlope(1) = True Then
                    '    frmDescribeRoom.optFlatCeiling.Checked = False
                    'Else
                    '    frmDescribeRoom.optFlatCeiling.Checked = True
                    'End If

                    'If HaveWallSubstrate(1) = True Then
                    '    frmDescribeRoom.chkAddWallSubstrate.Checked = True
                    'Else
                    '    frmDescribeRoom.chkAddWallSubstrate.Checked = False
                    'End If

                    'If HaveFloorSubstrate(1) = True Then
                    '    frmDescribeRoom.chkAddFloorSubstrate.Checked = True
                    'Else
                    '    frmDescribeRoom.chkAddFloorSubstrate.Checked = False
                    'End If

                    'If HaveCeilingSubstrate(1) = True Then
                    '    frmDescribeRoom.chkAddSubstrate.Checked = True
                    'Else
                    '    frmDescribeRoom.chkAddSubstrate.Checked = False
                    'End If

                    If frmFire.lstObjectLocation.Items.Count > 0 Then
                        frmFire.lstObjectLocation.SelectedIndex = FireLocation(1)
                    Else
                        frmFire.lstObjectLocation.SelectedIndex = -1
                    End If

                    If corner = 1 Then
                        frmOptions1.optKarlsson.Checked = True
                    ElseIf corner = 2 Then
                        frmOptions1.optQuintiere.Checked = True
                    Else
                        frmOptions1.optRCNone.Checked = True
                    End If

                    'frmSprink.cboDetectorType.SelectedIndex = DetectorType(fireroom)

                    If IgniteNextRoom = True Then frmQuintiere.chkSpreadAdjacentRoom.CheckState = System.Windows.Forms.CheckState.Checked

                    If NumberRooms > 1 Then
                        'frmVentGeom.shpWall(0).Height = VB6.TwipsToPixelsY(1000 * WallLength1(1, 2, 1))
                        'frmVentGeom.shpWall(1).Height = VB6.TwipsToPixelsY(1000 * WallLength2(1, 2, 1))
                    End If

                    If VM2 = True Then
                        MDIFrmMain.NZBCVM2ToolStripMenuItem.Checked = CHECKED
                        MDIFrmMain.RiskSimulatorToolStripMenuItem.Checked = UNCHECKED
                    Else
                        MDIFrmMain.RiskSimulatorToolStripMenuItem.Checked = CHECKED
                        MDIFrmMain.NZBCVM2ToolStripMenuItem.Checked = UNCHECKED
                    End If

                    'matname = Dir(MaterialsDatabaseName)
                    matname = Path.GetFileName(MaterialsDatabaseName)
                    'firename = Dir(FireDatabaseName)
                    firename = Path.GetFileName(FireDatabaseName)

                    'acknowledge to the user that file is successfully loaded
                    If opendatafile = UserPersonalDataFolder & gcs_folder_ext & "\data\ISO9705.tem" Then
                        DataFile = UserPersonalDataFolder & gcs_folder_ext & "\data\ISO9705.xml"
                        MsgBox("ISO9705 template file Successfully Loaded", MB_ICONINFORMATION, ProgramTitle)
                    Else
                        If batch = False Then MsgBox(DataFile & " Successfully Loaded", MB_ICONINFORMATION, ProgramTitle)
                    End If
                End If

            End If

            'MDIFrmMain.ToolStripStatusLabel1.Text = MDIFrmMain.OpenFileDialog1.SafeFileName
            'MDIFrmMain.ToolStripStatusLabel2.Text = Description
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in Read_File_xml")

        End Try
    End Sub

    Public Sub Save_output_xml()
        'save simulation output to xml file

        Dim room, i, j As Integer
        Dim myXMLsettings As New XmlWriterSettings
        Dim outputfile As String = DataFile.Replace("input", "output")

        Try

            myXMLsettings.Indent = True
            myXMLsettings.NewLineOnAttributes = True

            'outputfile = UserAppDataFolder & gcs_folder_ext & "\data\" & outputfile
            outputfile = RiskDataDirectory & outputfile

            Using DFW As XmlWriter = XmlWriter.Create(outputfile, myXMLsettings)

                DFW.WriteComment("Created by B-RISK Version " & Version)

                DFW.WriteStartElement("run")
                DFW.WriteAttributeString("id", DataFile)
                DFW.WriteAttributeString("runtime", runtime)
                If NumberTimeSteps > 0 Then
                    For room = 1 To NumberRooms
                        DFW.WriteStartElement("room")
                        DFW.WriteAttributeString("id", room)

                        DFW.WriteStartElement("items")
                        DFW.WriteAttributeString("value", NumberObjects)
                        'add ignition times of secondary objects
                        For i = 2 To NumberObjects
                            DFW.WriteStartElement("secondary item ignition time")
                            DFW.WriteAttributeString("id", ObjectItemID(i))
                            DFW.WriteAttributeString("description", ObjectDescription(i))
                            DFW.WriteAttributeString("value", ObjectIgnTime(i))
                            DFW.WriteAttributeString("units", "sec")
                            DFW.WriteAttributeString("value", ObjectIgnMode(i))
                            DFW.WriteEndElement()
                        Next

                        For i = 1 To NumberTimeSteps + 1
                            DFW.WriteStartElement("time")
                            DFW.WriteAttributeString("value", tim(i, 1))
                            DFW.WriteAttributeString("units", "sec")

                            DFW.WriteStartElement("HeatRelease")
                            DFW.WriteAttributeString("value", HeatRelease(room, i, 2))
                            DFW.WriteAttributeString("units", "kW")
                            'DFW.WriteStartElement("value", HeatRelease(room, i, 2))
                            'DFW.WriteEndElement()
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("layerheight")
                            DFW.WriteAttributeString("value", layerheight(room, i))
                            DFW.WriteAttributeString("units", "m")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("uppertemp")
                            DFW.WriteAttributeString("value", uppertemp(room, i) - 273)
                            DFW.WriteAttributeString("units", "C")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("lowertemp")
                            DFW.WriteAttributeString("value", lowertemp(room, i) - 273)
                            DFW.WriteAttributeString("units", "C")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("FuelMassLossRate")
                            DFW.WriteAttributeString("value", FuelMassLossRate(i, 1))
                            DFW.WriteAttributeString("units", "kg/s")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("massplumeflow")
                            DFW.WriteAttributeString("value", massplumeflow(i, room))
                            DFW.WriteAttributeString("units", "kg/s")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("ventfire")
                            DFW.WriteAttributeString("value", ventfire(room, i))
                            DFW.WriteAttributeString("units", "kW")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("CO2VolumeFraction-u")
                            DFW.WriteAttributeString("value", CO2VolumeFraction(room, i, 1) * 100)
                            DFW.WriteAttributeString("units", "%")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("COVolumeFraction-u")
                            DFW.WriteAttributeString("value", COVolumeFraction(room, i, 1) * 1000000)
                            DFW.WriteAttributeString("units", "ppm")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("O2VolumeFraction-u")
                            DFW.WriteAttributeString("value", O2VolumeFraction(room, i, 1) * 100)
                            DFW.WriteAttributeString("units", "%")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("CO2VolumeFraction-l")
                            DFW.WriteAttributeString("value", CO2VolumeFraction(room, i, 2) * 100)
                            DFW.WriteAttributeString("units", "%")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("COVolumeFraction-l")
                            DFW.WriteAttributeString("value", COVolumeFraction(room, i, 2) * 1000000)
                            DFW.WriteAttributeString("units", "ppm")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("O2VolumeFraction-l")
                            DFW.WriteAttributeString("value", O2VolumeFraction(room, i, 2) * 100)
                            DFW.WriteAttributeString("units", "%")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("FEDSum")
                            DFW.WriteAttributeString("value", FEDSum(room, i))
                            DFW.WriteAttributeString("units", "-")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("Upperwalltemp")
                            DFW.WriteAttributeString("value", Upperwalltemp(room, i) - 273)
                            DFW.WriteAttributeString("units", "C")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("lowerwalltemp")
                            DFW.WriteAttributeString("value", LowerWallTemp(room, i) - 273)
                            DFW.WriteAttributeString("units", "C")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("CeilingTemp")
                            DFW.WriteAttributeString("value", CeilingTemp(room, i) - 273)
                            DFW.WriteAttributeString("units", "C")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("FloorTemp")
                            DFW.WriteAttributeString("value", FloorTemp(room, i) - 273)
                            DFW.WriteAttributeString("units", "C")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("target")
                            DFW.WriteAttributeString("value", Target(room, i))
                            DFW.WriteAttributeString("units", "kW")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("Y_pyrolysis")
                            DFW.WriteAttributeString("value", Y_pyrolysis(room, i))
                            DFW.WriteAttributeString("units", "m")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("X_pyrolysis")
                            DFW.WriteAttributeString("value", X_pyrolysis(room, i))
                            DFW.WriteAttributeString("units", "m")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("Z_pyrolysis")
                            DFW.WriteAttributeString("value", Z_pyrolysis(room, i))
                            DFW.WriteAttributeString("units", "m")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("UpwardFlameVelocity")
                            DFW.WriteAttributeString("value", FlameVelocity(room, 1, i))
                            DFW.WriteAttributeString("units", "m/s")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("LateralFlameVelocity")
                            DFW.WriteAttributeString("value", FlameVelocity(room, 2, i))
                            DFW.WriteAttributeString("units", "m/s")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("RoomPressure")
                            DFW.WriteAttributeString("value", RoomPressure(room, i))
                            DFW.WriteAttributeString("units", "Pa")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("Visibility")
                            DFW.WriteAttributeString("value", Visibility(room, i))
                            DFW.WriteAttributeString("units", "m")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("FlowToUpper")
                            DFW.WriteAttributeString("value", FlowToUpper(room, i))
                            DFW.WriteAttributeString("units", "kg/s")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("FlowToLower")
                            DFW.WriteAttributeString("value", FlowToLower(room, i))
                            DFW.WriteAttributeString("units", "kg/s")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("SurfaceRad")
                            DFW.WriteAttributeString("value", SurfaceRad(room, i))
                            DFW.WriteAttributeString("units", "kW")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("FEDRadSum")
                            DFW.WriteAttributeString("value", FEDRadSum(room, i))
                            DFW.WriteAttributeString("units", "-")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("OD_upper")
                            DFW.WriteAttributeString("value", OD_upper(room, i))
                            DFW.WriteAttributeString("units", "1/m")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("OD_lower")
                            DFW.WriteAttributeString("value", OD_lower(room, i))
                            DFW.WriteAttributeString("units", "1/m")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("k_upper")
                            DFW.WriteAttributeString("value", 2.3 * OD_upper(room, i))
                            DFW.WriteAttributeString("units", "1/m")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("k_lower")
                            DFW.WriteAttributeString("value", 2.3 * OD_lower(room, i))
                            DFW.WriteAttributeString("units", "1/m")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("UFlowToOutside")
                            DFW.WriteAttributeString("value", UFlowToOutside(room, i))
                            DFW.WriteAttributeString("units", "m3/s")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("HCNVolumeFraction-u")
                            DFW.WriteAttributeString("value", HCNVolumeFraction(room, i, 1) * 1000000)
                            DFW.WriteAttributeString("units", "ppm")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("HCNVolumeFraction-l")
                            DFW.WriteAttributeString("value", HCNVolumeFraction(room, i, 2) * 1000000)
                            DFW.WriteAttributeString("units", "ppm")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("SPR")
                            DFW.WriteAttributeString("value", SPR(room, i))
                            DFW.WriteAttributeString("units", "m2/s")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("UnexposedUpperwalltemp")
                            DFW.WriteAttributeString("value", UnexposedUpperwalltemp(room, i) - 273)
                            DFW.WriteAttributeString("units", "C")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("UnexposedLowerwalltemp")
                            DFW.WriteAttributeString("value", UnexposedLowerwalltemp(room, i) - 273)
                            DFW.WriteAttributeString("units", "C")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("UnexposedLowerwalltemp")
                            DFW.WriteAttributeString("value", UnexposedCeilingtemp(room, i) - 273)
                            DFW.WriteAttributeString("units", "C")
                            DFW.WriteEndElement()

                            DFW.WriteStartElement("UnexposedFloortemp")
                            DFW.WriteAttributeString("value", UnexposedFloortemp(room, i) - 273)
                            DFW.WriteAttributeString("units", "C")
                            DFW.WriteEndElement()

                            If room = fireroom Then
                                DFW.WriteStartElement("CJetVel")
                                DFW.WriteAttributeString("value", CJetTemp(j, 0, 0))
                                DFW.WriteAttributeString("units", "m/s")
                                DFW.WriteEndElement()

                                DFW.WriteStartElement("CJetTemp")
                                DFW.WriteAttributeString("value", CJetTemp(j, 1, 0) - 273)
                                DFW.WriteAttributeString("units", "C")
                                DFW.WriteEndElement()

                                DFW.WriteStartElement("MaxCJetTemp")
                                DFW.WriteAttributeString("value", CJetTemp(j, 2, 0) - 273)
                                DFW.WriteAttributeString("units", "C")
                                DFW.WriteEndElement()

                                DFW.WriteStartElement("GlobalER")
                                DFW.WriteAttributeString("value", GlobalER(i))
                                DFW.WriteAttributeString("units", "-")
                                DFW.WriteEndElement()

                                DFW.WriteStartElement("OD_outside")
                                DFW.WriteAttributeString("value", OD_outside(room, i))
                                DFW.WriteAttributeString("units", "1/m")
                                DFW.WriteEndElement()

                                DFW.WriteStartElement("OD_inside")
                                DFW.WriteAttributeString("value", OD_inside(room, i))
                                DFW.WriteAttributeString("units", "1/m")
                                DFW.WriteEndElement()

                                DFW.WriteStartElement("LinkTemp")
                                DFW.WriteAttributeString("value", LinkTemp(room, i) - 273)
                                DFW.WriteAttributeString("units", "C")
                                DFW.WriteEndElement()
                            End If

                            DFW.WriteEndElement() 'end room
                        Next
                        DFW.WriteEndElement() 'end time
                    Next
                End If
                DFW.WriteEndElement() 'end run
            End Using

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in MODULE1.vb save_output_xml")
        End Try

    End Sub

    Public Sub Save_File_xml(ByVal filename As String, ByVal counter As Integer, ByVal oSprinklers As Object, ByVal oSmokeDets As Object, ByVal oFans As Object)
        'create input file
        'input.xml

        Dim room As Short
        Dim HCNcalc As Boolean

        Dim plume, i, k, corner As Short
        Dim SuppressCeiling As Boolean
        Dim post As Boolean
        Dim tempvar1 As String = ""
        Dim tempvar2 As String = ""

        Try

            If frmOptions1.optPostFlashover.Checked = True Then
                post = True
                g_post = post
            Else
                post = False
                g_post = post
            End If

            If frmOptions1.optStrongPlume.Checked = True Then
                plume = 1
            Else
                plume = 2
            End If

            If frmOptions1.optEnhanceOn.Checked = True Then
                Enhance = True
            Else
                Enhance = False
            End If

            If frmOptions1.optKarlsson.Checked = True Then
                corner = 1
            ElseIf frmOptions1.optQuintiere.Checked = True Then
                corner = 2
            Else
                corner = 0
            End If

            'comode = frmOptions1.optCOman.Checked
            'sootmode = frmOptions1.optSootman.Checked

            If frmOptions1.optIlluminatedSign.Checked = True Then
                illumination = True
            Else
                illumination = False
            End If

            If frmQuintiere.chkSpreadAdjacentRoom.CheckState = System.Windows.Forms.CheckState.Checked Then
                IgniteNextRoom = True
            Else
                IgniteNextRoom = False
            End If

            ventlog = frmInputs.chkSaveVentFlow.CheckState
            If frmOptions1.chkWallFlowDisable.CheckState = 1 Then
                nowallflow = True
            Else
                nowallflow = False
            End If

            HCNcalc = frmOptions1.chkHCNcalc.CheckState

            Dim myXMLsettings As New XmlWriterSettings With {
                .Indent = True,
                .NewLineOnAttributes = True
            }

            'Using DFW As XmlWriter = XmlWriter.Create("datafile.xml", myXMLsettings)
            Using DFW As XmlWriter = XmlWriter.Create(filename, myXMLsettings)
                DFW.WriteComment("Created by B-RISK Version " & Version)
                DFW.WriteComment("Input File " & Convert.ToString(MDIFrmMain.Text))

                DFW.WriteStartElement("simulation")

                DFW.WriteStartElement("general_settings")
                DFW.WriteElementString("version", Version)
                DFW.WriteElementString("file_type", "input")
                DFW.WriteElementString("description", Description)
                DFW.WriteElementString("user_mode", VM2)
                DFW.WriteElementString("temp_interior", InteriorTemp)
                DFW.WriteElementString("temp_exterior", ExteriorTemp)
                DFW.WriteElementString("rel_humidity", RelativeHumidity)
                DFW.WriteElementString("simulation_duration", SimTime)
                DFW.WriteElementString("display_interval", DisplayInterval)
                DFW.WriteElementString("ceiling_nodes", Ceilingnodes)
                DFW.WriteElementString("wall_nodes", Wallnodes)
                DFW.WriteElementString("floor_nodes", Floornodes)
                DFW.WriteElementString("enhance_burning", Enhance)
                DFW.WriteElementString("job_number", JobNumber)
                DFW.WriteElementString("excel_interval", ExcelInterval)
                DFW.WriteElementString("time_step", Timestep)
                DFW.WriteElementString("error_control", Error_Control)
                DFW.WriteElementString("error_control_ventflow", Error_Control_ventflow)
                DFW.WriteElementString("fire_dbase", FireDatabaseName)
                DFW.WriteElementString("mat_dbase", MaterialsDatabaseName)
                DFW.WriteElementString("ceiling_jet", cjModel)
                DFW.WriteElementString("vent_logfile", ventlog)
                DFW.WriteElementString("LE_Solver", LEsolver)
                DFW.WriteElementString("no_wall_flow", nowallflow)

                'If frmOptions1.optLUdecom.Checked = True Then
                ' DFW.WriteElementString("LE_solver", "LU decomposition")
                'Else
                'DFW.WriteElementString("LE_solver", "Gauss-Jordan")
                'End If

                DFW.WriteEndElement()

                DFW.WriteStartElement("rooms")
                DFW.WriteAttributeString("number_rooms", NumberRooms)

                For room = 1 To NumberRooms
                    DFW.WriteStartElement("room")
                    DFW.WriteAttributeString("id", room)
                    DFW.WriteAttributeString("ceilingslope", CeilingSlope(room))
                    DFW.WriteElementString("width", RoomWidth(room))
                    DFW.WriteElementString("length", RoomLength(room))
                    DFW.WriteElementString("max_height", RoomHeight(room))
                    DFW.WriteElementString("description", RoomDescription(room))
                    DFW.WriteElementString("min_height", MinStudHeight(room))
                    DFW.WriteElementString("floor_elevation", FloorElevation(room))
                    DFW.WriteElementString("two_zones", TwoZones(room))
                    DFW.WriteElementString("abs_X", RoomAbsX(room))
                    DFW.WriteElementString("abs_Y", RoomAbsY(room))

                    DFW.WriteStartElement("wall_lining")
                    DFW.WriteElementString("description", WallSurface(room))
                    DFW.WriteElementString("thickness", WallThickness(room))
                    DFW.WriteElementString("conductivity", WallConductivity(room))
                    DFW.WriteElementString("specific_heat", WallSpecificHeat(room))
                    DFW.WriteElementString("density", WallDensity(room))
                    DFW.WriteElementString("emissivity", Surface_Emissivity(2, room))
                    'DFW.WriteElementString("emissivity_lower", Surface_Emissivity(3, room))
                    DFW.WriteElementString("cone_file", WallConeDataFile(room))
                    DFW.WriteElementString("min_temp_spread", WallTSMin(room))
                    DFW.WriteElementString("flame_spread_parameter", WallFlameSpreadParameter(room))
                    DFW.WriteElementString("eff_heat_of_combustion", WallEffectiveHeatofCombustion(room))
                    DFW.WriteElementString("soot_yield", WallSootYield(room))
                    DFW.WriteElementString("CO2_yield", WallCO2Yield(room))
                    DFW.WriteElementString("H20_yield", WallH2OYield(room))
                    DFW.WriteElementString("HCN_yield", WallHCNYield(room))
                    DFW.WriteElementString("Pessimise_comb_wall", PessimiseCombWall)
                    DFW.WriteEndElement()

                    DFW.WriteStartElement("wall_substrate")
                    DFW.WriteAttributeString("present", HaveWallSubstrate(room))
                    If HaveWallSubstrate(room) = True Then
                        DFW.WriteElementString("description", WallSubstrate(room))
                        DFW.WriteElementString("thickness", WallSubThickness(room))
                        DFW.WriteElementString("conductivity", WallSubConductivity(room))
                        DFW.WriteElementString("specific_heat", WallSubSpecificHeat(room))
                        DFW.WriteElementString("density", WallSubDensity(room))
                    End If
                    DFW.WriteEndElement()

                    DFW.WriteStartElement("ceiling_lining")
                    DFW.WriteElementString("description", CeilingSurface(room))
                    DFW.WriteElementString("thickness", CeilingThickness(room))
                    DFW.WriteElementString("conductivity", CeilingConductivity(room))
                    DFW.WriteElementString("specific_heat", CeilingSpecificHeat(room))
                    DFW.WriteElementString("density", CeilingDensity(room))
                    DFW.WriteElementString("emissivity", Surface_Emissivity(1, room))
                    DFW.WriteElementString("ceiling_cone_file", CeilingConeDataFile(room))
                    DFW.WriteElementString("eff_heat_of_combustion", WallEffectiveHeatofCombustion(room))
                    DFW.WriteElementString("soot_yield", CeilingSootYield(room))
                    DFW.WriteElementString("CO2_yield", CeilingCO2Yield(room))
                    DFW.WriteElementString("H20_yield", CeilingH2OYield(room))
                    DFW.WriteElementString("HCN_yield", CeilingHCNYield(room))
                    DFW.WriteEndElement()

                    DFW.WriteStartElement("ceiling_substrate")
                    DFW.WriteAttributeString("present", HaveCeilingSubstrate(room))
                    If HaveCeilingSubstrate(room) = True Then
                        DFW.WriteElementString("description", CeilingSubstrate(room))
                        DFW.WriteElementString("thickness", CeilingSubThickness(room))
                        DFW.WriteElementString("conductivity", CeilingSubConductivity(room))
                        DFW.WriteElementString("specific_heat", CeilingSubSpecificHeat(room))
                        DFW.WriteElementString("density", CeilingSubDensity(room))
                    End If
                    DFW.WriteEndElement()

                    DFW.WriteStartElement("floor_lining")
                    DFW.WriteElementString("description", FloorSurface(room))
                    DFW.WriteElementString("thickness", FloorThickness(room))
                    DFW.WriteElementString("conductivity", FloorConductivity(room))
                    DFW.WriteElementString("specific_heat", FloorSpecificHeat(room))
                    DFW.WriteElementString("density", FloorDensity(room))
                    DFW.WriteElementString("emissivity", Surface_Emissivity(4, room))
                    DFW.WriteElementString("floor_cone_file", FloorConeDataFile(room))
                    DFW.WriteElementString("min_temp_spread", FloorTSMin(room))
                    DFW.WriteElementString("flame_spread_parameter", FloorFlameSpreadParameter(room))
                    DFW.WriteElementString("eff_heat_of_combustion", FloorEffectiveHeatofCombustion(room))
                    DFW.WriteElementString("soot_yield", FloorSootYield(room))
                    DFW.WriteElementString("CO2_yield", FloorCO2Yield(room))
                    DFW.WriteElementString("H20_yield", FloorH2OYield(room))
                    DFW.WriteElementString("HCN_yield", FloorHCNYield(room))

                    DFW.WriteEndElement()

                    DFW.WriteStartElement("floor_substrate")
                    DFW.WriteAttributeString("present", HaveFloorSubstrate(room))
                    If HaveFloorSubstrate(room) = True Then
                        DFW.WriteElementString("description", FloorSubstrate(room))
                        DFW.WriteElementString("thickness", FloorSubThickness(room))
                        DFW.WriteElementString("conductivity", FloorSubConductivity(room))
                        DFW.WriteElementString("specific_heat", FloorSubSpecificHeat(room))
                        DFW.WriteElementString("density", FloorSubDensity(room))
                    End If
                    DFW.WriteEndElement()

                    DFW.WriteEndElement()
                Next

                DFW.WriteEndElement()

                DFW.WriteStartElement("flamespread")
                DFW.WriteAttributeString("algorithm", corner)
                'DFW.WriteElementString("flamespread_algorithm", corner)
                If corner > 0 Then
                    DFW.WriteElementString("suppress_ceiling_hrr", SuppressCeiling)
                    DFW.WriteElementString("flame_area_constant", FlameAreaConstant)
                    DFW.WriteElementString("flame_length_power", FlameLengthPower)
                    DFW.WriteElementString("burner_width", BurnerWidth)
                    DFW.WriteElementString("wall_heat_flux", WallHeatFlux)
                    DFW.WriteElementString("ceiling_heat_flux", CeilingHeatFlux)
                    DFW.WriteElementString("ignite_next_room", IgniteNextRoom)
                    DFW.WriteElementString("one_cone_curve", UseOneCurve)
                    DFW.WriteElementString("ign_correlation", IgnCorrelation)
                    DFW.WriteElementString("pessimise_comb_wall", PessimiseCombWall)
                End If
                DFW.WriteEndElement()

                DFW.WriteStartElement("tenability")
                DFW.WriteElementString("monitor_height", MonitorHeight)
                DFW.WriteElementString("activity_level", Activity)
                DFW.WriteElementString("endpoint_radiation", TargetEndPoint)
                DFW.WriteElementString("endpoint_temp", TempEndPoint)
                DFW.WriteElementString("endpoint_visibility", VisibilityEndPoint)
                DFW.WriteElementString("endpoint_FED", FEDEndPoint)
                DFW.WriteElementString("endpoint_convect", ConvectEndPoint)
                DFW.WriteElementString("FED_start_time", StartOccupied)
                DFW.WriteElementString("FED_end_time", EndOccupied)
                DFW.WriteElementString("illumination", illumination)
                DFW.WriteEndElement()

                DFW.WriteStartElement("postflashover")
                DFW.WriteAttributeString("post", post)
                If post = True Then
                    DFW.WriteElementString("FLED", FLED)

                    DFW.WriteElementString("fuel_thickness", Fuel_Thickness)
                    DFW.WriteElementString("HOC_fuel", HoC_fuel)
                    DFW.WriteElementString("stick_spacing", Stick_Spacing)
                    DFW.WriteElementString("crib_height", Cribheight)
                    DFW.WriteElementString("excess_fuel_factor", ExcessFuelFactor)
                End If
                DFW.WriteEndElement()

                DFW.WriteStartElement("chemistry")
                DFW.WriteElementString("nC", nC)
                DFW.WriteElementString("nH", nH)
                DFW.WriteElementString("nO", nO)
                DFW.WriteElementString("nN", nN)
                DFW.WriteElementString("stoic", StoichAFratio)
                DFW.WriteElementString("fueltype", fueltype)
                DFW.WriteElementString("hcn_calc", HCNcalc)
                DFW.WriteElementString("soot_alpha", SootAlpha)
                DFW.WriteElementString("soot_epsilon", SootEpsilon)
                DFW.WriteElementString("emission_coefficient", EmissionCoefficient)
                DFW.WriteElementString("pre_CO", preCO)
                DFW.WriteElementString("post_CO", postCO)
                DFW.WriteElementString("pre_soot", preSoot)
                DFW.WriteElementString("post_soot", postSoot)
                DFW.WriteElementString("CO_mode", comode)
                DFW.WriteElementString("soot_mode", sootmode)
                DFW.WriteEndElement()

                DFW.WriteStartElement("fires")
                DFW.WriteElementString("fire_room", fireroom)
                'DFW.WriteElementString("radiant_loss_fraction", RadiantLossFraction)
                'DFW.WriteElementString("mass_loss_per_unit_area", MLUnitArea)
                DFW.WriteComment("plume, macaffrey=2, delichatsios=1")
                DFW.WriteElementString("plume_algorithm", plume)
                DFW.WriteElementString("number_objects", NumberObjects)

                Resize_Objects()

                For k = 1 To NumberObjects
                    DFW.WriteStartElement("fire")
                    DFW.WriteAttributeString("id", ObjectItemID(k))
                    DFW.WriteAttributeString("description", ObjectDescription(k))
                    DFW.WriteAttributeString("userlabel", ObjLabel(k))
                    DFW.WriteElementString("heat_of_combustion", EnergyYield(k))
                    DFW.WriteElementString("CO2_yield", CO2Yield(k))
                    DFW.WriteElementString("soot_yield", SootYield(k))
                    'DFW.WriteElementString("H2O_yield", WaterVaporYield(k))
                    DFW.WriteElementString("HCN_yield", HCNuserYield(k))
                    DFW.WriteElementString("fire_height", FireHeight(k))
                    DFW.WriteComment("fire location, corner=2, wall=1, centre=0")
                    DFW.WriteElementString("fire_location", FireLocation(k))
                    DFW.WriteElementString("data_points", NumberDataPoints(k))
                    DFW.WriteElementString("obj_CRF_pilot", ObjCRF(k))
                    DFW.WriteElementString("obj_FTPindex_pilot", ObjFTPindexpilot(k))
                    DFW.WriteElementString("obj_FTPlimit_pilot", ObjFTPlimitpilot(k))
                    DFW.WriteElementString("obj_CRF_auto", ObjCRFauto(k))
                    DFW.WriteElementString("obj_FTPindex_auto", ObjFTPindexauto(k))
                    DFW.WriteElementString("obj_FTPlimit_auto", ObjFTPlimitauto(k))
                    DFW.WriteElementString("obj_length", ObjLength(k))
                    DFW.WriteElementString("obj_width", ObjWidth(k))
                    DFW.WriteElementString("obj_height", ObjHeight(k))
                    DFW.WriteElementString("obj_x", ObjDimX(k))
                    DFW.WriteElementString("obj_y", ObjDimY(k))
                    DFW.WriteElementString("obj_elevation", ObjElevation(k))
                    DFW.WriteElementString("obj_igntime", ObjectIgnTime(k))
                    DFW.WriteElementString("obj_RLF", ObjectRLF(k))
                    DFW.WriteElementString("windeffect", ObjectWindEffect(k))
                    DFW.WriteElementString("pyrolysisoption", ObjectPyrolysisOption(k))

                    DFW.WriteElementString("pooldensity", ObjectPoolDensity(k))
                    DFW.WriteElementString("pooldiameter", ObjectPoolDiameter(k))
                    DFW.WriteElementString("poolFBMLR", ObjectPoolFBMLR(k))
                    DFW.WriteElementString("poolramp", ObjectPoolRamp(k))
                    DFW.WriteElementString("poolvolume", ObjectPoolVolume(k))

                    DFW.WriteElementString("constantA", ObjectMLUA(0, k))
                    DFW.WriteElementString("constantB", ObjectMLUA(1, k))
                    DFW.WriteElementString("HRRUA", ObjectMLUA(2, k))
                    DFW.WriteElementString("obj_LHoG", ObjectLHoG(k))
                    DFW.WriteElementString("x1", item_location(3, k))
                    DFW.WriteElementString("x2", item_location(4, k))
                    DFW.WriteElementString("y1", item_location(5, k))
                    DFW.WriteElementString("y2", item_location(6, k))

                    tempvar1 = ""
                    tempvar2 = ""

                    For i = 1 To NumberDataPoints(k) - 1
                        tempvar1 = tempvar1 & HeatReleaseData(1, i, k) & ","
                        tempvar2 = tempvar2 & HeatReleaseData(2, i, k) & ","
                    Next
                    tempvar1 = tempvar1 & HeatReleaseData(1, i, k)
                    tempvar2 = tempvar2 & HeatReleaseData(2, i, k)

                    DFW.WriteElementString("time", tempvar1)
                    DFW.WriteElementString("HRR", tempvar2)

                    DFW.WriteElementString("alphaT", AlphaT)
                    DFW.WriteElementString("peakHRR", PeakHRR)

                    DFW.WriteEndElement()
                Next k

                DFW.WriteEndElement()

                'needs to know number of vents in each wall segment
                Resize_Vents()
                Resize_CVents()

                DFW.WriteStartElement("hvents")
                For room = 1 To NumberRooms
                    For i = 2 To NumberRooms + 1
                        If room < i Then
                            If NumberVents(room, i) > 0 Then
                                For k = 1 To NumberVents(room, i)
                                    DFW.WriteStartElement("hvent")
                                    DFW.WriteElementString("room_1", room)
                                    DFW.WriteElementString("room_2", i)
                                    DFW.WriteElementString("id", k)
                                    DFW.WriteElementString("height", VentHeight(room, i, k))
                                    DFW.WriteElementString("width", VentWidth(room, i, k))
                                    DFW.WriteElementString("sill_height", VentSillHeight(room, i, k))
                                    DFW.WriteElementString("open_time", VentOpenTime(room, i, k))
                                    DFW.WriteElementString("close_time", VentCloseTime(room, i, k))
                                    'DFW.WriteElementString("close_time", VentCD(room, i, k))

                                    DFW.WriteElementString("wall_length_1", WallLength1(room, i, k))
                                    DFW.WriteElementString("wall_length_2", WallLength2(room, i, k))
                                    DFW.WriteElementString("offset", VentOffset(room, i, k))
                                    DFW.WriteElementString("face", VentFace(room, i, k))
                                    DFW.WriteElementString("holdopen_reliability", HOReliability(room, i, k))
                                    DFW.WriteElementString("discharge_coeff", VentCD(room, i, k))
                                    If VentCD(room, i, k) = 0 Then
                                        DFW.WriteElementString("vent_status", "CLOSED")
                                    Else
                                        DFW.WriteElementString("vent_status", "OPEN")
                                    End If

                                    DFW.WriteElementString("holdopen_status", HOactive(room, i, k))


                                    DFW.WriteStartElement("glassbreak")
                                    DFW.WriteAttributeString("autobreak", AutoBreakGlass(room, i, k))
                                    If AutoBreakGlass(room, i, k) = True Then
                                        DFW.WriteElementString("conductivity", GLASSconductivity(room, i, k))
                                        DFW.WriteElementString("emissivity", GLASSemissivity(room, i, k))
                                        DFW.WriteElementString("expansion_coefficient", GLASSexpansion(room, i, k))
                                        DFW.WriteElementString("thickness", GLASSthickness(room, i, k))
                                        DFW.WriteElementString("shading_depth", GLASSshading(room, i, k))
                                        DFW.WriteElementString("breaking_stress", GLASSbreakingstress(room, i, k))
                                        DFW.WriteElementString("thermal_diffusivity", GLASSalpha(room, i, k))
                                        DFW.WriteElementString("youngs_modulus", GlassYoungsModulus(room, i, k))
                                        DFW.WriteElementString("fallout_time", GLASSFalloutTime(room, i, k))
                                        DFW.WriteElementString("glasstoflame_distance", GLASSdistance(room, i, k))
                                        DFW.WriteElementString("flame_flux", GlassFlameFlux(room, i, k))
                                    End If
                                    DFW.WriteEndElement()

                                    DFW.WriteStartElement("spillplume")
                                    DFW.WriteAttributeString("use_spillplume", spillplume(room, i, k))
                                    If spillplume(room, i, k) <> 0 Then
                                        DFW.WriteElementString("downstand_depth", Downstand(room, i, k))
                                        DFW.WriteElementString("balc_extends", SpillPlumeBalc(room, i, k))
                                        DFW.WriteElementString("algorithm", spillplumemodel(room, i, k))
                                        DFW.WriteElementString("single_sided", spillbalconyprojection(room, i, k))
                                    End If
                                    DFW.WriteEndElement()
                                    DFW.WriteEndElement()
                                Next
                            End If
                        End If
                    Next i
                Next room
                DFW.WriteEndElement()

                DFW.WriteStartElement("vvents")
                For room = 1 To NumberRooms
                    For i = 2 To NumberRooms + 1
                        If room < i Then
                            If NumberCVents(i, room) > 0 Then
                                For k = 1 To NumberCVents(i, room)
                                    DFW.WriteStartElement("vvent")
                                    DFW.WriteElementString("room_upper", i)
                                    DFW.WriteElementString("room_lower", room)
                                    DFW.WriteElementString("id", k)
                                    DFW.WriteElementString("Auto", CVentAuto(i, room, k))
                                    DFW.WriteElementString("area", CVentArea(i, room, k))
                                    DFW.WriteElementString("time_open", CVentOpenTime(i, room, k))
                                    DFW.WriteElementString("time_close", CVentCloseTime(i, room, k))
                                    DFW.WriteEndElement()
                                Next k
                            End If
                        End If
                    Next i
                Next room
                DFW.WriteEndElement()
                ' sprfile = DataDirectory + "sprinklers.xml"
                DFW.WriteStartElement("sprinklers")
                DFW.WriteAttributeString("sprink_mode", sprink_mode)
                DFW.WriteAttributeString("spr_reliability", Format(SprReliability, "0.0000"))
                DFW.WriteAttributeString("spr_suppression_prob", Format(SprSuppressionProb, "0.0000"))
                DFW.WriteAttributeString("spr_cooling_coefficient", Format(SprCooling, "0.0000"))
                DFW.WriteAttributeString("NumOperatingSpr", NumOperatingSpr)

                'should be able to pass osprinklers to this sub instead of reading file again?
                'osprinklers = SprinklerDB.GetSprinklers
                If mc_RTI IsNot Nothing Then

                    For Each oSprinkler In oSprinklers
                        DFW.WriteStartElement("sprinkler")
                        DFW.WriteAttributeString("id", oSprinkler.sprid)
                        'DFW.WriteElementString("RTI", oSprinkler.rti)
                        DFW.WriteElementString("RTI", CSng(mc_RTI(oSprinkler.sprid - 1, counter - 1)))
                        DFW.WriteElementString("c_factor", CSng(mc_cfactor(oSprinkler.sprid - 1, counter - 1)))
                        'DFW.WriteElementString("c_factor", oSprinkler.cfactor)
                        DFW.WriteElementString("radial_distance", CSng(mc_RadialDist(oSprinkler.sprid - 1, counter - 1)))
                        'DFW.WriteElementString("radial_distance", oSprinkler.sprr)
                        DFW.WriteElementString("actuation_temp", CSng(mc_acttemp(oSprinkler.sprid - 1, counter - 1)))
                        'DFW.WriteElementString("actuation_temp", oSprinkler.acttemp)
                        DFW.WriteElementString("water_spray_density", CSng(mc_waterdensity(oSprinkler.sprid - 1, counter - 1)))
                        'DFW.WriteElementString("water_spray_density", oSprinkler.sprdensity)
                        DFW.WriteElementString("depth", CSng(mc_dist(oSprinkler.sprid - 1, counter - 1)))
                        'DFW.WriteElementString("depth", oSprinkler.sprz)
                        DFW.WriteElementString("x-dim", oSprinkler.sprx)
                        DFW.WriteElementString("y-dim", oSprinkler.spry)
                        DFW.WriteEndElement() 'clear sprinkler
                    Next

                End If
                DFW.WriteEndElement() 'clear sprinklers

                DFW.WriteStartElement("smoke_detectors")
                ' DFW.WriteAttributeString("sd_reliability", Format(SDReliability, "0.0000"))
                DFW.WriteAttributeString("sys_reliability", Format(SDReliability, "0.0000"))
                DFW.WriteAttributeString("operational_status", sd_mode.ToString)

                If mc_SDOD IsNot Nothing Then
                    For Each oSmokeDet In oSmokeDets
                        DFW.WriteStartElement("smoke_detector")
                        DFW.WriteAttributeString("id", oSmokeDet.sdid)
                        DFW.WriteAttributeString("room", oSmokeDet.room)
                        DFW.WriteElementString("OD", oSmokeDet.od)
                        DFW.WriteElementString("radial_distance", oSmokeDet.sdr)
                        DFW.WriteElementString("depth", oSmokeDet.sdz)
                        DFW.WriteElementString("x-dim", oSmokeDet.sdx)
                        DFW.WriteElementString("y-dim", oSmokeDet.sdy)
                        DFW.WriteElementString("charlength", oSmokeDet.charlength)

                        DFW.WriteEndElement() 'clear sd
                    Next
                End If
                DFW.WriteEndElement() 'clear SD

                DFW.WriteStartElement("fans")
                DFW.WriteAttributeString("sys_reliability", Format(FanReliability, "0.0000"))
                DFW.WriteAttributeString("operational_status", mv_mode.ToString)
                If mc_fanflowrate IsNot Nothing Then
                    For Each ofan In oFans
                        DFW.WriteStartElement("fan")
                        DFW.WriteAttributeString("id", ofan.fanid)
                        DFW.WriteAttributeString("room", ofan.fanroom)
                        DFW.WriteElementString("flow_rate", ofan.fanflowrate)
                        DFW.WriteElementString("start_time", ofan.fanstarttime)
                        DFW.WriteElementString("fan_reliability", ofan.fanreliability)
                        DFW.WriteElementString("pressure_limit", ofan.fanpressurelimit)
                        DFW.WriteElementString("elevation", ofan.fanelevation)
                        DFW.WriteElementString("extract", ofan.fanextract)
                        DFW.WriteElementString("manual", ofan.fanmanual)
                        DFW.WriteElementString("fan_curve", ofan.fancurve)
                        DFW.WriteElementString("operational_status", FANactive(ofan.fanid))
                        DFW.WriteEndElement()
                    Next
                End If
                DFW.WriteEndElement() 'clear fan

                'DFW.WriteStartElement("fans")
                'For room = 1 To NumberRooms
                '    If fanon(room) = True Then
                '        DFW.WriteStartElement("fan")
                '        DFW.WriteElementString("location", room)
                '        DFW.WriteElementString("fan_on", fanon(room))
                '        DFW.WriteElementString("extract_rate", ExtractRate(room))
                '        DFW.WriteElementString("start_time", ExtractStartTime(room))
                '        DFW.WriteElementString("pressure_limit", MaxPressure(room))
                '        DFW.WriteElementString("number_fans", NumberFans(room))
                '        DFW.WriteElementString("extract_dir", Extract(room))
                '        DFW.WriteElementString("use_fan_curve", UseFanCurve(room))
                '        DFW.WriteElementString("elevation", FanElevation(room))
                '        DFW.WriteElementString("fan_auto", FanAutoStart(room))
                '        DFW.WriteEndElement()
                '    End If
                'Next room
                'DFW.WriteEndElement()

                DFW.WriteEndElement()

            End Using

            MDIFrmMain.ToolStripStatusLabel1.Text = Path.GetFileName(filename)
            MDIFrmMain.ToolStripStatusLabel2.Text = Description

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in MODULE1.vb Save_File_xml")
        End Try
    End Sub

    Public Sub Save_File(ByRef opendatafile As Object)
        '*  ====================================================================
        '*  Save the model variables in an ascii file
        '*  ====================================================================

        'Dim Floor, wall, ceiling, Lining As System.Windows.Forms.ComboBox
        'Dim WallSub, CeilingSub As System.Windows.Forms.ComboBox
        Dim room As Short
        'Dim illumination, HCNcalc As Boolean
        Dim HCNcalc As Boolean
        Dim plume, i, k As Short
        Dim SuppressCeiling As Boolean
        Dim post As Boolean

        If frmOptions1.optPostFlashover.Checked = True Then
            post = True
            g_post = post
        Else
            post = False
            g_post = post
        End If

        If frmOptions1.optStrongPlume.Checked = True Then
            plume = 1
        Else
            plume = 2
        End If

        If frmOptions1.optEnhanceOn.Checked = True Then
            Enhance = True
        Else
            Enhance = False
        End If

        If frmOptions1.optKarlsson.Checked = True Then
            corner = 1
        ElseIf frmOptions1.optQuintiere.Checked = True Then
            corner = 2
        Else
            corner = 0
        End If

        comode = frmOptions1.optCOman.Checked
        sootmode = frmOptions1.optSootman.Checked

        If frmOptions1.optIlluminatedSign.Checked = True Then
            illumination = True
        Else
            illumination = False
        End If
        'Get activity level
        'If frmOptions1.optAtRest.Value = True Then Activity = "Rest"
        'If frmOptions1.optLightWork.Value = True Then Activity = "Light"
        'If frmOptions1.optHeavyWork.Value = True Then Activity = "Heavy"

        If frmQuintiere.chkSpreadAdjacentRoom.CheckState = System.Windows.Forms.CheckState.Checked Then
            IgniteNextRoom = True
        Else
            IgniteNextRoom = False
        End If

        ventlog = frmInputs.chkSaveVentFlow.CheckState
        If frmOptions1.chkWallFlowDisable.CheckState = 1 Then
            nowallflow = True
        Else
            nowallflow = False
        End If

        HCNcalc = frmOptions1.chkHCNcalc.CheckState

        FileOpen(1, opendatafile, OpenMode.Output)
        WriteLine(1, "Created by B-RISK Version ", Version)
        WriteLine(1, Description)
        WriteLine(1, "room corner model, none=0, karlsson=1, quintiere=2", corner)
        WriteLine(1, "Number rooms", NumberRooms)
        For room = 1 To NumberRooms
            WriteLine(1, "Room Number", room)
            WriteLine(1, "room width (m)", RoomWidth(room))
            WriteLine(1, "room length (m)", RoomLength(room))
            WriteLine(1, "room description (m)", RoomDescription(room))
            WriteLine(1, "Max room Height(m)", RoomHeight(room))
            WriteLine(1, "Min room Height (m)", MinStudHeight(room))
            WriteLine(1, "floor elevation (m)", FloorElevation(room))
            WriteLine(1, "wall lining", WallSurface(room))
            WriteLine(1, "wall substrate", WallSubstrate(room))
            WriteLine(1, "ceiling lining", CeilingSurface(room))
            WriteLine(1, "ceiling substrate", CeilingSubstrate(room))
            WriteLine(1, "floor substrate", FloorSubstrate(room))
            WriteLine(1, "wall lining thickness (mm)", WallThickness(room))
            WriteLine(1, "ceiling lining thickness (mm)", CeilingThickness(room))
            WriteLine(1, "floor thickness (mm)", FloorThickness(room))
            WriteLine(1, "wall lining conductivity (W/mK)", WallConductivity(room))
            WriteLine(1, "ceiling lining conductivity (W/mK)", CeilingConductivity(room))
            WriteLine(1, "floor conductivity (W/mK)", FloorConductivity(room))
            WriteLine(1, "floor", FloorSurface(room))
            WriteLine(1, "wall lining specific heat (J/kgK)", WallSpecificHeat(room))
            WriteLine(1, "wall lining density (kg/m3)", WallDensity(room))
            WriteLine(1, "wall substrate thickness (mm)", WallSubThickness(room))
            WriteLine(1, "ceiling substrate thickness (mm)", CeilingSubThickness(room))
            WriteLine(1, "floor substrate thickness (mm)", FloorSubThickness(room))
            WriteLine(1, "wall substrate conductivity (W/mK)", WallSubConductivity(room))
            WriteLine(1, "floor substrate conductivity (W/mK)", FloorSubConductivity(room))
            WriteLine(1, "wall substrate specific heat (J/kgK)", WallSubSpecificHeat(room))
            WriteLine(1, "floor substrate specific heat (J/kgK)", FloorSubSpecificHeat(room))
            WriteLine(1, "wall substrate density (kg/m3)", WallSubDensity(room))
            WriteLine(1, "floor substrate density (kg/m3)", FloorSubDensity(room))
            WriteLine(1, "floor specific heat (kJ/kgK)", FloorSpecificHeat(room))
            WriteLine(1, "floor density (kg/m3)", FloorDensity(room))
            WriteLine(1, "ceiling lining specific heat (J/kgK)", CeilingSpecificHeat(room))
            WriteLine(1, "ceiling lining density (kg/m3)", CeilingDensity(room))
            WriteLine(1, "ceiling substrate conductivity (W/mK)", CeilingSubConductivity(room))
            WriteLine(1, "ceiling substrate specific heat (J/kgK)", CeilingSubSpecificHeat(room))
            WriteLine(1, "ceiling substrate density (kg/m3)", CeilingSubDensity(room))
            WriteLine(1, "have ceiling substrate? TRUE/FALSE", HaveCeilingSubstrate(room))
            WriteLine(1, "have wall substrate? TRUE/FALSE", HaveWallSubstrate(room))
            WriteLine(1, "have floor substrate? TRUE/FALSE", HaveFloorSubstrate(room))
            WriteLine(1, "ceiling sloped, 0= flat, -1=sloping", CeilingSlope(room))
            WriteLine(1, "ceiling emissivity", Surface_Emissivity(1, room))
            WriteLine(1, "upper wall emissivity", Surface_Emissivity(2, room))
            WriteLine(1, "lower wall emissivity", Surface_Emissivity(3, room))
            WriteLine(1, "floor emissivity", Surface_Emissivity(4, room))
        Next room

        WriteLine(1, "interior temp (K)", InteriorTemp)
        WriteLine(1, "exterior temp (K)", ExteriorTemp)
        WriteLine(1, "relative humidity", RelativeHumidity)
        WriteLine(1, "tenability monitoring height (m)", MonitorHeight)
        WriteLine(1, "activity level", Activity)
        'WriteLine(1, "radiant loss fraction", RadiantLossFraction)
        'WriteLine(1, "mass loss per unit area (kg/s)", MLUnitArea)
        WriteLine(1, "emission coefficient", EmissionCoefficient)
        WriteLine(1, "simulation time (s)", SimTime)
        WriteLine(1, "display interval (s)", DisplayInterval)
        WriteLine(1, "plume, macaffrey=2, delichatsios=1", plume)
        WriteLine(1, "suppress ceiling HRR", SuppressCeiling)
        WriteLine(1, "flame area constant (m2/kW)", FlameAreaConstant)
        WriteLine(1, "flame length power", FlameLengthPower)
        WriteLine(1, "burner width (m)", BurnerWidth)
        WriteLine(1, "wall heat flux (kW/m2)", WallHeatFlux)
        WriteLine(1, "ceiling heat flux (kW/m2)", CeilingHeatFlux)
        For room = 1 To NumberRooms
            For i = 2 To NumberRooms + 1
                If room < i Then WriteLine(1, "number vents", NumberVents(room, i))
            Next i
        Next room
        Resize_Vents()
        Resize_CVents()
        For room = 1 To NumberRooms
            For i = 2 To NumberRooms + 1
                If room < i Then
                    If NumberVents(room, i) > 0 Then
                        For k = 1 To NumberVents(room, i)
                            WriteLine(1, "Room ", room, " to ", i, " Vent ", k)
                            WriteLine(1, "vent height (m)", VentHeight(room, i, k))
                            WriteLine(1, "vent width (m)", VentWidth(room, i, k))
                            WriteLine(1, "vent sill height (m)", VentSillHeight(room, i, k))
                            WriteLine(1, "vent open time (s)", VentOpenTime(room, i, k))
                            WriteLine(1, "vent close time (s)", VentCloseTime(room, i, k))
                            WriteLine(1, "glass conductivity(s)", GLASSconductivity(room, i, k))
                            WriteLine(1, "glass emissivity(-)", GLASSemissivity(room, i, k))
                            WriteLine(1, "glass linear coefficient of expansion (/C)", GLASSexpansion(room, i, k))
                            WriteLine(1, "glass thickness (mm)", GLASSthickness(room, i, k))
                            WriteLine(1, "glass shading depth (mm)", GLASSshading(room, i, k))
                            WriteLine(1, "glass breaking stress (MPa)", GLASSbreakingstress(room, i, k))
                            WriteLine(1, "glass thermal diffusivity (m2/s)", GLASSalpha(room, i, k))
                            WriteLine(1, "glass Young's modulus (MPa)", GlassYoungsModulus(room, i, k))
                            WriteLine(1, "Auto Break Glass", AutoBreakGlass(room, i, k))
                            WriteLine(1, "Glass fallout time (sec)", GLASSFalloutTime(room, i, k))
                            WriteLine(1, "Glass to flame distance (m)", GLASSdistance(room, i, k))
                            WriteLine(1, "Glass heated hot layer only?", GlassFlameFlux(room, i, k))
                            WriteLine(1, "downstand depth", Downstand(room, i, k))
                            WriteLine(1, "balcony extend beyond compartment opening?", SpillPlumeBalc(room, i, k))
                            WriteLine(1, "Use Spill Plume?", spillplume(room, i, k))
                            WriteLine(1, "Spill Plume Model?", spillplumemodel(room, i, k))
                            WriteLine(1, "Spill Plume Single Sided?", spillbalconyprojection(room, i, k))

                            If i <> NumberRooms + 1 Then
                                WriteLine(1, "wall length 1 (m)", WallLength1(room, i, k))
                                WriteLine(1, "wall length 2 (m)", WallLength2(room, i, k))
                            End If

                            '2009.20
                            WriteLine(1, "auto opening vents?", AutoOpenVent(room, i, k))
                            WriteLine(1, "fire size for human detection", HRR_threshold(room, i, k))
                            WriteLine(1, "time to reach fire size threshold", HRR_threshold_time(room, i, k))
                            WriteLine(1, "flag for HRR threshold", HRRFlag(room, i, k))
                            WriteLine(1, "vent opening delay time", trigger_ventopendelay(room, i, k))
                            WriteLine(1, "vent opening duration", trigger_ventopenduration(room, i, k))
                            WriteLine(1, "hrr vent open delay time", HRR_ventopendelay(room, i, k))
                            WriteLine(1, "hrr vent open duration", HRR_ventopenduration(room, i, k))
                            WriteLine(1, "room for SD trigger", SDtriggerroom(room, i, k))
                            WriteLine(1, "trigger device id", trigger_device(0, room, i, k))
                            WriteLine(1, "trigger device id", trigger_device(1, room, i, k))
                            WriteLine(1, "trigger device id", trigger_device(2, room, i, k))
                            WriteLine(1, "trigger device id", trigger_device(3, room, i, k))
                            WriteLine(1, "trigger device id", trigger_device(4, room, i, k))
                            WriteLine(1, "trigger device id", trigger_device(5, room, i, k))
                            WriteLine(1, "trigger device id", trigger_device(6, room, i, k))
                        Next k
                    End If
                End If
            Next i
        Next room

        WriteLine(1, "number objects", NumberObjects)

        Resize_Objects()

        For k = 1 To NumberObjects
            WriteLine(1, "number data points", NumberDataPoints(k))
            WriteLine(1, "energy yield (kJ/g)", EnergyYield(k))
            WriteLine(1, "CO yield (g/g)", COYield)
            WriteLine(1, "CO2 yield (g/g)", CO2Yield(k))
            WriteLine(1, "soot yield (g/g)", SootYield(k))
            WriteLine(1, "water vapour yield (g/g)", WaterVaporYield(k))
            WriteLine(1, "Fire height (m)", FireHeight(k))
            WriteLine(1, "fire location, corner=2, wall=1, centre=0", FireLocation(k))
            WriteLine(1, "HCN yield", HCNuserYield(k))
            WriteLine(1, "obj CRF pilot", ObjCRF(k))
            WriteLine(1, "obj FTP index pilot", ObjFTPindexpilot(k))
            WriteLine(1, "obj FTP limit pilot", ObjFTPlimitpilot(k))
            WriteLine(1, "obj CRF auto", ObjCRFauto(k))
            WriteLine(1, "obj FTP index auto", ObjFTPindexauto(k))
            WriteLine(1, "obj FTP limit auto", ObjFTPlimitauto(k))
            WriteLine(1, "obj length", ObjLength(k))
            WriteLine(1, "obj width", ObjWidth(k))
            WriteLine(1, "obj height", ObjHeight(k))
            WriteLine(1, "obj x", ObjDimX(k))
            WriteLine(1, "obj y", ObjDimY(k))
            WriteLine(1, "obj elevation", ObjElevation(k))
            WriteLine(1, "HRR data")
            For i = 1 To NumberDataPoints(k)
                If i < 501 Then WriteLine(1, HeatReleaseData(1, i, k), HeatReleaseData(2, i, k))
            Next i
        Next k

        WriteLine(1, "Detector Type", DetectorType(fireroom))
        WriteLine(1, "RTI", RTI(fireroom))
        WriteLine(1, "C-factor", cfactor(fireroom))
        WriteLine(1, "radial distance (m)", RadialDistance(fireroom))
        WriteLine(1, "actuation temp (K)", ActuationTemp(fireroom))
        WriteLine(1, "water discharge rate", WaterSprayDensity(fireroom))
        'Write #1, "sprinkler setting", frmSprink.optControl.Value, frmSprink.optSprinkOff.Value, frmSprink.optSuppression.Value
        'WriteLine(1, "sprinkler setting", Sprinkler(1, fireroom), Sprinkler(2, fireroom), Sprinkler(3, fireroom))
        WriteLine(1, "sprink_mode", sprink_mode)
        WriteLine(1, "target radiation endpoint (kW/m2)", TargetEndPoint)
        WriteLine(1, "upper temp endpoint (K)", TempEndPoint)
        WriteLine(1, "visibility endpoint (m)", VisibilityEndPoint)
        WriteLine(1, "FED endpoint", FEDEndPoint)
        WriteLine(1, "convective endpoint (K)", ConvectEndPoint)
        For room = 1 To NumberRooms
            WriteLine(1, WallConeDataFile(room))
            WriteLine(1, FloorConeDataFile(room))
            WriteLine(1, CeilingConeDataFile(room))
            WriteLine(1, "wall min temp for spread (k)", WallTSMin(room))
            WriteLine(1, "wall flame spread parameter", WallFlameSpreadParameter(room))
            WriteLine(1, "wall effective heat of combustion", WallEffectiveHeatofCombustion(room))
            WriteLine(1, "ceiling effective heat of combustion", CeilingEffectiveHeatofCombustion(room))
            WriteLine(1, "floor effective heat of combustion", FloorEffectiveHeatofCombustion(room))
            WriteLine(1, "fan extract rate (m3/s)", ExtractRate(room))
            WriteLine(1, "fan start time (sec)", ExtractStartTime(room))
            WriteLine(1, "fan on?", fanon(room))
            WriteLine(1, "Max Pressure (Pa)", MaxPressure(room))
            WriteLine(1, "Extract?", Extract(room))
            WriteLine(1, "Number Fans", NumberFans(room))
            WriteLine(1, "Wall Soot Yield", WallSootYield(room))
            WriteLine(1, "Ceiling Soot Yield", CeilingSootYield(room))
            WriteLine(1, "Floor Soot Yield", FloorSootYield(room))
            WriteLine(1, "Wall CO2 Yield", WallCO2Yield(room))
            WriteLine(1, "Ceiling CO2 Yield", CeilingCO2Yield(room))
            WriteLine(1, "Floor CO2 Yield", FloorCO2Yield(room))
            WriteLine(1, "Wall H2O Yield", WallH2OYield(room))
            WriteLine(1, "Ceiling H2O Yield", CeilingH2OYield(room))
            WriteLine(1, "Floor H2O Yield", FloorH2OYield(room))
            WriteLine(1, "Floor min temp for spread (k)", FloorTSMin(room))
            WriteLine(1, "Floor flame spread parameter", FloorFlameSpreadParameter(room))
            'Write #1, "Floor HCN Yield"; FloorHCNYield(room)
            'Write #1, "Wall HCN Yield"; WallHCNYield(room)
            'Write #1, "Ceiling HCN Yield"; CeilingHCNYield(room)
        Next room
        WriteLine(1, "fire in room", fireroom)
        If NumberRooms > 1 Then WriteLine(1, "Ignite adjacent rooms?", IgniteNextRoom)
        WriteLine(1, "FED Start time", StartOccupied)
        WriteLine(1, "FED end time", EndOccupied)
        WriteLine(1, "Illuminated signage", illumination)
        For room = 1 To NumberRooms + 1
            For i = 1 To NumberRooms + 1
                If room <> i Then WriteLine(1, "number cVents", NumberCVents(room, i))
            Next i
        Next room
        For room = 1 To NumberRooms + 1
            For i = 1 To NumberRooms + 1
                If room <> i Then
                    For k = 1 To NumberCVents(room, i)
                        If NumberCVents(room, i) > 0 Then
                            WriteLine(1, "Upper space ", room, "lower space ", i, "Vent ", k)
                            WriteLine(1, "cVent Area", CVentArea(room, i, k))
                            WriteLine(1, "cVent Opening Time", CVentOpenTime(room, i, k))
                            WriteLine(1, "cVent Closing Time", CVentCloseTime(room, i, k))
                            WriteLine(1, "cVent Auto opening?", CVentAuto(room, i, k))
                        End If
                    Next k
                End If
            Next i
        Next room
        For room = 1 To NumberRooms
            WriteLine(1, "Use fan curve?", UseFanCurve(room))
            WriteLine(1, "Fan Elevation", FanElevation(room))
        Next room
        'End If
        WriteLine(1, "Ceiling Nodes", Ceilingnodes)
        WriteLine(1, "Wall Nodes", Wallnodes)
        WriteLine(1, "Floor Nodes", Floornodes)
        WriteLine(1, "LE Solver", LEsolver)
        'If frmOptions1.optLUdecom.Checked = True Then
        ' WriteLine(1, "LE solver", "LU decomposition")
        ' Else
        'WriteLine(1, "LE solver", "Gauss-Jordan")
        'End If
        WriteLine(1, "Enhanced Burning Rate", Enhance)
        WriteLine(1, "Job Number", JobNumber)
        WriteLine(1, "Excel Interval (s)", ExcelInterval)
        For room = 1 To NumberRooms
            WriteLine(1, "Two Zones? ", TwoZones(room))
        Next room
        WriteLine(1, "Time Step", Timestep)
        WriteLine(1, "Error Control", Error_Control)
        WriteLine(1, "Fire Objects Database", FireDatabaseName)
        WriteLine(1, "Materials Database", MaterialsDatabaseName)
        'Write #1, "Fire Objects Database", Dir(FireDatabaseName)
        'Write #1, "Materials Database", Dir(MaterialsDatabaseName)
        For room = 1 To NumberRooms
            WriteLine(1, "Have Smoke Detector?", HaveSD(room))
            WriteLine(1, "Alarm OD", SmokeOD(room))
            WriteLine(1, "Alarm delay", SDdelay(room))
            WriteLine(1, "Detector Sensitivity", DetSensitivity(room))
            WriteLine(1, "Radial Distance", SDRadialDist(room))
            WriteLine(1, "Depth", SDdepth(room))
            WriteLine(1, "Use OD inside detector for response", SDinside(room))
            WriteLine(1, "Fan Auto Start?", FanAutoStart(room))
            WriteLine(1, "Specify Alarm OD?", SpecifyOD(room))
        Next room
        WriteLine(1, "Ceiling Jet Model", cjModel)
        WriteLine(1, "Use One Cone Curve Only?", UseOneCurve)
        WriteLine(1, "Ignition Correlation", IgnCorrelation)
        WriteLine(1, "Sprinkler Distance", SprinkDistance(fireroom))
        WriteLine(1, "Vent Log File", ventlog)
        WriteLine(1, "Underventilated Soot Yield Factor", SootFactor)
        WriteLine(1, "Postflashover Model", post)
        WriteLine(1, "FLED", FLED)

        WriteLine(1, "Fuel Thickness", Fuel_Thickness)
        WriteLine(1, "Heat of Combustion", HoC_fuel)
        WriteLine(1, "Stick Spacing", Stick_Spacing)
        WriteLine(1, "Crib Height", Cribheight)
        WriteLine(1, "Excess Fuel Factor", ExcessFuelFactor)
        WriteLine(1, "Soot Alpha Coefficient", SootAlpha)
        WriteLine(1, "Soot Epsilon Coefficient", SootEpsilon)
        WriteLine(1, "Carbon atoms in fuel", nC)
        WriteLine(1, "Hydrogen atoms in fuel", nH)
        WriteLine(1, "Oxygen atoms in fuel", nO)
        WriteLine(1, "Nitrogen atoms in fuel", nN)

        WriteLine(1, "fuel type", fueltype)
        WriteLine(1, "Disable wall flow", nowallflow)
        WriteLine(1, "Calculate HCN yield", HCNcalc)
        WriteLine(1, "preflashover CO yield", preCO)
        WriteLine(1, "postflashover CO yield", postCO)
        WriteLine(1, "preflashover soot yield", preSoot)
        WriteLine(1, "postflashover soot yield", postSoot)
        WriteLine(1, "CO mode", comode)
        WriteLine(1, "soot mode", sootmode)
        WriteLine(1, "user mode", MDIFrmMain.NZBCVM2ToolStripMenuItem.Checked)

        FileClose(1)
    End Sub
    Public Sub Save_plt()
        ''save simulation output to plt file

        'Dim room, i As Integer

        'Dim outputfile As String = DataFile.Replace(".xml", ".plt")
        ''outputfile = "C:\Documents and Settings\branzcw\My Documents\BRANZFIRE-RISK\data\test.plt"

        '' Split the string on the . character
        'Dim parts As String() = DataFile.Split(New Char() {"."c})
        'outputfile = parts(0) & ".plt"


        'Dim binaryout As New BinaryWriter(New FileStream(outputfile, FileMode.Create, FileAccess.Write))

        'binaryout.Write(CInt(1)) 'version
        'binaryout.Write(NumberRooms)
        'binaryout.Write(CInt(1)) 'number of fires

        'For i = 1 To NumberTimeSteps + 1

        '    binaryout.Write(tim(i, 1))
        '    For room = 1 To NumberRooms
        '        binaryout.Write(RoomPressure(room, i))
        '        binaryout.Write(layerheight(room, i))
        '        binaryout.Write(lowertemp(room, i))
        '        binaryout.Write(uppertemp(room, i))
        '    Next

        '    binaryout.Write(1.0) 'should be flame height i think
        '    binaryout.Write(HeatRelease(fireroom, i, 2))

        'Next

        'binaryout.Close()

    End Sub
    Public Sub Save_Smokeview(ByVal filename As String)

        Dim room As Integer
        Dim searchfolder As String = ""
        Dim str1, str2, str3, str4, str5, str6, str7 As String
        Dim yfolder As String = ""
        ' Dim oFile As System.IO.File
        ' Dim oWrite As System.IO.StreamWriter
        'Dim oFile As File
        Dim oWrite As StreamWriter

        'oWrite = oFile.CreateText(filename)
        oWrite = File.CreateText(filename)

        Try
            Dim part As String
            part = Path.GetFileName(filename)
            part = part.Replace(".smv", "_zone.csv")

            'Dim parts As String() = part.Split(New Char() {"."c})
            'part = parts(0) & "_zone.csv"

            oWrite.WriteLine("{0,1}", "ZONE")
            oWrite.WriteLine("{0,1}", part)
            oWrite.WriteLine("{0,1}", "P")
            oWrite.WriteLine("{0,1}", "Pa")
            oWrite.WriteLine("{0,1}", "Layer Height")
            oWrite.WriteLine("{0,1}", "ylay")
            oWrite.WriteLine("{0,1}", "m")
            oWrite.WriteLine("{0,1}", "TEMPERATURE")
            oWrite.WriteLine("{0,1}", "TEMP")
            oWrite.WriteLine("{0,1}", "C")
            oWrite.WriteLine("{0,1}", "TEMPERATURE")
            oWrite.WriteLine("{0,1}", "TEMP")
            oWrite.WriteLine("{0,1}", "C")
            oWrite.WriteLine("{0,1}", "AMBIENT")
            str1 = String.Format("{0:E}", Atm_Pressure * 1000)
            str2 = String.Format("{0:E}", 0)
            str3 = String.Format("{0:E}", InteriorTemp)

            oWrite.WriteLine("{0,15}{1,15}{2,15}", str1, str2, str3)

            For room = 1 To NumberRooms
                str1 = String.Format("{0:d}", room)
                str2 = String.Format("{0}", "ROOM")
                oWrite.WriteLine("{1,4}{0,4}", str1, str2)
                str1 = String.Format("{0:E4}", RoomLength(room))
                str2 = String.Format("{0:E4}", RoomWidth(room))
                str3 = String.Format("{0:E4}", RoomHeight(room))
                oWrite.WriteLine("{0,12}{1,12}{2,12}", str1, str2, str3)
                str1 = String.Format("{0:E4}", RoomAbsX(room))
                str2 = String.Format("{0:E4}", RoomAbsY(room))
                str3 = String.Format("{0:E4}", FloorElevation(room))
                oWrite.WriteLine("{0,12}{1,12}{2,12}", str1, str2, str3)
            Next

            For room = 1 To NumberRooms
                str2 = String.Format("{0}", "LABEL")
                oWrite.WriteLine("{0,4}", str2)
                oWrite.WriteLine("{0,12}{1,12}{2,12}", RoomAbsX(room) + RoomLength(room) / 4, RoomAbsY(room) + RoomWidth(room) / 4, FloorElevation(room), 0, 0, 0, 0, SimTime)
                str1 = String.Format("{0:d}", RoomDescription(room))
                oWrite.WriteLine("{0,4}", str1)
            Next

            For room = 1 To NumberRooms
                For i = 2 To NumberRooms + 1
                    If room < i Then
                        If NumberVents(room, i) > 0 Then
                            For k = 1 To NumberVents(room, i)
                                oWrite.WriteLine("{0,1}", "VENTGEOM")
                                str1 = String.Format("{0:d}", room)
                                str2 = String.Format("{0:d}", i)
                                'vent face 0 = front, length dim
                                'ventface 1 = right, width dim
                                'vent face 2 = rear, length dim
                                'vent face 3 = left, width dim
                                str3 = String.Format("{0:d}", VentFace(room, i, k) + 1)  'face 
                                str4 = String.Format("{0:E4}", VentWidth(room, i, k))

                                str5 = String.Format("{0:E4}", VentOffset(room, i, k)) 'offset


                                'If WallLength1(room, i, k) < WallLength2(room, i, k) Or i = NumberRooms + 1 Then
                                '    If VentFace(room, i, k) = 1 Then
                                '        str5 = String.Format("{0:E4}", RoomWidth(room) / 2 - VentWidth(room, i, k) / 2) 'offset
                                '    ElseIf VentFace(room, i, k) = 0 Then
                                '        str5 = String.Format("{0:E4}", RoomLength(room) / 2 - VentWidth(room, i, k) / 2) 'offset
                                '    ElseIf VentFace(room, i, k) = 2 Then
                                '        str5 = String.Format("{0:E4}", RoomLength(room) / 2 - VentWidth(room, i, k) / 2) 'offset
                                '    ElseIf VentFace(room, i, k) = 3 Then
                                '        str5 = String.Format("{0:E4}", RoomWidth(room) / 2 - VentWidth(room, i, k) / 2) 'offset
                                '    End If
                                'Else
                                '    If VentFace(room, i, k) = 1 Then
                                '        str5 = String.Format("{0:E4}", RoomWidth(i) / 2 - VentWidth(room, i, k) / 2) 'offset
                                '    ElseIf VentFace(room, i, k) = 0 Then
                                '        str5 = String.Format("{0:E4}", RoomLength(i) / 2 - VentWidth(room, i, k) / 2) 'offset
                                '    ElseIf VentFace(room, i, k) = 2 Then
                                '        str5 = String.Format("{0:E4}", RoomLength(i) / 2 - VentWidth(room, i, k) / 2) 'offset
                                '    ElseIf VentFace(room, i, k) = 3 Then
                                '        str5 = String.Format("{0:E4}", RoomWidth(i) / 2 - VentWidth(room, i, k) / 2) 'offset
                                '    End If
                                'End If

                                str6 = String.Format("{0:E4}", VentSillHeight(room, i, k))
                                str7 = String.Format("{0:E4}", VentSillHeight(room, i, k) + VentHeight(room, i, k))
                                oWrite.WriteLine("{0,4}{1,4}{2,4}{3,12}{4,12}{5,12}{6,12}", str1, str2, str3, str4, str5, str6, str7)

                            Next
                        End If
                    End If
                Next
            Next

            If FireHeight.GetUpperBound(0) < NumberObjects Then
                Resize_Objects()
            End If
            For i = 1 To NumberObjects
                oWrite.WriteLine("{0,1}", "FIRE")
                str1 = String.Format("{0:d}", fireroom)
                str2 = String.Format("{0:E4}", CSng(ObjDimX(i) + ObjLength(i) / 2))
                str3 = String.Format("{0:E4}", CSng(ObjDimY(i) + ObjWidth(i) / 2))
                str4 = String.Format("{0:E4}", FloorElevation(fireroom) + FireHeight(i))
                oWrite.WriteLine("{0,4}{1,12}{2,12}{3,12}", str1, str2, str3, str4)
            Next



            For room = 1 To NumberRooms + 1 'lower room
                'For i = 2 To NumberRooms + 1 'upper room
                For i = 1 To NumberRooms + 1 'upper room
                    'If room < i Then

                    If NumberCVents(i, room) > 0 Then
                        If room < NumberRooms + 1 Then
                            For k = 1 To NumberCVents(i, room)
                                oWrite.WriteLine("{0,1}", "VFLOWGEOM")
                                str1 = String.Format("{0:d}", room)
                                str2 = String.Format("{0:d}", i)
                                str3 = String.Format("{0:d}", 6)  'face 
                                str4 = String.Format("{0:E4}", CVentArea(i, room, k)) 'area
                                str5 = String.Format("{0:d}", 1)  'shape 
                                oWrite.WriteLine("{0,4}{1,4}{2,4}{3,12}{4,4}", str1, str2, str3, str4, str5)
                            Next
                        Else
                            'csn't get vent to display in the floor
                            For k = 1 To NumberCVents(i, room)
                                oWrite.WriteLine("{0,1}", "VFLOWGEOM")
                                str1 = String.Format("{0:d}", room)
                                str2 = String.Format("{0:d}", i)
                                str3 = String.Format("{0:d}", 1)  'face 
                                str4 = String.Format("{0:E4}", CVentArea(i, room, k)) 'area
                                str5 = String.Format("{0:d}", 1)  'shape 
                                oWrite.WriteLine("{0,4}{1,4}{2,4}{3,12}{4,4}", str1, str2, str3, str4, str5)
                            Next
                        End If

                    End If
                Next
            Next

            If NumberObjects > 0 Then
                oWrite.WriteLine("{0,1}", "THCP")
                str1 = String.Format("{0:d}", NumberObjects)
                oWrite.WriteLine("{0,4}", str1)
                For i = 1 To NumberObjects
                    str2 = String.Format("{0:F3}", RoomAbsX(fireroom) + ObjDimX(i) + ObjLength(i) / 2)
                    str3 = String.Format("{0:F3}", RoomAbsY(fireroom) + ObjDimY(i) + ObjWidth(i) / 2)
                    str4 = String.Format("{0:F1}", 0)
                    str5 = String.Format("{0:F1}", 0)
                    str6 = String.Format("{0:F1}", 0)
                    str7 = String.Format("{0:F1}", 1)
                    oWrite.WriteLine("{0,7}{1,10}{2,10}{3,10}{4,10}{5,10}", str2, str3, str4, str5, str6, str7)
                Next
            End If

            oWrite.WriteLine("{0,1}", "TIME")
            str1 = String.Format("{0:d0}", SimTime) & "."
            oWrite.WriteLine("{0,7}{1,11}", NumberTimeSteps, str1)

            Dim device_counter As Integer = 0
            oWrite.WriteLine("{0,1}", "DEVICE")
            device_counter = device_counter + 1
            oWrite.WriteLine("{0,7}", "TIME")
            str1 = String.Format("{0:F1}", 0)
            oWrite.WriteLine("{0,6}{1,6}{2,6}", str1, str1, str1)

            For i = 1 To NumberRooms
                oWrite.WriteLine("{0,1}", "DEVICE")
                device_counter = device_counter + 1
                str1 = "ULT_" & i.ToString
                oWrite.WriteLine("{0,7}", str1)
                str1 = String.Format("{0:F1}", 0)
                oWrite.WriteLine("{0,6}{1,6}{2,6}", str1, str1, str1)

                oWrite.WriteLine("{0,1}", "DEVICE")
                device_counter = device_counter + 1
                str1 = "LLT_" & i.ToString
                oWrite.WriteLine("{0,7}", str1)
                str1 = String.Format("{0:F1}", 0)
                oWrite.WriteLine("{0,6}{1,6}{2,6}", str1, str1, str1)

                oWrite.WriteLine("{0,1}", "DEVICE")
                device_counter = device_counter + 1
                str1 = "HGT_" & i.ToString
                oWrite.WriteLine("{0,7}", str1)
                str1 = String.Format("{0:F1}", 0)
                oWrite.WriteLine("{0,6}{1,6}{2,6}", str1, str1, str1)

                oWrite.WriteLine("{0,1}", "DEVICE")
                device_counter = device_counter + 1
                str1 = "PRS_" & i.ToString
                oWrite.WriteLine("{0,7}", str1)
                str1 = String.Format("{0:F1}", 0)
                oWrite.WriteLine("{0,6}{1,6}{2,6}", str1, str1, str1)

                oWrite.WriteLine("{0,1}", "DEVICE")
                device_counter = device_counter + 1
                str1 = "ULOD_" & i.ToString
                oWrite.WriteLine("{0,7}", str1)
                str1 = String.Format("{0:F1}", 0)
                oWrite.WriteLine("{0,6}{1,6}{2,6}", str1, str1, str1)

                oWrite.WriteLine("{0,1}", "DEVICE")
                device_counter = device_counter + 1
                str1 = "LLOD_" & i.ToString
                oWrite.WriteLine("{0,7}", str1)
                str1 = String.Format("{0:F1}", 0)
                oWrite.WriteLine("{0,6}{1,6}{2,6}", str1, str1, str1)
            Next

            If NumberObjects > 0 Then

                For i = 1 To NumberObjects
                    oWrite.WriteLine("{0,1}", "DEVICE")
                    device_counter = device_counter + 1
                    str1 = "HRR_" & i.ToString
                    oWrite.WriteLine("{0,7}", str1)
                    str1 = String.Format("{0:F1}", 0)
                    oWrite.WriteLine("{0,6}{1,6}{2,6}", str1, str1, str1)

                    oWrite.WriteLine("{0,1}", "DEVICE")
                    device_counter = device_counter + 1
                    str1 = "FLHGT_" & i.ToString
                    oWrite.WriteLine("{0,7}", str1)
                    str1 = String.Format("{0:F1}", 0)
                    oWrite.WriteLine("{0,6}{1,6}{2,6}", str1, str1, str1)

                    oWrite.WriteLine("{0,1}", "DEVICE")
                    device_counter = device_counter + 1
                    str1 = "FBASE_" & i.ToString
                    oWrite.WriteLine("{0,7}", str1)
                    str1 = String.Format("{0:F1}", 0)
                    oWrite.WriteLine("{0,6}{1,6}{2,6}", str1, str1, str1)

                    oWrite.WriteLine("{0,1}", "DEVICE")
                    device_counter = device_counter + 1
                    str1 = "FAREA_" & i.ToString
                    oWrite.WriteLine("{0,7}", str1)
                    str1 = String.Format("{0:F1}", 0)
                    oWrite.WriteLine("{0,6}{1,6}{2,6}", str1, str1, str1)
                Next
            End If

            For i = 1 To number_vents
                oWrite.WriteLine("{0,1}", "DEVICE")
                device_counter = device_counter + 1
                str1 = "HVENT_" & i.ToString
                oWrite.WriteLine("{0,7}", str1)
                str1 = String.Format("{0:F1}", 0)
                oWrite.WriteLine("{0,6}{1,6}{2,6}", str1, str1, str1)
            Next

            Dim m As Integer = 0
            For room = 1 To NumberRooms + 1
                For i = 1 To NumberRooms + 1
                    If NumberCVents(i, room) > 0 Then
                        For k = 1 To NumberCVents(i, room)
                            m = m + 1
                            oWrite.WriteLine("{0,1}", "DEVICE")
                            device_counter = device_counter + 1
                            str1 = "VVENT_" & m.ToString
                            oWrite.WriteLine("{0,7}", str1)
                            str1 = String.Format("{0:F1}", 0)
                            oWrite.WriteLine("{0,6}{1,6}{2,6}", str1, str1, str1)
                        Next
                    End If
                Next
            Next

            Dim oSprinklers As New List(Of oSprinkler)
            oSprinklers = SprinklerDB.GetSprinklers2
            Dim osprdistributions As New List(Of oDistribution)
            osprdistributions = SprinklerDB.GetSprDistributions

            For Each oSprinkler In oSprinklers
                For Each oDistribution In osprdistributions
                    If oDistribution.id = oSprinkler.sprid Then
                        If oDistribution.varname = "sprz" Then
                            oSprinkler.sprz = oDistribution.varvalue
                            Exit For
                        End If
                    End If
                Next

                oWrite.WriteLine("{0,1}", "DEVICE")
                device_counter = device_counter + 1
                str1 = "sprinkler_upright"
                oWrite.WriteLine("{0,7}", str1)
                str1 = String.Format("{0:F5}", 0)
                str2 = String.Format("{0:F5}", 1)
                str3 = String.Format("{0:F5}", RoomAbsX(oSprinkler.room) + oSprinkler.sprx)
                str4 = String.Format("{0:F5}", RoomAbsY(oSprinkler.room) + oSprinkler.spry)
                str5 = String.Format("{0:F5}", FloorElevation(oSprinkler.room) + RoomHeight(oSprinkler.room) - oSprinkler.sprz)
                oWrite.WriteLine("{0,12}{1,12}{2,12}{3,12}{4,12}{5,12}{6,6}", str3, str4, str5, str1, str1, str2, "  0  0 % null")

                oWrite.WriteLine("{0,1}", "DEVICE_ACT") 'not working?
                str1 = String.Format("{0:F0}", device_counter) 'idevice
                str2 = String.Format("{0:F1}", oSprinkler.responsetime) 'time
                str3 = String.Format("{0:F0}", 1) 'state
                oWrite.WriteLine("{0,12}{1,12}{2,12}", str1, str2, str3, "  0  0 % null")

            Next

            oWrite.Close()

            If filename.Length > 4 Then
                Dim myProcess As System.Diagnostics.Process = New System.Diagnostics.Process()

                ' make a reference to a directory

                'searchfolder = FileIO.SpecialDirectories.ProgramFiles
                'Dim di As New IO.DirectoryInfo(searchfolder & "\FDS\FDS5")
                ''Dim diar1 As IO.FileInfo() = di.GetFiles("smokeview.exe")
                ''Dim dra As IO.FileInfo
                'Dim found As Boolean

                'For Each y In di.GetDirectories
                '    yfolder = di.FullName
                '    For Each x In y.GetFiles

                '        Dim name As String = x.FullName
                '        If InStr(name, "smokeview") And InStr(name, ".exe") Then
                '            found = True
                '            myProcess.StartInfo.FileName = name
                '            Exit For
                '        End If
                '    Next
                '    If found = True Then Exit For
                'Next

                myProcess.StartInfo.Arguments = filename
                myProcess.StartInfo.WorkingDirectory = RiskDataDirectory
                myProcess.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal
                myProcess.StartInfo.FileName = filename
                myProcess.Start(ApplicationPath & "\smokeview.exe", filename)
                myProcess.Close()
            End If
            Exit Sub

        Catch ex As Exception

            MsgBox(ex.Message & " This feature requires Smokeview 6.0 or later to be installed. You can download this program from http://fire.nist.gov/fds/downloads.html", MsgBoxStyle.Critical)
        End Try

        oWrite.Close()
    End Sub

    Public Sub Flame_Spread_Props(ByRef room As Integer, ByRef id As Integer, ByRef high As Single, ByRef ConeDataFile As Object, ByRef emissivity As Double, ByRef ThermalInertia As Double, ByRef IgnitionTemp As Double, ByRef EffectiveHeatofCombustion As Double, ByRef HeatofGasification As Double, ByRef AreaUnderCurve As Double, ByRef FTP As Double, ByRef QCritical As Double, ByRef Nmax As Single)
        '========================================================================
        '   find flame spread properties
        '   4 April 1998
        '========================================================================
        Dim radflux(20) As Double
        Dim IgnitionTime(20) As Double
        Dim Peak(20) As Double
        Dim IgnPoints As Short
        Dim slope, IO, slopemax As Double
        Dim Ignition(20) As Double
        Dim Yintercept, RMax, R2 As Double
        Dim n As Single
        Dim i As Short
        'Dim caption As String

        AreaUnderCurve = 0
        IgnitionTemp = 0
        ThermalInertia = 0
        QCritical = 0

        If ConeDataFile = "null.txt" Then
            Exit Sub
        End If

        Try


            'read data from cone
            Call Read_Cone_Data(room, id, high, ConeDataFile, radflux, IgnitionTime, IgnPoints, Peak, IgnitionTemp, ThermalInertia, HeatofGasification, AreaUnderCurve, QCritical)

            If frmQuintiere.optUseOneCurve.Checked = True Then
                'normalise the cone data
                Call Normalise_Data(id, room)
            End If
            If AreaUnderCurve <= 0 Then
                'find the area under the cone hrr curve if the user has not specified the value
                Call Integrate_HRR(room, id, 0, high, IO)
                AreaUnderCurve = IO
                If IO = 0 Then MsgBox("Problem integrating the area under the cone data curve.")
            End If

            If IgnPoints > 1 Then
                'determine ignition temperature and thermal inertia from a piloted ignition data

                RMax = 0
                Nmax = 0

                If IgnCorrelation = vbJanssens Then
                    For n = 0.547 To 1 Step 0.001
                        For i = 1 To IgnPoints
                            If IgnitionTime(i) > 0 Then
                                Ignition(i) = (1 / IgnitionTime(i)) ^ n
                            End If
                        Next

                        'fit a linear regression line to the data, plotting vs external flux
                        Call RegressionL(radflux, Ignition, IgnPoints, Yintercept, slope, R2)

                        If R2 * R2 > RMax Then
                            RMax = R2 * R2
                            Nmax = n
                            slopemax = slope
                        End If
                    Next
                Else 'FTP method
                    For n = 1 To 0.5 Step -0.01
                        For i = 1 To IgnPoints
                            If IgnitionTime(i) > 0 Then
                                Ignition(i) = (1 / IgnitionTime(i)) ^ n
                            End If
                        Next

                        'fit a linear regression line to the data, plotting vs external flux
                        Call RegressionL(radflux, Ignition, IgnPoints, Yintercept, slope, R2)

                        If CDbl(VB6.Format(R2 * R2, "0.00")) > RMax Then
                            'If R2 * R2 > Format(RMax, "0.00") Then
                            RMax = CDbl(VB6.Format(R2 * R2, "0.00"))
                            Nmax = n
                            slopemax = slope
                        End If
                    Next
                End If

                FTP = (1 / slopemax) ^ (1 / Nmax) 'use this even when janssens method used
                'sensitivity analysis
                'FTP = 1.4 * FTP
                For i = 1 To IgnPoints
                    If IgnitionTime(i) > 0 Then
                        Ignition(i) = (1 / IgnitionTime(i)) ^ Nmax
                    End If
                Next

                Call Ignition_Correlation(radflux, Ignition, IgnitionTime, IgnPoints, emissivity, ThermalInertia, IgnitionTemp, QCritical)

                'hardwire Qcritical for sensitivity analysis
                'QCritical = 0.5 * QCritical

                'determine effective heat of gasification using Quintiere's method
                If frmQuintiere.optUseOneCurve.Checked = True Then
                    Call EHoG_Correlation(radflux, Peak, IgnPoints, EffectiveHeatofCombustion, HeatofGasification)
                End If
                Nmax = 1 / Nmax
            Else
                FTP = 0
            End If

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in Flame_Spread_Props")
        End Try
    End Sub
    Public Sub Flame_Spread_Props2(ByRef room As Integer, ByRef id As Integer, ByRef high As Single, ByRef ConeDataFile As Object, ByRef emissivity As Double, ByRef ThermalInertia As Double, ByRef IgnitionTemp As Double, ByRef EffectiveHeatofCombustion As Double, ByRef HeatofGasification As Double, ByRef AreaUnderCurve As Double, ByRef FTP As Double, ByRef QCritical As Double, ByRef Nmax As Single)
        '========================================================================
        '   find flame spread properties
        '   
        '========================================================================
        Dim radflux(20) As Double
        Dim IgnitionTime(20) As Double
        Dim Peak(20) As Double
        Dim IgnPoints As Short
        Dim slope, IO, slopemax, interceptmax, mintestflux As Double
        Dim Ignition(20) As Double
        Dim Yintercept, RMax, R2 As Double
        Dim n As Single
        Dim i As Short
        'Dim caption As String

        AreaUnderCurve = 0
        IgnitionTemp = 0
        ThermalInertia = 0
        QCritical = 0

        If ConeDataFile = "null.txt" Then
            Exit Sub
        End If

        Try

            'read data from cone
            Call Read_Cone_Data(room, id, high, ConeDataFile, radflux, IgnitionTime, IgnPoints, Peak, IgnitionTemp, ThermalInertia, HeatofGasification, AreaUnderCurve, QCritical)

            'find min test flux
            mintestflux = radflux(1)

            If frmQuintiere.optUseOneCurve.Checked = True Then
                'normalise the cone data
                Call Normalise_Data(id, room)
            End If
            If AreaUnderCurve <= 0 Then
                'find the area under the cone hrr curve if the user has not specified the value
                Call Integrate_HRR(room, id, 0, high, IO)
                AreaUnderCurve = IO
                If IO = 0 Then MsgBox("Problem integrating the area under the cone data curve.")
            End If

            If IgnPoints > 1 Then
                'determine ignition temperature and thermal inertia from a piloted ignition data

                RMax = 0
                Nmax = 0

                If IgnCorrelation = vbJanssens Then
                    'For n = 0.547 To 1 Step 0.001
                    '    For i = 1 To IgnPoints
                    '        If IgnitionTime(i) > 0 Then
                    '            Ignition(i) = (1 / IgnitionTime(i)) ^ n
                    '        End If
                    '    Next

                    '    'fit a linear regression line to the data, plotting vs external flux
                    '    Call RegressionL(radflux, Ignition, IgnPoints, Yintercept, slope, R2)

                    '    If R2 * R2 > RMax Then
                    '        RMax = R2 * R2
                    '        Nmax = n
                    '        slopemax = slope
                    '    End If
                    'Next
                Else 'FTP method
                    ' For n = 1 To 0.5 Step -0.01
                    For n = 0.5 To 3 Step 0.1
                        n = CDec(n)
                        For i = 1 To IgnPoints
                            If IgnitionTime(i) > 0 Then
                                Ignition(i) = 1 / (IgnitionTime(i) ^ (1 / n))
                            End If
                        Next

                        'fit a linear regression line to the data, plotting vs external flux
                        Call RegressionL(Ignition, radflux, IgnPoints, Yintercept, slope, R2)

                        If mintestflux < Yintercept Then
                            'qcrit must be greater than yintercept
                            Continue For
                        End If

                        If CDbl(VB6.Format(R2 * R2, "0.00")) > RMax Then
                            'If R2 * R2 > Format(RMax, "0.00") Then
                            RMax = CDbl(VB6.Format(R2 * R2, "0.00"))
                            Nmax = n
                            slopemax = slope
                            interceptmax = Yintercept
                        End If


                    Next
                End If

                ' FTP = (1 / slopemax) ^ (1 / Nmax) 'use this even when janssens method used
                FTP = (slopemax) ^ (Nmax)
                QCritical = interceptmax

                For i = 1 To IgnPoints
                    If IgnitionTime(i) > 0 Then
                        'Ignition(i) = (1 / IgnitionTime(i)) ^ Nmax
                        Ignition(i) = 1 / (IgnitionTime(i) ^ (1 / Nmax))
                    End If
                Next

                Call Ignition_Correlation2(radflux, Ignition, IgnitionTime, IgnPoints, emissivity, ThermalInertia, IgnitionTemp, QCritical)

                'determine effective heat of gasification using Quintiere's method
                If frmQuintiere.optUseOneCurve.Checked = True Then
                    Call EHoG_Correlation(radflux, Peak, IgnPoints, EffectiveHeatofCombustion, HeatofGasification)
                End If

                'Nmax = 1 / Nmax

            Else
                FTP = 0
            End If

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in Flame_Spread_Props")
        End Try
    End Sub

    Public Sub Resize_CVents()
        '*  =================================================
        '*      This procedure resizes the vent arrays to
        '*      optimise their size.
        '*
        '*      Global variables used:
        '*          NumberVents, SoffitHeight, VentHeight
        '*          VentWidth, VentSillHeight
        '*  =================================================

        Dim i, j As Short

        Try


            MaxNumberCVents = 0
            For i = 1 To NumberRooms + 1
                For j = 1 To NumberRooms + 1
                    If NumberCVents(i, j) > MaxNumberCVents Then
                        MaxNumberCVents = NumberCVents(i, j)
                    End If
                Next j
            Next i

            If MaxNumberCVents > 0 Then
                ReDim Preserve CVentOpenTime(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)
                ReDim Preserve CVentArea(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)
                ReDim Preserve CVentAuto(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)
                ReDim Preserve CVentCloseTime(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)
                ReDim Preserve trigger_device2(0 To 6, MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)
                ReDim Preserve FRintegrity2(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)
                ReDim Preserve FRMaxOpening2(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)
                ReDim Preserve FRMaxOpeningTime2(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)
                ReDim Preserve FRgastemp2(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)

                ReDim Preserve AutoOpenCVent(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)
                ReDim Preserve SDtriggerroom2(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)
                ReDim Preserve HRR_ventopenduration2(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)
                ReDim Preserve HRR_ventopendelay2(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)
                ReDim Preserve trigger_ventopenduration2(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)
                ReDim Preserve HRR_threshold2(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)
                ReDim Preserve trigger_ventopendelay2(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)
                ReDim Preserve HRR_threshold_time2(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)
                ReDim Preserve HRRFlag2(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)
                ReDim Preserve CVentBreakTime(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)
                ReDim Preserve CVentBreakTimeClose(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)

            End If

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Module1.vb Resize_CVents")
        End Try
    End Sub


    Function CeilingJet_MaxTemp(ByVal Q As Double) As Double
        '*  =============================================================
        '*      This function returns the value of the maximum temperature
        '*      in the ceiling jet at the location of the sprinkler or
        '*      detector using the Alpert quasi-steady correlations.
        '*
        '*      Arguments passed to the function:
        '*      Q = heat release rate for object 1
        '*  =============================================================

        Dim h As Single

        'ceiling height above base of fire
        h = RoomHeight(fireroom) - FireHeight(1)

        If FireLocation(1) = 0 Then
            'in the centre of the room
            CeilingJet_MaxTemp = 16.9 * Q ^ (2 / 3) * h ^ (-5 / 3) + InteriorTemp
        ElseIf FireLocation(1) = 1 Then
            'against wall
            CeilingJet_MaxTemp = 16.9 * (2 * Q) ^ (2 / 3) * h ^ (-5 / 3) + InteriorTemp
        Else
            'in corner
            CeilingJet_MaxTemp = 16.9 * (4 * Q) ^ (2 / 3) * h ^ (-5 / 3) + InteriorTemp
            'CeilingJet_MaxTemp = 16.9 * (4 * Q) ^ (2 / 3) * h ^ (-5 / 3) + uppertemp(fireroom, stepcount)
        End If

    End Function

    Public Sub Resize_Rooms()
        '*  =================================================
        '*      This procedure resizes the room arrays to
        '*      optimise their size.
        '*
        '   C Wade 8/8/98
        '*  =================================================

        If NumberRooms > 0 Then
            ReDim Preserve F(4, 4, NumberRooms)
            ReDim Preserve TransmissionFactor(4, 4, NumberRooms)
            ReDim Preserve RoomHeight(NumberRooms)
            ReDim Preserve RoomWidth(NumberRooms)
            ReDim Preserve FloorElevation(NumberRooms)
            ReDim Preserve RoomAbsX(NumberRooms)
            ReDim Preserve RoomAbsY(NumberRooms)
            ReDim Preserve RoomLength(NumberRooms)
            ReDim Preserve RoomDescription(NumberRooms)
            ReDim Preserve MinStudHeight(NumberRooms)
            ReDim Preserve RoomVolume(NumberRooms)
            ReDim Preserve RoomFloorArea(NumberRooms)
            ReDim Preserve CeilingSlope(NumberRooms)
            ReDim Preserve FanAutoStart(NumberRooms)
            ReDim Preserve TwoZones(0 To NumberRooms)
            ReDim Preserve WallSubThickness(NumberRooms)
            ReDim Preserve CeilingSubThickness(NumberRooms)
            ReDim Preserve FloorSubThickness(NumberRooms)
            ReDim Preserve WallThickness(NumberRooms)
            ReDim Preserve CeilingThickness(NumberRooms)
            ReDim Preserve FloorThickness(NumberRooms)
            ReDim Preserve WallDensity(NumberRooms)
            ReDim Preserve CeilingDensity(NumberRooms)
            ReDim Preserve WallSubDensity(NumberRooms)
            ReDim Preserve FloorSubDensity(NumberRooms)
            ReDim Preserve CeilingSubDensity(NumberRooms)
            ReDim Preserve FloorDensity(NumberRooms)
            ReDim Preserve WallSubSpecificHeat(NumberRooms)
            ReDim Preserve FloorSubSpecificHeat(NumberRooms)
            ReDim Preserve CeilingSubSpecificHeat(NumberRooms)
            ReDim Preserve WallSpecificHeat(NumberRooms)
            ReDim Preserve CeilingSpecificHeat(NumberRooms)
            ReDim Preserve FloorSpecificHeat(NumberRooms)
            ReDim Preserve HaveWallSubstrate(NumberRooms)
            ReDim Preserve HaveFloorSubstrate(NumberRooms)
            ReDim Preserve SDFlag(NumberRooms + 1)
            ReDim Preserve SDTime(NumberRooms)
            ReDim Preserve HaveSD(NumberRooms)
            ReDim Preserve SDinside(NumberRooms)
            ReDim Preserve SpecifyOD(NumberRooms)
            ReDim Preserve HaveCeilingSubstrate(NumberRooms)
            ReDim Preserve WallConductivity(NumberRooms)
            ReDim Preserve CeilingConductivity(NumberRooms)
            ReDim Preserve WallSubConductivity(NumberRooms)
            ReDim Preserve FloorSubConductivity(NumberRooms)
            ReDim Preserve CeilingSubConductivity(NumberRooms)
            ReDim Preserve FloorConductivity(NumberRooms)
            ReDim Preserve WallSubstrate(NumberRooms)
            ReDim Preserve CeilingSubstrate(NumberRooms)
            ReDim Preserve FloorSubstrate(NumberRooms)
            ReDim Preserve WallSurface(NumberRooms)
            ReDim Preserve CeilingSurface(NumberRooms)
            ReDim Preserve FloorSurface(NumberRooms)
            ReDim Preserve WallConeDataFile(NumberRooms)
            ReDim Preserve FloorConeDataFile(NumberRooms)
            ReDim Preserve CeilingConeDataFile(NumberRooms)
            ReDim Preserve WallSootYield(NumberRooms)
            ReDim Preserve CeilingSootYield(NumberRooms)
            ReDim Preserve FloorSootYield(NumberRooms)
            ReDim Preserve WallH2OYield(NumberRooms)
            ReDim Preserve CeilingH2OYield(NumberRooms)
            ReDim Preserve FloorH2OYield(NumberRooms)
            ReDim Preserve WallHCNYield(NumberRooms)
            ReDim Preserve CeilingHCNYield(NumberRooms)
            ReDim Preserve FloorHCNYield(NumberRooms)
            ReDim Preserve WallCO2Yield(NumberRooms)
            ReDim Preserve CeilingCO2Yield(NumberRooms)
            ReDim Preserve FloorCO2Yield(NumberRooms)
            ReDim Preserve AreaUnderCeilingCurve(NumberRooms)
            ReDim Preserve AreaUnderFloorCurve(NumberRooms)
            ReDim Preserve AreaUnderWallCurve(NumberRooms)
            ReDim Preserve WallEffectiveHeatofCombustion(NumberRooms)
            ReDim Preserve CeilingEffectiveHeatofCombustion(NumberRooms)
            ReDim Preserve FloorEffectiveHeatofCombustion(NumberRooms)
            ReDim Preserve WallHeatofGasification(NumberRooms)
            ReDim Preserve WallFTP(NumberRooms)
            ReDim Preserve Walln(NumberRooms)
            ReDim Preserve WallQCrit(NumberRooms)
            ReDim Preserve FloorQCrit(NumberRooms)
            ReDim Preserve CeilingQCrit(NumberRooms)
            ReDim Preserve CeilingFTP(NumberRooms)
            ReDim Preserve FloorFTP(NumberRooms)
            ReDim Preserve Ceilingn(NumberRooms)
            ReDim Preserve Floorn(NumberRooms)
            ReDim Preserve CeilingHeatofGasification(NumberRooms)
            ReDim Preserve FloorHeatofGasification(NumberRooms)
            ReDim Preserve WallTSMin(NumberRooms)
            ReDim Preserve FloorTSMin(NumberRooms)
            ReDim Preserve CurveNumber_F(NumberRooms)
            ReDim Preserve CurveNumber_W(NumberRooms)
            ReDim Preserve CurveNumber_C(NumberRooms)
            ReDim Preserve PeakFloorHRR(MaxConeCurves, NumberRooms)
            ReDim Preserve PeakCeilingHRR(MaxConeCurves, NumberRooms)
            ReDim Preserve PeakWallHRR(MaxConeCurves, NumberRooms)
            ReDim Preserve WallFlameSpreadParameter(NumberRooms)
            ReDim Preserve FloorFlameSpreadParameter(NumberRooms)
            ReDim Preserve ThermalInertiaCeiling(NumberRooms)
            ReDim Preserve ThermalInertiaFloor(NumberRooms)
            ReDim Preserve IgTempF(NumberRooms)
            ReDim Preserve IgTempC(NumberRooms)
            ReDim Preserve IgTempW(NumberRooms)
            ReDim Preserve ThermalInertiaWall(NumberRooms)
            ReDim Preserve AlphaFloor(NumberRooms)
            ReDim Preserve AlphaWall(NumberRooms)
            ReDim Preserve AlphaCeiling(NumberRooms)
            ReDim Preserve WallDeltaX(NumberRooms)
            ReDim Preserve CeilingDeltaX(NumberRooms)
            ReDim Preserve FloorDeltaX(NumberRooms)
            ReDim Preserve WallOutsideBiot(NumberRooms)
            ReDim Preserve CeilingOutsideBiot(NumberRooms)
            ReDim Preserve FloorOutsideBiot(NumberRooms)
            ReDim Preserve FloorFourier(NumberRooms)
            ReDim Preserve CeilingFourier(NumberRooms)
            ReDim Preserve WallFourier(NumberRooms)
            ReDim Preserve Surface_Emissivity(4, NumberRooms)
            ReDim Preserve ExtractRate(NumberRooms)
            ReDim Preserve NumberFans(NumberRooms)
            ReDim Preserve ExtractStartTime(NumberRooms)
            ReDim Preserve FanElevation(NumberRooms)
            ReDim Preserve MaxPressure(NumberRooms)
            ReDim Preserve fanon(NumberRooms)
            ReDim Preserve Extract(NumberRooms)
            ReDim Preserve UseFanCurve(NumberRooms)
            ReDim Preserve RTI(NumberRooms)
            ReDim Preserve RadialDistance(NumberRooms)
            ReDim Preserve SmokeOD(NumberRooms)
            ReDim Preserve SDdelay(NumberRooms)
            ReDim Preserve DetSensitivity(NumberRooms)
            ReDim Preserve SDRadialDist(NumberRooms)
            ReDim Preserve SDdepth(NumberRooms)
            ReDim Preserve WaterSprayDensity(NumberRooms)
            ReDim Preserve ActuationTemp(NumberRooms)
            ReDim Preserve cfactor(NumberRooms)
            ReDim Preserve SprinkDistance(NumberRooms)
            ReDim Preserve DetectorType(NumberRooms)
            ReDim Preserve Sprinkler(3, NumberRooms)
        End If
    End Sub
    Public Sub Graph_Data_cdf(ByVal Title As String, ByRef datatobeplotted(,) As Double, ByVal datashift As Double, ByVal datamultiplier As Double)
        '*  ====================================================================
        '*  This function plots a CDF graph
        '*  ====================================================================

        Dim j As Integer
        Dim room As Integer
        Dim ydata(0 To 100) As Double

        'if no data exists
        If NumberTimeSteps < 1 Then
            MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Sub
        End If

        Try

            frmPlot.Chart1.Series.Clear()
            room = 1
            'For room = 1 To NumberRooms

            frmPlot.Chart1.Series.Add("Room " & room)

            frmPlot.Chart1.Series("Room " & room).ChartType = SeriesChartType.FastLine

            For j = 1 To 100
                ydata(j) = datatobeplotted(j, room) * datamultiplier + datashift 'data to be plotted
                frmPlot.Chart1.Series("Room " & room).Points.AddXY(ydata(j), j / 100)
            Next

            'Next room
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Maximum = 1
            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = "Cumulative Density Function"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.LabelStyle.Format = "0.0"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = Title
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            'frmPlot.Chart1.Titles("Title1").Text = Title
            ' Disable X axis margin

            frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Graph_Data_cdf")
        End Try
    End Sub
    Public Sub Graph_Data_2D(ByVal Title As String, ByRef datatobeplotted(,) As Double, ByVal datashift As Double, ByVal datamultiplier As Double, ByVal timeunit As Integer)
        '*  ====================================================================
        '*  This function takes data for a variable from a two-dimensional array
        '*  and displays it in a graph
        '*  ====================================================================

        Dim j As Integer
        Dim room As Integer
        Dim ydata(0 To NumberTimeSteps) As Double

        'if no data exists
        If (NumberTimeSteps < 1) Then
            MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Sub
        End If
        If datatobeplotted Is Nothing Then
            MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Sub
        End If

        Try

            frmPlot.Chart1.Series.Clear()

            For room = 1 To NumberRooms

                frmPlot.Chart1.Series.Add("Room " & room)

                frmPlot.Chart1.Series("Room " & room).ChartType = SeriesChartType.FastLine

                If Not roomcolor(room - 1).IsEmpty Then
                    frmPlot.Chart1.Series("Room " & room).Color = roomcolor(room - 1) '=line color
                End If

                For j = 1 To NumberTimeSteps

                    ydata(j) = datatobeplotted(room, j) * datamultiplier + datashift 'data to be plotted
                    frmPlot.Chart1.Series("Room " & room).Points.AddXY(tim(j, 1) / timeunit, ydata(j))
                Next

            Next room

            If Title = "Vent Fires (kW)" Then
                frmPlot.Chart1.Series.Add("Outside")

                frmPlot.Chart1.Series("Outside").ChartType = SeriesChartType.FastLine

                For j = 1 To NumberTimeSteps
                    ydata(j) = datatobeplotted(NumberRooms + 1, j) * datamultiplier + datashift 'data to be plotted
                    frmPlot.Chart1.Series("Outside").Points.AddXY(tim(j, 1) / timeunit, ydata(j))

                Next
            End If


            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            'frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            If timeunit = 60 Then
                frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (min)"
            Else
                frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            End If
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            'frmPlot.Chart1.Titles("Title1").Text = Title
            ' Disable X axis margin

            frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()


            If frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title.Contains("char depth (mm)") = True Then
                Dim getname = ""
                Dim sname = ""
                Dim oExcel As Object
                Dim oBook As Object
                Dim oSheet As Object

                getname = RiskDataDirectory & "excel_chardepth_" & Convert.ToString(frmInputs.txtBaseName.Text)

                If frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title.Contains("Ceiling") = True Then
                    sname = "Ceiling Char Depth"
                ElseIf frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title.Contains("Upper wall") = True Then
                    sname = "Upper Wall Char Depth"
                End If

                oExcel = CreateObject("Excel.Application")
                oBook = oExcel.Workbooks.Add
                oSheet = oBook.Worksheets(1)
                oSheet.name = sname
                oSheet.Range("A1").Value = "time (sec)"
                oSheet.Range("B1").Value = frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title
                Dim DataArray(0 To NumberTimeSteps, 0 To 1) As Double
                For j = 1 To NumberTimeSteps
                    DataArray(j, 0) = j * Timestep 's
                    DataArray(j, 1) = datatobeplotted(fireroom, j) 'depth
                Next
                oSheet.Range("A2").Resize(NumberTimeSteps + 1, 2).Value = DataArray
                oBook.worksheets.add()

                'Save the Workbook and Quit Excel
                oBook.SaveAs(getname)
                oExcel.Quit()
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Graph_Data_2D")
        End Try

    End Sub
    Public Sub Graph_Data_CharDepth(ByVal temp As String, ByVal Title As String, ByRef datatobeplotted(,) As Double, ByVal datashift As Double, ByVal datamultiplier As Double, ByVal timeunit As Integer)
        '*  ====================================================================
        '*  This function takes data for a variable from a two-dimensional array
        '*  and displays it in a graph
        '*  ====================================================================

        Dim j As Integer
        Dim room As Integer
        Dim ydata(0 To NumberTimeSteps) As Double

        'if no data exists
        If (NumberTimeSteps < 1) Then
            MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Sub
        End If
        If datatobeplotted Is Nothing Then
            MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Sub
        End If

        Try

            frmPlot.Chart1.Series.Clear()

            For room = 1 To NumberRooms

                frmPlot.Chart1.Series.Add("Room " & room)

                frmPlot.Chart1.Series("Room " & room).ChartType = SeriesChartType.FastLine

                If Not roomcolor(room - 1).IsEmpty Then
                    frmPlot.Chart1.Series("Room " & room).Color = roomcolor(room - 1) '=line color
                End If

                For j = 1 To NumberTimeSteps

                    ydata(j) = datatobeplotted(room, j) * datamultiplier + datashift 'data to be plotted
                    frmPlot.Chart1.Series("Room " & room).Points.AddXY(tim(j, 1) / timeunit, ydata(j))
                Next

            Next room


            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title

            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            If timeunit = 60 Then
                frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (min)"
            Else
                frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            End If
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right


            frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()


            If frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title.Contains("char depth (mm)") = True Then
                Dim getname = ""
                Dim sname = ""
                Dim oExcel As Object
                Dim oBook As Object
                Dim oSheet As Object

                If frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title.Contains("Ceiling") = True Then
                    sname = "Ceiling Char Depth"
                ElseIf frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title.Contains("Upper wall") = True Then
                    sname = "Upper Wall Char Depth"
                End If

                getname = RiskDataDirectory & sname.ToString & "_" & temp.ToString & "C"
                oExcel = CreateObject("Excel.Application")
                oBook = oExcel.Workbooks.Add
                oSheet = oBook.Worksheets(1)
                oSheet.name = sname
                oSheet.Range("A1").Value = "time (sec)"
                oSheet.Range("B1").Value = frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title
                oSheet.Range("C1").Value = "time (min)"
                Dim DataArray(0 To NumberTimeSteps, 0 To 2) As Double
                For j = 1 To NumberTimeSteps
                    DataArray(j, 0) = j * Timestep 's
                    DataArray(j, 1) = datatobeplotted(fireroom, j) 'depth
                    DataArray(j, 2) = j * Timestep / 60 'min
                Next
                oSheet.Range("A2").Resize(NumberTimeSteps + 1, 3).Value = DataArray
                oBook.worksheets.add()

                'Save the Workbook and Quit Excel
                oBook.SaveAs(getname)
                oExcel.Quit()
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Graph_Data_CharDepth")
        End Try

    End Sub
    Public Sub Graph_Data_FED(ByVal Title As String, ByRef datatobeplotted(,) As Double, ByVal datashift As Double, ByVal datamultiplier As Double, ByVal timeunit As Integer)
        '*  ====================================================================
        '*  This function takes data for a variable from a two-dimensional array
        '*  and displays it in a graph
        '*  ====================================================================

        Dim j As Integer
        Dim room As Integer
        Dim ydata(0 To NumberTimeSteps) As Double

        'if no data exists
        If (NumberTimeSteps < 1) Then
            MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Sub
        End If
        If datatobeplotted Is Nothing Then
            MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Sub
        End If

        Try

            frmPlot.Chart1.Series.Clear()

            'For room = 1 To NumberRooms

            'frmPlot.Chart1.Series.Add("Room " & fireroom)
            frmPlot.Chart1.Series.Add("Egress Path")
            frmPlot.Chart1.Series("Egress Path").ChartType = SeriesChartType.FastLine

            room = fireroom

            For j = 1 To NumberTimeSteps

                ydata(j) = datatobeplotted(room, j) * datamultiplier + datashift 'data to be plotted
                frmPlot.Chart1.Series("Egress Path").Points.AddXY(tim(j, 1) / timeunit, ydata(j))
            Next

            ' Next room

            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            'frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            If timeunit = 60 Then
                frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (min)"
            Else
                frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            End If
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right

            'frmPlot.Chart1.Annotations.Clear()
            Dim mytext As TextAnnotation = New TextAnnotation With {
                .Text = "Room " & FEDPath(2, 0) & " from " & FEDPath(0, 0) & " to " & FEDPath(1, 0) & " sec"
            }
            If FEDPath(1, 1) > FEDPath(1, 0) Then
                mytext.Text = mytext.Text & "; Room " & FEDPath(2, 1) & " from " & FEDPath(0, 1) & " to " & FEDPath(1, 1) & " sec"
            End If
            If FEDPath(1, 2) > FEDPath(1, 1) Then
                mytext.Text = mytext.Text & "; Room " & FEDPath(2, 2) & " from " & FEDPath(0, 2) & " to " & FEDPath(1, 2) & " sec;"
            End If
            'mytext.X = 0
            'mytext.Y = 0
            'mytext.AllowMoving = True
            'mytext.Alignment = ContentAlignment.TopCenter
            'frmPlot.Chart1.Annotations.Add(mytext)

            frmPlot.Chart1.Titles(0).Text = mytext.Text
            frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Graph_Data_FED")
        End Try

    End Sub

    Public Sub Graph_Data_2Dsprink(ByVal Title As String, ByRef datatobeplotted(,) As Double, ByVal datashift As Double, ByVal datamultiplier As Double)
        '*  ====================================================================
        '*  This function takes data for a variable from a two-dimensional array
        '*  and displays it in a graph
        '*  ====================================================================

        Dim j As Integer
        Dim id As Integer
        Dim ydata(0 To NumberTimeSteps) As Double

        'if no data exists
        If NumberTimeSteps < 1 Then
            MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Sub
        End If

        Try

            frmPlot.Chart1.Series.Clear()
            id = 1
            For id = 1 To NumSprinklers

                frmPlot.Chart1.Series.Add("Sprink ID " & id)

                frmPlot.Chart1.Series("Sprink ID " & id).ChartType = SeriesChartType.FastLine

                For j = 1 To NumberTimeSteps
                    ydata(j) = datatobeplotted(id, j) * datamultiplier + datashift 'data to be plotted
                    frmPlot.Chart1.Series("Sprink ID " & id).Points.AddXY(tim(j, 1), ydata(j))
                Next

            Next id
            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            'frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            'frmPlot.Chart1.Titles("Title1").Text = Title
            ' Disable X axis margin

            frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Graph_Data_2Dsprink")
        End Try

    End Sub

    Public Function O2_limit_cfast(ByVal room As Integer, ByVal Mass_Upper As Double, ByVal Q As Double, ByVal mplume As Double, ByVal UT As Double, ByVal YO2 As Double, ByVal YO2Lower As Double, ByVal mw_upper As Double, ByVal mw_lower As Double, ByRef Qplume As Double) As Double
        '*  =================================================
        '*  This function determines if the burning is limited
        '*  by the available air supply and if so returns a reduced
        '*  value for the heat release rate.
        '*
        '*  Arguments passes to the function:
        '*
        '*      Q   =   theoretical heat release of fire (kW)
        '*      MPlume = mass flow of air in plume
        '*      UT = upper layer temperature
        '*      YO2 = mass fraction of oxygen in upper layer
        '*
        '*  Revised: 13 April 1996  Colleen Wade
        '*  =================================================

        Dim O2needed, O2critical_U, oxygen, O2actual, CLOL, LOL As Double
        Dim Y As Double = 0
        'how much O2 in plume flow needed for complete combustion?
        O2needed = Q / 13100 'kg O2 /sec

        'oxygen in the upper layer needed to support combustion
        'O2critical_U = DMax1(2, (873 - UT) / (873 - InteriorTemp) * (10 - 2) + 2) '%
        O2critical_U = (873 - UT) / (873 - InteriorTemp) * (10 - 2) + 2 '%
        If O2critical_U < 2 Then O2critical_U = 2
        'O2critical_L = DMax1(2, (873 - UT) / (873 - InteriorTemp) * (10 - 2) + 2) '%
        'O2critical = 0
        'LOL = MolecularWeightO2 / MolecularWeightAir / 100 * O2critical
        'LOL = MolecularWeightO2 / mw_upper / 100 * O2critical
        LOL = MolecularWeightO2 / mw_lower / 100 * 10 'critical mass fraction in lower layer

        CLOL = (TANH(800 * (YO2Lower - LOL) - 4) + 1) / 2
        'CLOL = (TANH(800 * (YO2Lower - 0.1) - 4) + 1) / 2

        'how much do we have?
        'mass fraction of O2 in lower layer (ambient)
        'O2actual = MPlume * O2Ambient  'kg O2 /sec
        O2actual = mplume * YO2Lower * CLOL 'kg O2 /sec

        If O2needed <= O2actual Or Q <= 0.000001 Then
            'enough air in the plume flow for complete combustion
            If g_post = True And Flashover = True And burnmode = True Then
                'if using wood crib model under ventilation control, then we want to burn as much as the oxygen can support
                '???? don't understand this, changed 27/05/16
                O2_limit_cfast = (O2actual) * 13100
                Qplume = O2actual * 13100

                'these had been changed, i've changed them back 18/6/16 for the reason above
                'O2_limit_cfast = Q
                'Qplume = Q
            Else
                'If g_post = True And Flashover = True Then
                '    If Q > O2actual * 13100 Then Stop
                'End If


                O2_limit_cfast = Q
                Qplume = Q
            End If
        Else
            'not enough oxygen in plume for complete combustion
            'volume % of oxygen in the upper layer
            oxygen = YO2 * mw_upper / MolecularWeightO2 * 100

            'combustion can occur in upper layer if o2 concentration
            'is above a critical value
            If oxygen > O2critical_U Then
                'enough o2 in the layer for complete combustion
                'instead of burning everything, burn only that using excess oxygen (>2%) in the upper layer
                'mass fraction of o2 corresponding to 2% by vol

                If Mass_Upper > 0 Then
                    'Y = O2critical_U / 100 * MolecularWeightO2 / mw_upper
                    'CLOL = (TANH(800 * (YO2 - Y) - 4) + 1) / 2

                    ''use all oxygen in upper layer in one second
                    'O2_limit_cfast = (CLOL * (YO2 - Y) * Mass_Upper + O2actual) * 13100

                    'the above was removed 
                    O2_limit_cfast = (O2actual) * 13100

                    'use all oxygen in upper layer in one timestep
                    'O2_limit_cfast = CLOL * (YO2 - Y) * Mass_Upper / Timestep * 13100 + O2actual * 13100
                    'O2_limit_cfast = O2actual * 13100
                Else
                    O2_limit_cfast = O2actual * 13100
                    'O2_limit_cfast = Q '24/2/2005 if the upper layer hardly exists, just assume that unrestricted burning can occur
                End If

                If Q < O2_limit_cfast Then
                    O2_limit_cfast = Q
                End If

            Else
                'not enough air for complete combustion, so restrict to what the plume
                'flow can support
                O2_limit_cfast = O2actual * 13100 'kW
                'O2_limit_cfast = Q
            End If

            Qplume = O2actual * 13100
        End If
    End Function
    Public Function O2_limit_cfast2018(ByVal room As Integer, ByVal Mass_Upper As Double, ByVal Q As Double, ByVal mplume As Double, ByVal UT As Double, ByVal YO2 As Double, ByVal YO2Lower As Double, ByVal layerz As Double, ByRef Qplume As Double) As Double
        '*  =================================================
        '*  This function determines if the burning is limited
        '*  by the available air supply and if so returns a reduced
        '*  value for the heat release rate.
        '*
        '*  Arguments passes to the function:
        '*
        '*      Q   =   theoretical heat release of fire (kW)
        '*      MPlume = mass flow of air in plume
        '*      UT = upper layer temperature
        '*      YO2 = mass fraction of oxygen in upper layer
        '*
        '*  =================================================

        Dim O2needed, O2actual, CLOLL, CLOLU, LOL As Double
        Dim Y As Double = 0
        LOL = 0.15

        'how much O2 in plume flow needed for complete combustion?
        O2needed = Q / 13100 'kg O2 /sec

        'If fire in lower layer
        CLOLL = (TANH(800 * (YO2Lower - LOL) - 4) + 1) / 2

        'If fire in upper layer
        CLOLU = (TANH(800 * (YO2 - LOL) - 4) + 1) / 2

        'how much do we have in plume?
        O2actual = CLOLL * mplume * YO2Lower 'kg O2 /sec

        If O2needed <= O2actual Or Q <= 0.000001 Then
            'enough air in the plume flow for complete combustion
            If g_post = True And Flashover = True And burnmode = True Then
                'if using wood crib model under ventilation control, then we want to burn as much as the oxygen can support
                '???? don't understand this, changed 27/05/16
                O2_limit_cfast2018 = (O2actual) * 13100
                Qplume = O2actual * 13100

            Else

                O2_limit_cfast2018 = Q
                Qplume = Q
            End If
        Else
            'not enough oxygen in plume for complete combustion

            'combustion can occur in upper layer if o2 concentration
            'is above a critical value
            If YO2 > LOL Then
                'keep burning while there is enough o2 in upper layer
                O2_limit_cfast2018 = Q

            Else
                'not enough air for complete combustion, so restrict to what the plume
                'flow can support
                O2_limit_cfast2018 = O2actual * 13100 'kW

            End If

            Qplume = O2actual * 13100

        End If
    End Function
    Public Sub View_output(ByRef filename As String)
        '*  ====================================================================
        '*  View the results 
        '*  ====================================================================

        Dim room, k, i, howmany, j, X, Y, m, vent As Integer
        Dim energy, frr As Double
        Dim tempvar As Single
        On Error GoTo errorhandler
        ReDim Preserve CeilingIgniteFlag(NumberRooms)

        Call view_input(filename)
        If IsNothing(tim) Then Exit Sub

        FileOpen(1, (filename), OpenMode.Append)

        Dim dummy As String
        If NumberTimeSteps <> 0 Then

            PrintLine(1, "====================================================================")
            PrintLine(1, "Results from Fire Simulation")
            PrintLine(1, "====================================================================")
            PrintLine(1)

            If DisplayInterval < Timestep Then
                MsgBox("The specified display time interval must be equal or greater than " & CStr(Timestep) & " seconds.")
                FileClose(1)
                Exit Sub
            End If
            howmany = Int((NumberTimeSteps + 1) * Timestep / DisplayInterval)
            j = 1
            For k = 1 To howmany + 1
                If j > NumberTimeSteps + 1 Then Exit For

                Print(1, VB6.Format(Int(tim(j, 1) / 60), "0") & " min", TAB(10), VB6.Format(tim(j, 1) Mod 60, "00") & " sec", TAB(10), "(" & VB6.Format(Int(tim(j, 1)), "0") & " sec)")
                X = 40
                For room = 1 To NumberRooms
                    Print(1, TAB(X), "Room " & CStr(room))
                    X = X + 10
                Next room
                PrintLine(1, TAB(X), "Outside")
                PrintLine(1)

                X = 40
                Y = 40
                room = 1
                If frmprintvar.chkLH.CheckState = System.Windows.Forms.CheckState.Checked Then
                    Print(1, TAB(10), "Layer (m)", TAB(X), VB6.Format(layerheight(room, j), "0.000"))
                End If
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkLH.CheckState = System.Windows.Forms.CheckState.Checked Then
                        Print(1, TAB(X), VB6.Format(layerheight(room, j), "0.000"))
                    End If
                Next room
                If frmprintvar.chkLH.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkUT.CheckState = System.Windows.Forms.CheckState.Checked Then
                    Print(1, TAB(10), "Upper Temp (C)", TAB(X), VB6.Format(uppertemp(room, j) - 273, "0.0"))
                End If
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkUT.CheckState = System.Windows.Forms.CheckState.Checked Then
                        Print(1, TAB(X), VB6.Format(uppertemp(room, j) - 273, "0.0"))
                    End If
                Next room
                If frmprintvar.chkUT.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkLT.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Lower Temp (C)", TAB(X), VB6.Format(lowertemp(room, j) - 273, "0.0"))
                For room = 2 To NumberRooms
                    If TwoZones(room) = True Then
                        X = X + 10
                        If frmprintvar.chkLT.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(lowertemp(room, j) - 273, "0.0"))
                    End If
                Next room
                If frmprintvar.chkLT.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkHRR_input.CheckState = System.Windows.Forms.CheckState.Checked Then
                    If room = fireroom Then
                        Print(1, TAB(10), "Unconstrained HRR (kW)", TAB(X), VB6.Format(HeatRelease(fireroom, j, 1), "0.0"))
                    Else
                        Print(1, TAB(10), "Unconstrained HRR (kW)", TAB(X), VB6.Format(0, "0.0"))
                    End If
                End If
                For room = 2 To NumberRooms
                    X = X + 10
                    If room = fireroom Then
                        If frmprintvar.chkHRR_input.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(HeatRelease(fireroom, j, 1), "0.0"))
                    Else
                        If frmprintvar.chkHRR_input.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(0, "0.0"))
                    End If
                Next room
                If frmprintvar.chkHRR_input.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)


                X = Y
                room = 1
                If frmprintvar.chkHRR.CheckState = System.Windows.Forms.CheckState.Checked Then
                    If room = fireroom Then
                        Print(1, TAB(10), "HRR (kW)", TAB(X), VB6.Format(HeatRelease(fireroom, j, 2), "0.0"))
                    Else
                        Print(1, TAB(10), "HRR (kW)", TAB(X), VB6.Format(0, "0.0"))
                    End If
                End If
                For room = 2 To NumberRooms
                    X = X + 10
                    If room = fireroom Then
                        If frmprintvar.chkHRR.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(HeatRelease(fireroom, j, 2), "0.0"))
                    Else
                        If frmprintvar.chkHRR.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(0, "0.0"))
                    End If
                Next room
                If frmprintvar.chkHRR.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkQstar.Checked = True Then
                    If room = fireroom Then
                        'Print(1, TAB(10), "Q* =", TAB(X), VB6.Format(HeatRelease(fireroom, j, 2) / (1110 * RoomHeight(room) ^ (5 / 2)), "0.0000"))
                        Print(1, TAB(10), "Q* =", TAB(X), Format(HeatRelease(fireroom, j, 2) / (1110 * (MinStudHeight(room) / 2 + RoomHeight(room) / 2) ^ (5 / 2)), "0.0000"))
                    Else
                        Print(1, TAB(10), "Q* =", TAB(X), Format(0, "-"))
                    End If
                    For room = 2 To NumberRooms
                        X = X + 10
                        If room = fireroom Then
                            'Print(1, TAB(X), VB6.Format(HeatRelease(fireroom, j, 2) / (1110 * RoomHeight(room) ^ (5 / 2)), "0.0000"))
                            Print(1, TAB(10), "Q* =", TAB(X), Format(HeatRelease(fireroom, j, 2) / (1110 * (MinStudHeight(room) / 2 + RoomHeight(room) / 2) ^ (5 / 2)), "0.0000"))
                        Else
                            Print(1, TAB(X), VB6.Format(0, "-"))
                        End If
                    Next room
                End If
                If frmprintvar.chkQstar.Checked = True Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkMassLoss.CheckState = System.Windows.Forms.CheckState.Checked Then
                    If room = fireroom Then
                        Print(1, TAB(10), "Mass Loss (kg/s)", TAB(X), VB6.Format(FuelMassLossRate(j, room), "0.000"))
                    Else
                        Print(1, TAB(10), "Mass Loss (kg/s)", TAB(X), VB6.Format(FuelMassLossRate(j, room), "0.000"))
                    End If
                End If
                For room = 2 To NumberRooms
                    X = X + 10
                    If room = fireroom Then
                        If frmprintvar.chkMassLoss.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(FuelMassLossRate(j, room), "0.000"))
                    Else
                        If frmprintvar.chkMassLoss.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(0, "0.000"))
                    End If
                Next room
                If frmprintvar.chkMassLoss.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                If frmOptions1.optPostFlashover.Checked = True Then

                    X = Y
                    room = 1
                    If frmprintvar.chkFuelMassRemaining.CheckState = System.Windows.Forms.CheckState.Checked Then
                        If room = fireroom Then
                            Print(1, TAB(10), "Fuel Mass Remaining (kg)", TAB(X), VB6.Format(InitialFuelMass - TotalFuel(j), "0.0"))
                        Else
                            Print(1, TAB(10), "Fuel Mass Remaining (kg)", TAB(X), VB6.Format(0, "0.0"))
                        End If
                    End If
                    For room = 2 To NumberRooms
                        X = X + 10
                        If room = fireroom Then
                            If frmprintvar.chkFuelMassRemaining.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(InitialFuelMass - TotalFuel(j), "0.0"))
                        Else
                            If frmprintvar.chkFuelMassRemaining.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(0, "0.0"))
                        End If
                    Next room
                    If frmprintvar.chkFuelMassRemaining.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)
                End If

                X = Y
                room = 1
                If frmprintvar.chkPlume.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Plume Flow (kg/s)", TAB(X), VB6.Format(massplumeflow(j, room), "0.000"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkPlume.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(massplumeflow(j, room), "0.000"))
                Next room
                If frmprintvar.chkPlume.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkPressure.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Pressure (Pa)", TAB(X), VB6.Format(RoomPressure(room, j), "0.00"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkPressure.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(RoomPressure(room, j), "0.00"))
                Next room
                If frmprintvar.chkPressure.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkLCO.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "CO Lower (ppm)", TAB(X), VB6.Format(COVolumeFraction(room, j, 2) * 1000000, "0"))
                For room = 2 To NumberRooms
                    If room = fireroom Then
                        X = X + 10
                        If frmprintvar.chkLCO.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(COVolumeFraction(room, j, 2) * 1000000, "0"))
                    End If
                Next room
                If frmprintvar.chkLCO.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkUCO.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "CO Upper (ppm)", TAB(X), VB6.Format(COVolumeFraction(room, j, 1) * 1000000, "0"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkUCO.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(COVolumeFraction(room, j, 1) * 1000000, "0"))
                Next room
                If frmprintvar.chkUCO.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkLhcn.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "HCN Lower (ppm)", TAB(X), VB6.Format(HCNVolumeFraction(room, j, 2) * 1000000, "0.0"))
                For room = 2 To NumberRooms
                    If TwoZones(room) = True Then
                        X = X + 10
                        If frmprintvar.chkLhcn.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(HCNVolumeFraction(room, j, 2) * 1000000, "0.0"))
                    End If
                Next room
                If frmprintvar.chkLhcn.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkUhcn.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "HCN Upper (ppm)", TAB(X), VB6.Format(HCNVolumeFraction(room, j, 1) * 1000000, "0.0"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkUhcn.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(HCNVolumeFraction(room, j, 1) * 1000000, "0.0"))
                Next room
                If frmprintvar.chkUhcn.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkLCO2.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "CO2 Lower (%)", TAB(X), VB6.Format(CO2VolumeFraction(room, j, 2) * 100, "0.0"))
                For room = 2 To NumberRooms
                    If TwoZones(room) = True Then
                        X = X + 10
                        If frmprintvar.chkLCO2.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(CO2VolumeFraction(room, j, 2) * 100, "0.0"))
                    End If
                Next room
                If frmprintvar.chkLCO2.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkUCO2.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "CO2 Upper (%)", TAB(X), VB6.Format(CO2VolumeFraction(room, j, 1) * 100, "0.0"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkUCO2.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(CO2VolumeFraction(room, j, 1) * 100, "0.0"))
                Next room
                If frmprintvar.chkUCO2.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkLO.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "O2 Lower (%)", TAB(X), VB6.Format(O2VolumeFraction(room, j, 2) * 100, "0.0"))
                For room = 2 To NumberRooms
                    If TwoZones(room) = True Then
                        X = X + 10
                        If frmprintvar.chkLO.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(O2VolumeFraction(room, j, 2) * 100, "0.0"))
                    End If
                Next room
                If frmprintvar.chkLO.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkUO.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "O2 Upper (%)", TAB(X), VB6.Format(O2VolumeFraction(room, j, 1) * 100, "0.0"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkUO.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(O2VolumeFraction(room, j, 1) * 100, "0.0"))
                Next room
                If frmprintvar.chkUO.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkLH2O.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "H2O Lower (%)", TAB(X), VB6.Format(H2OVolumeFraction(room, j, 2) * 100, "0.0"))
                For room = 2 To NumberRooms
                    If TwoZones(room) = True Then
                        X = X + 10
                        If frmprintvar.chkLH2O.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(H2OVolumeFraction(room, j, 2) * 100, "0.0"))
                    End If
                Next room
                If frmprintvar.chkLH2O.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkUH2O.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "H2O Upper (%)", TAB(X), VB6.Format(H2OVolumeFraction(room, j, 1) * 100, "0.0"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkUH2O.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(H2OVolumeFraction(room, j, 1) * 100, "0.0"))
                Next room
                If frmprintvar.chkUH2O.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkTar.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Radiation on Floor (kW/m2)", TAB(X), VB6.Format(Target(room, j), "0.00"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkTar.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(Target(room, j), "0.00"))
                Next room
                If frmprintvar.chkTar.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkTargetRad.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Radiation on Target (kW/m2)", TAB(X), VB6.Format(SurfaceRad(room, j), "0.00"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkTargetRad.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(SurfaceRad(room, j), "0.00"))
                Next room
                If frmprintvar.chkTargetRad.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkVentFire.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Vent Fire (kW)", TAB(X), VB6.Format(ventfire(room, j), "0.0"))
                For room = 2 To NumberRooms + 1
                    X = X + 10
                    If frmprintvar.chkVentFire.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(ventfire(room, j), "0.0"))
                Next room
                If frmprintvar.chkVentFire.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkVentLL.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Vent Flow to Lower (kg/s)", TAB(X), VB6.Format(FlowToLower(room, j), "0.000"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkVentLL.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(FlowToLower(room, j), "0.000"))
                Next room
                If frmprintvar.chkVentLL.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkVentUL.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Vent Flow to Upper (kg/s)", TAB(X), VB6.Format(FlowToUpper(room, j), "0.000"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkVentUL.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(FlowToUpper(room, j), "0.000"))
                Next room
                If frmprintvar.chkVentUL.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkVentOut.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Vent Flow to Outside (m3/s)", TAB(X), VB6.Format(UFlowToOutside(room, j), "0.000"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkVentOut.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(UFlowToOutside(room, j), "0.000"))
                Next room
                If frmprintvar.chkVentOut.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If Visibility(room, j) < 20 Then
                    If frmprintvar.chkVisi.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Visibility (m) at " & MonitorHeight & "m", TAB(X), VB6.Format(Visibility(room, j), "0.00"))
                Else
                    If frmprintvar.chkVisi.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Visibility (m) at " & MonitorHeight & "m", TAB(X), VB6.Format(Visibility(room, j), "0") & "+")
                End If
                For room = 2 To NumberRooms
                    X = X + 10
                    If Visibility(room, j) < 20 Then
                        If frmprintvar.chkVisi.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(Visibility(room, j), "0.00"))
                    Else
                        If frmprintvar.chkVisi.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(Visibility(room, j), "0") & "+")
                    End If
                Next room
                If frmprintvar.chkVisi.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkSPR.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "SPR (m²/s)", TAB(X), VB6.Format(SPR(room, j), "0.00"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkSPR.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(SPR(room, j), "0.00"))
                Next room
                If frmprintvar.chkSPR.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkGER.CheckState = System.Windows.Forms.CheckState.Checked And fireroom = 1 Then
                    Print(1, TAB(10), "GER", TAB(X), VB6.Format(GlobalER(j), "0.0"))
                End If
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkGER.CheckState = System.Windows.Forms.CheckState.Checked And room = fireroom Then
                        Print(1, TAB(X), VB6.Format(GlobalER(j), "0.0"))
                    Else
                        Print(1, TAB(X))
                    End If
                Next room
                If frmprintvar.chkGER.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                If frmprintvar.chkCOYield.CheckState = System.Windows.Forms.CheckState.Checked Then
                    'calculate co yields
                    If tim(j, 1) >= flashover_time Then
                        tempvar = Get_CO_yield(GlobalER(j), comode, True)
                    Else
                        If (tim(j, 1) >= alphaTfitted(4, itcounter - 1)) And (VM2 = True) Then
                            tempvar = Get_CO_yield(GlobalER(j), comode, True)
                        Else
                            tempvar = Get_CO_yield(GlobalER(j), comode, False)
                        End If

                    End If
                End If
                X = Y
                room = 1
                If frmprintvar.chkCOYield.CheckState = System.Windows.Forms.CheckState.Checked And fireroom = 1 Then
                    Print(1, TAB(10), "CO Yield (g/g)", TAB(X), VB6.Format(tempvar, "0.000"))
                End If
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkCOYield.CheckState = System.Windows.Forms.CheckState.Checked And room = fireroom Then
                        Print(1, TAB(X), VB6.Format(tempvar, "0.000"))
                    Else
                        Print(1, TAB(X))
                    End If
                Next room
                If frmprintvar.chkCOYield.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkODupper.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "OD Upper (1/m)", TAB(X), VB6.Format(OD_upper(room, j), "0.000"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkODupper.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(OD_upper(room, j), "0.000"))
                Next room
                If frmprintvar.chkODupper.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkODlower.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "OD Lower (1/m)", TAB(X), VB6.Format(OD_lower(room, j), "0.000"))
                For room = 2 To NumberRooms
                    If TwoZones(room) = True Then
                        X = X + 10
                        If frmprintvar.chkODlower.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(OD_lower(room, j), "0.000"))
                    End If
                Next room
                If frmprintvar.chkODlower.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkExtUpper.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Ext. Coef. Upper (1/m)", TAB(X), VB6.Format(2.3 * OD_upper(room, j), "0.000"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkExtUpper.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(2.3 * OD_upper(room, j), "0.000"))
                Next room
                If frmprintvar.chkExtUpper.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkExtLower.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Ext. Coef. Lower (1/m)", TAB(X), VB6.Format(2.3 * OD_lower(room, j), "0.000"))
                For room = 2 To NumberRooms
                    If TwoZones(room) = True Then
                        X = X + 10
                        If frmprintvar.chkExtLower.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(2.3 * OD_lower(room, j), "0.000"))
                    End If
                Next room
                If frmprintvar.chkExtLower.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)


                X = Y
                room = 1
                If frmprintvar.chkCST.CheckState = System.Windows.Forms.CheckState.Checked Then
                    Print(1, TAB(10), "Ceiling Temp (C)", TAB(X), VB6.Format(CeilingTemp(room, j) - 273, "0.0"))
                End If
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkCST.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(CeilingTemp(room, j) - 273, "0.0"))
                Next room
                If frmprintvar.chkCST.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkCST.CheckState = System.Windows.Forms.CheckState.Checked Then
                    Print(1, TAB(10), "Unexposed Ceiling Temp (C)", TAB(X), VB6.Format(UnexposedCeilingtemp(room, j) - 273, "0.0"))
                End If
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkCST.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(UnexposedCeilingtemp(room, j) - 273, "0.0"))
                Next room
                If frmprintvar.chkCST.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkFST.CheckState = System.Windows.Forms.CheckState.Checked Then
                    Print(1, TAB(10), "Floor Temp (C)", TAB(X), VB6.Format(FloorTemp(room, j) - 273, "0.0"))
                End If
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkFST.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(FloorTemp(room, j) - 273, "0.0"))
                Next room
                If frmprintvar.chkFST.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkFST.CheckState = System.Windows.Forms.CheckState.Checked Then
                    Print(1, TAB(10), "Unexposed Floor Temp (C)", TAB(X), VB6.Format(UnexposedFloortemp(room, j) - 273, "0.0"))
                End If
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkFST.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(UnexposedFloortemp(room, j) - 273, "0.0"))
                Next room
                If frmprintvar.chkFST.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkUWST.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Upper Wall Temp (C)", TAB(X), VB6.Format(Upperwalltemp(room, j) - 273, "0.0"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkUWST.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(Upperwalltemp(room, j) - 273, "0.0"))
                Next room
                If frmprintvar.chkUWST.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkUWST.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Unexposed U Wall Temp (C)", TAB(X), VB6.Format(UnexposedUpperwalltemp(room, j) - 273, "0.0"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkUWST.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(UnexposedUpperwalltemp(room, j) - 273, "0.0"))
                Next room
                If frmprintvar.chkUWST.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkLWST.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Lower Wall Temp (C)", TAB(X), VB6.Format(LowerWallTemp(room, j) - 273, "0.0"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkLWST.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(LowerWallTemp(room, j) - 273, "0.0"))
                Next room
                If frmprintvar.chkLWST.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkLWST.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Unexposed L Wall Temp (C)", TAB(X), VB6.Format(UnexposedLowerwalltemp(room, j) - 273, "0.0"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkLWST.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(UnexposedLowerwalltemp(room, j) - 273, "0.0"))
                Next room
                If frmprintvar.chkLWST.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkTotalLosses.CheckState = CheckState.Checked Then Print(1, TAB(10), "Total heat loss to bound (kW)", TAB(X), VB6.Format(TotalLosses(room, j), "0.0"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkTotalLosses.CheckState = CheckState.Checked Then Print(1, TAB(X), VB6.Format(TotalLosses(room, j), "0.0"))
                Next room
                If frmprintvar.chkTotalLosses.CheckState = CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If calcFRR = True Then Print(1, TAB(10), "Normalised Heat Load, avg (s1/2 K)", TAB(X), VB6.Format(NHL(1, room, j), "0"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If calcFRR = True Then Print(1, TAB(X), VB6.Format(NHL(1, room, j), "0"))
                Next room
                If calcFRR = True Then PrintLine(1)

                X = Y
                room = 1
                If calcFRR = True Then Print(1, TAB(10), "NHL, ceiling (s1/2 K)", TAB(X), VB6.Format(NHL(3, room, j), "0"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If calcFRR = True Then Print(1, TAB(X), VB6.Format(NHL(3, room, j), "0"))
                Next room
                If calcFRR = True Then PrintLine(1)

                X = Y
                room = 1
                If calcFRR = True Then Print(1, TAB(10), "NHL, upper wall (s1/2 K)", TAB(X), VB6.Format(NHL(4, room, j), "0"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If calcFRR = True Then Print(1, TAB(X), VB6.Format(NHL(4, room, j), "0"))
                Next room
                If calcFRR = True Then PrintLine(1)

                X = Y
                room = 1
                If calcFRR = True Then Print(1, TAB(10), "NHL, lower wall (s1/2 K)", TAB(X), VB6.Format(NHL(5, room, j), "0"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If calcFRR = True Then Print(1, TAB(X), VB6.Format(NHL(5, room, j), "0"))
                Next room
                If calcFRR = True Then PrintLine(1)

                X = Y
                room = 1
                If calcFRR = True Then Print(1, TAB(10), "NHL, floor (s1/2 K)", TAB(X), VB6.Format(NHL(6, room, j), "0"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If calcFRR = True Then Print(1, TAB(X), VB6.Format(NHL(6, room, j), "0"))
                Next room
                If calcFRR = True Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkAST.CheckState = CheckState.Checked Then Print(1, TAB(10), "Ceiling AST (C)", TAB(X), Format(QCeilingAST(room, 2, j) - 273, "0.0"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkAST.CheckState = CheckState.Checked Then Print(1, TAB(X), Format(QCeilingAST(room, 2, j) - 273, "0.0"))
                Next room
                If frmprintvar.chkAST.CheckState = CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkAST.CheckState = CheckState.Checked Then Print(1, TAB(10), "Upper Wall AST (C)", TAB(X), Format(QUpperWallAST(room, 2, j) - 273, "0.0"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkAST.CheckState = CheckState.Checked Then Print(1, TAB(X), Format(QUpperWallAST(room, 2, j) - 273, "0.0"))
                Next room
                If frmprintvar.chkAST.CheckState = CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkAST.CheckState = CheckState.Checked Then Print(1, TAB(10), "Lower Wall AST (C)", TAB(X), Format(QLowerWallAST(room, 2, j) - 273, "0.0"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkAST.CheckState = CheckState.Checked Then Print(1, TAB(X), Format(QLowerWallAST(room, 2, j) - 273, "0.0"))
                Next room
                If frmprintvar.chkAST.CheckState = CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkAST.CheckState = CheckState.Checked Then Print(1, TAB(10), "Floor AST (C)", TAB(X), Format(QFloorAST(room, 2, j) - 273, "0.0"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chkAST.CheckState = CheckState.Checked Then Print(1, TAB(X), Format(QFloorAST(room, 2, j) - 273, "0.0"))
                Next room
                If frmprintvar.chkAST.CheckState = CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chk_U_tuhc.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Unburned Fuel Upper (g/g)", TAB(X), VB6.Format(TUHC(room, j, 1), "0.000"))
                For room = 2 To NumberRooms
                    X = X + 10
                    If frmprintvar.chk_U_tuhc.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(TUHC(room, j, 1), "0.000"))
                Next room
                If frmprintvar.chk_U_tuhc.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chk_L_tuhc.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Unburned Fuel Lower (g/g)", TAB(X), VB6.Format(TUHC(room, j, 2), "0.000"))
                For room = 2 To NumberRooms
                    If TwoZones(room) = True Then
                        X = X + 10
                        If frmprintvar.chk_L_tuhc.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(TUHC(room, j, 2), "0.000"))
                    End If
                Next room
                If frmprintvar.chk_L_tuhc.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.ChkCJtemp.CheckState = System.Windows.Forms.CheckState.Checked Then
                    For s = 1 To NumSprinklers
                        If (CJetTemp(j, 1, s - 1)) > 0 Then Print(1, TAB(10), "Ceiling jet temp at device " & s.ToString & " (C)", TAB(X + 10), Format(CJetTemp(j, 1, s - 1) - 273, "0.0"))
                    Next s
                End If
                If frmprintvar.ChkCJtemp.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkCJVel.CheckState = System.Windows.Forms.CheckState.Checked Then
                    For s = 1 To NumSprinklers
                        If (CJetTemp(j, 0, s - 1)) > 0 Then Print(1, TAB(10), "Ceiling Jet Vel at device " & s.ToString & " (m/s)", TAB(X + 10), Format(CJetTemp(j, 0, s - 1), "0.000"))
                    Next s
                End If
                If frmprintvar.chkCJVel.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkCJmax.CheckState = System.Windows.Forms.CheckState.Checked Then
                    For s = 1 To NumSprinklers
                        If (CJetTemp(j, 2, s - 1)) > 0 Then Print(1, TAB(10), "Max Ceiling Jet Temp at device " & s.ToString & " (C)", TAB(X + 10), Format(CJetTemp(j, 2, s - 1) - 273, "0.0"))
                    Next s
                End If
                If frmprintvar.chkCJmax.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                X = Y
                room = 1
                If frmprintvar.chkLink.CheckState = System.Windows.Forms.CheckState.Checked Then
                    For s = 1 To NumSprinklers
                        If (SprinkTemp(s, j)) > 0 Then Print(1, TAB(10), "Link Temp at device " & s.ToString & " (C)", TAB(X + 10), Format(SprinkTemp(s, j) - 273, "0.0"))
                    Next s
                End If
                If frmprintvar.chkLink.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                'X = Y
                'room = 1
                '    If frmprintvar.chkCJVel.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Ceiling Jet Vel at Link (m/s)", TAB(X), VB6.Format(CJetTemp(j, 0, 0), "0.000"))
                '    For room = 2 To NumberRooms
                '        X = X + 10
                '        If frmprintvar.chkCJVel.CheckState = System.Windows.Forms.CheckState.Checked And room = fireroom Then Print(1, TAB(X), VB6.Format(CJetTemp(j, 0, 0), "0.000"))
                '    Next room
                '    If frmprintvar.chkCJVel.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)


                '    X = Y
                '    room = 1
                '    If frmprintvar.chkCJmax.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Max Ceiling Jet Temp at Link (C)", TAB(X), VB6.Format(CJetTemp(j, 2, 0) - 273, "0.0"))
                '    For room = 2 To NumberRooms
                '        X = X + 10
                '        If frmprintvar.chkCJmax.CheckState = System.Windows.Forms.CheckState.Checked And room = fireroom Then Print(1, TAB(X), VB6.Format(CJetTemp(j, 2, 0) - 273, "0.0"))
                '    Next room
                '    If frmprintvar.chkCJmax.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                '    X = Y
                '    room = 1
                '    'If frmprintvar.chkLink.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Link Temp(C)", TAB(X), VB6.Format(LinkTemp(room, j) - 273, "0.0"))
                '    If frmprintvar.chkLink.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Link Temp(C)", TAB(X), VB6.Format(LinkTemp(room, j) - 273, "0.0"))
                '    For room = 2 To NumberRooms
                '        X = X + 10
                '        If frmprintvar.chkLink.CheckState = System.Windows.Forms.CheckState.Checked And room = fireroom Then Print(1, TAB(X), VB6.Format(LinkTemp(room, j) - 273, "0.0"))
                '    Next room
                '    If frmprintvar.chkLink.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                If frmOptions1.optQuintiere.Checked = True Then
                    X = Y
                    room = 1
                    If frmprintvar.chkXPyrol.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "X Pyrolysis (m)", TAB(X), VB6.Format(X_pyrolysis(room, j), "0.000"))
                    For room = 2 To NumberRooms
                        X = X + 10
                        If frmprintvar.chkXPyrol.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(X_pyrolysis(room, j), "0.000"))
                    Next room
                    If frmprintvar.chkXPyrol.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                    X = Y
                    room = 1
                    If frmprintvar.chkYPyrol.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Y Pyrolysis (m)", TAB(X), VB6.Format(Y_pyrolysis(room, j), "0.000"))
                    For room = 2 To NumberRooms
                        X = X + 10
                        If frmprintvar.chkYPyrol.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(Y_pyrolysis(room, j), "0.000"))
                    Next room
                    If frmprintvar.chkYPyrol.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)

                    X = Y
                    room = 1
                    If frmprintvar.chkZPyrol.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(10), "Z Pyrolysis (m)", TAB(X), VB6.Format(Z_pyrolysis(room, j), "0.000"))
                    For room = 2 To NumberRooms
                        X = X + 10
                        If frmprintvar.chkZPyrol.CheckState = System.Windows.Forms.CheckState.Checked Then Print(1, TAB(X), VB6.Format(Z_pyrolysis(room, j), "0.000"))
                    Next room
                    If frmprintvar.chkZPyrol.CheckState = System.Windows.Forms.CheckState.Checked Then PrintLine(1)
                End If

                If frmprintvar.chkFED.CheckState = System.Windows.Forms.CheckState.Checked Then
                    PrintLine(1, TAB(10), "FED gases on egress path = " & VB6.Format(FEDSum(fireroom, j), "0.000"))
                End If

                If frmprintvar.chkFED.CheckState = System.Windows.Forms.CheckState.Checked Then
                    PrintLine(1, TAB(10), "FED thermal on egress path = " & VB6.Format(FEDRadSum(fireroom, j), "0.000"))
                End If

                If calcFRR = True Then
                    Call FRating(tim(j, 1), energy, fireroom)
                    Call ISO_energy(energy, frr)
                    'PrintLine(1, TAB(10), "Equivalent thermal exposure in fire resistance test furnace based on VM 2.4.4 (Km=1.0) = " & VB6.Format(TEeurocode, "0") & " minutes.")
                    PrintLine(1, TAB(10), "Equivalent thermal exposure in fire resistance test furnace based on emissive power = " & VB6.Format(frr, "0") & " minutes.")
                    PrintLine(1, TAB(10), "Equivalent thermal exposure in fire resistance test furnace based on NHL (Harmathy Furnace)= " & VB6.Format(NHL(2, fireroom, j), "0") & " minutes.")
                    PrintLine(1, TAB(10), "Equivalent thermal exposure in fire resistance test furnace based on ceiling NHL = " & VB6.Format(NHLte(1, j), "0") & " minutes.")
                    PrintLine(1, TAB(10), "Equivalent thermal exposure in fire resistance test furnace based on wall NHL = " & VB6.Format(NHLte(2, j), "0") & " minutes.")
                    PrintLine(1, TAB(10), "Equivalent thermal exposure in fire resistance test furnace based on floor NHL = " & VB6.Format(NHLte(4, j), "0") & " minutes.")
                End If

                PrintLine(1)

                j = j + DisplayInterval / Timestep
            Next k
        End If
        If calcFRR = True Then
            PrintLine(1, TAB(10), "Equivalent thermal exposure in fire resistance test furnace based on NZBC C/VM2 2.4.4 (Km=1.0) = " & VB6.Format(TEeurocode, "0") & " minutes.")
        End If
        PrintLine(1)
        PrintLine(1, "====================================================================")
        PrintLine(1, "Event Log")
        PrintLine(1, "====================================================================")

        PrintLine(1, frmInputs.rtb_log.Text)

        PrintLine(1)
        PrintLine(1, "====================================================================")
        PrintLine(1, "Computer Run-Time = " & VB6.Format(runtime, "0.0") & " seconds.")
        PrintLine(1, "====================================================================")
        FileClose(1)
        Exit Sub

errorhandler:
        'Err.Clear
        FileClose(1)
        Exit Sub

    End Sub

    Public Sub Graph_data_ventfire(ByRef room As Short, ByRef Title As String, ByRef datatobeplotted(,) As Double, ByRef DataShift As Double, ByRef DataMultiplier As Double, ByRef GraphSty As Short, ByRef MaxYValue As Double)
        '*  ====================================================================
        '*  This function takes data for a variable from a two-dimensional array
        '*  and displays it in a graph
        '*
        '*  Revised 13/1/98 Colleen Wade
        '*  ====================================================================
        Dim i As Integer, j As Integer

        'if no data exists
        If NumberTimeSteps < 1 Then
            MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Sub
        End If

        Dim chdata(0 To NumberTimeSteps, 0 To 2 * NumberRooms + 1) As Object

        ' With frmGraph.AxMSChart1
        '.chartType = MSChart20Lib.VtChChartType.VtChChartType2dXY
        ' .Visible = True
        ' .TitleText = Left(Description, 73)
        ' .Plot.UniformAxis = False
        ' .Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).AxisTitle.Text = "Time (sec)"
        ' .Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).AxisTitle.Text = Title
        ' .Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).AxisTitle.TextLayout.Orientation = MSChart20Lib.VtOrientation.VtOrientationHorizontal

        room = 1
        For i = 0 To 2 * NumberRooms + 1 Step 2
            If room = NumberRooms + 1 Then
                chdata(0, i) = "Outside"
            Else
                chdata(0, i) = "Room " & CStr(room)
            End If

            For j = 1 To NumberTimeSteps
                chdata(j, i) = tim(j, 1) 'time
                chdata(j, i + 1) = datatobeplotted(room, j) * DataMultiplier + DataShift 'data to be plotted
            Next
            If room = NumberRooms + 1 Then chdata(0, i) = "Outside"
            room = room + 1
        Next i

        ' .ChartData = chdata

        'End With
        frmGraph.Text = Title
        frmGraph.Show()
        '    Dim graph As Control
        '    Dim i As Integer, j As Integer
        '    Dim maxpoints As Long, factor As Integer
        '    On Error GoTo graphhandler
        '
        '    'define an object variable for this control
        '    Set graph = frmGraphs.Graph1
        '    graph.Visible = True
        '
        '    'if no data exists
        '    If NumberTimeSteps < 1 Then
        '        MsgBox "There is no data to plot, please run the simulation first.", vbExclamation
        '        Exit Sub
        '    End If
        '
        '    frmGraphs.Show
        '    graph.AutoInc = 1
        '    graph.LeftTitle = Title
        '    graph.GraphTitle = Left$(Description, 80)
        '    graph.ThickLines = 1
        '    graph.FontStyle = 0
        '    graph.Ticks = 1
        '    graph.FontUse = gphAllText
        '    graph.FontFamily = gphSwiss
        '    graph.BottomTitle = "Time (sec)"
        '    graph.GraphStyle = 4
        '    graph.LineStats = 0
        '    graph.DataReset = 9
        '
        '    factor = 1
        '    maxpoints = (NumberRooms + 1) * (NumberTimeSteps + 1)
        '    Do While maxpoints >= 3800 Or (NumberTimeSteps + 1) / factor > 1200
        '        'too many data points
        '        factor = factor * 2
        '        maxpoints = maxpoints / factor
        '    Loop
        '
        '    graph.NumSets = NumberRooms + 1
        '    graph.NumPoints = (NumberTimeSteps + 1) / factor
        '
        '    If GraphSty = 2 Then
        '        graph.YAxisMax = MaxYValue
        '        graph.YAxisTicks = 9
        '    End If
        '
        'here:
        '
        '    For j = 1 To graph.NumSets
        '        graph.Graphdata = (DataMultiplier * datatobeplotted(j, 1) + DataShift)
        '        For i% = 2 To graph.NumPoints
        '            graph.Graphdata = (DataMultiplier * datatobeplotted(j, factor * i%) + DataShift)
        '        Next i%
        '    Next j
        '    For j = 1 To graph.NumSets
        '        graph.XPosData = (tim(1, 1))
        '        For i% = 2 To graph.NumPoints
        '            graph.XPosData = (tim(factor * i%, 1))
        '        Next i%
        '    Next j
        '
        '    For j = 1 To NumberRooms
        '        graph.LegendText = "Room " + CStr(j)
        '    Next j
        '    graph.LegendText = "Outside"
        '    graph.DrawMode = 2
        '    Exit Sub
        '
        'graphhandler:
        '    graph.NumPoints = graph.NumPoints - 1
        '    GoTo here
        'Exit Sub
        ''
    End Sub

    Public Sub AdjacentRoom_Ignition(ByRef qrad As Double, ByRef ventfire As Double, ByRef vwidth As Single, ByRef layer1 As Double, ByRef ceilingheight2 As Double, ByRef flameheight As Double, ByRef projection As Double)
        '==========================================================
        '   calculate the lining heat fluxes to use in determining
        '   ignition in a adjacent room. Assumes ignition is a result
        '   of a vent fire exposing the ceiling in the adjacent room.
        '
        '   qrad = the radiant flux per unit area  kw/m2
        '   ventfire = size of vent fire kw
        '   vwidth = width of the vent, taken as diameter of fire m
        '   layer1 = layer height in originating room m
        '   ceilingheight2 = ceiling height in destination room m
        '
        '   C A Wade 26/3/99
        '==========================================================

        Dim qstar, SurfaceArea As Double

        'estimate the flame height of a vent fire using a hasemi wall correlation
        qstar = ventfire / (1110 * vwidth ^ (5 / 2))
        If stepcount > 1 Then qstar = (ventfire - SpecificHeat_air * uppertemp(fireroom, stepcount - 1) * FlowToUpper(fireroom, stepcount - 1)) / (1110 * vwidth ^ (5 / 2))
        flameheight = 2.8 * vwidth * (qstar / vwidth) ^ (2 / 3)
        If flameheight <= 0 Then Exit Sub

        'SurfaceArea = vwidth * flameheight + Pi * vwidth / 2 + Pi * vwidth ^ 2 / 4
        SurfaceArea = 2 * (vwidth * flameheight + projection * flameheight + projection * vwidth)

        qrad = ObjectRLF(1) * ventfire / SurfaceArea
        qrad = ObjectRLF(1) * (ventfire - SpecificHeat_air * uppertemp(fireroom, stepcount - 1) * FlowToUpper(fireroom, stepcount - 1)) / SurfaceArea

    End Sub

    Public Function CeilingJet_MaxTemp_Adj(ByVal Q As Double, ByVal room As Integer, ByVal layer1 As Double) As Double
        '*  =============================================================
        '*      This function returns the value of the maximum temperature
        '*      in the ceiling jet at the location of the sprinkler or
        '*      detector using the Alpert quasi-steady correlations.
        '*
        '*      Arguments passed to the function:
        '*      Q = vent fire
        '       room = destination room
        '*  =============================================================

        Dim h As Single

        'ceiling height in destination room above layer height in source room
        h = RoomHeight(room) - layer1
        If h > 0 Then
            'against wall
            'UPGRADE_WARNING: Couldn't resolve default property of object CeilingJet_MaxTemp_Adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            CeilingJet_MaxTemp_Adj = 16.9 * (2 * Q) ^ (2 / 3) * h ^ (-5 / 3) + uppertemp(room, stepcount)
        Else
            CeilingJet_MaxTemp_Adj = uppertemp(room, stepcount)
        End If
    End Function

    Public Function Graph_Data_2D_reverse(ByVal thisroom As Short, ByVal Title As String, ByVal datatobeplotted(,) As Double, ByVal DataShift As Double, ByVal DataMultiplier As Double, ByVal timeunit As Integer) As Double
        '*  ====================================================================
        '*  This function takes data for a variable from a two-dimensional array
        '*  and displays it in a graph. Only plots one room.
        '*  ====================================================================
        Dim j As Integer
        Dim room As Integer
        Dim ydata(0 To NumberTimeSteps) As Double

        'if no data exists
        If NumberTimeSteps < 1 Then
            MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Function
        End If

        Try

            frmPlot.Chart1.Series.Clear()
            room = 1
            For room = 1 To NumberRooms

                frmPlot.Chart1.Series.Add("Room " & room)

                frmPlot.Chart1.Series("Room " & room).ChartType = SeriesChartType.FastLine

                If Not roomcolor(room - 1).IsEmpty Then
                    frmPlot.Chart1.Series("Room " & room).Color = roomcolor(room - 1) '=line color
                End If

                For j = 1 To NumberTimeSteps
                    ydata(j) = datatobeplotted(j, room) * DataMultiplier + DataShift 'data to be plotted
                    frmPlot.Chart1.Series("Room " & room).Points.AddXY(tim(j, 1) / timeunit, ydata(j))
                Next

            Next room
            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            'frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            If timeunit = 60 Then
                frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (min)"
            Else
                frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            End If
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            'frmPlot.Chart1.Titles("Title1").Text = Title

            'frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Graph_Data_2D_reverse")
        End Try

    End Function

    Public Sub View_input(ByRef filename As String)
        '*  ====================================================================
        '*  View the results in a rich text box
        '*  ====================================================================

        Dim j, i, k As Short
        Dim upperroom, lowerroom As Integer
        ReDim Preserve CeilingIgniteFlag(NumberRooms)
        Dim oSprinklers As New List(Of oSprinkler)
        oSprinklers = SprinklerDB.GetSprinklers2
        Dim osprdistributions As New List(Of oDistribution)
        osprdistributions = SprinklerDB.GetSprDistributions
        Dim oSmokeDetects As New List(Of oSmokeDet)
        oSmokeDetects = SmokeDetDB.GetSmokDets
        Dim osddistributions As New List(Of oDistribution)
        osddistributions = SmokeDetDB.GetSDDistributions
        Dim ovents As New List(Of oVent)
        ovents = VentDB.GetVents
        Dim oventdistributions As New List(Of oDistribution)
        oventdistributions = VentDB.GetVentDistributions
        Dim ocvents As New List(Of oCVent)
        ocvents = VentDB.GetCVents
        Dim ocventdistributions As New List(Of oDistribution)
        ocventdistributions = VentDB.GetCVentDistributions
        Dim ofans As New List(Of oFan)
        ofans = FanDB.GetFans
        Dim ofandistributions As New List(Of oDistribution)
        ofandistributions = FanDB.GetFanDistributions

        Derived_Variables()
        'get endpoint captions
        FileOpen(1, filename, OpenMode.Output)

        PrintLine(1)
        ' PrintLine(1, Format(Now, "dddd,mmmm dd,yyyy,hh:mm AM/PM"))
        PrintLine(1, Now.ToString("MMMM dd, yyyy H:mm:ss"))
        PrintLine(1, "B-RISK Fire Simulator and Design Fire Tool (Ver " & Version & ")")

        PrintLine(1)
        PrintLine(1, "Input Filename : " & DataFile)
        PrintLine(1, "Base File : " & basefile)
        PrintLine(1)
        If VM2 = True Then
            PrintLine(1, "User Mode : C/VM2")
        Else
            PrintLine(1, "User Mode : Risk Simulator")
        End If
        If TalkToEVACNZ = True Then
            PrintLine(1, "EvacuatioNZ is used to open/close vents.")
        End If
        PrintLine(1, "Simulation Time = " & VB6.Format(SimTime, "0.00") & " seconds.")
        PrintLine(1, "Initial Time-Step = " & VB6.Format(Timestep, "0.00") & " seconds.")
        PrintLine(1)
        PrintLine(1, Description)
        PrintLine(1)
        PrintLine(1, "====================================================================")
        PrintLine(1, "Description of Rooms")
        PrintLine(1, "====================================================================")

        For j = 1 To NumberRooms
            Print(1, "Room " & j & " : " & RoomDescription(j))
            If TwoZones(j) = False Then Print(1, TAB(10), "Room is modelled as a single zone.")
            PrintLine(1, TAB(10), "Room Length (m) =", TAB(60), VB6.Format(RoomLength(j), "0.00"))
            PrintLine(1, TAB(10), "Room Width (m) =", TAB(60), VB6.Format(RoomWidth(j), "0.00"))
            PrintLine(1, TAB(10), "Maximum Room Height (m) =", TAB(60), VB6.Format(RoomHeight(j), "0.00"))
            PrintLine(1, TAB(10), "Minimum Room Height (m) =", TAB(60), VB6.Format(MinStudHeight(j), "0.00"))
            PrintLine(1, TAB(10), "Floor Elevation (m) =", TAB(60), VB6.Format(FloorElevation(j), "0.000"))
            PrintLine(1, TAB(10), "Absolute X Position (m) =", TAB(60), VB6.Format(RoomAbsX(j), "0.000"))
            PrintLine(1, TAB(10), "Absolute Y Position (m) =", TAB(60), VB6.Format(RoomAbsY(j), "0.000"))

            If CeilingSlope(j) = False Then
                PrintLine(1, TAB(10), "Room " & j & " has a flat ceiling.")
            Else
                PrintLine(1, TAB(10), "Room " & j & " has a sloping ceiling.")
            End If

            'shape factor using average ceiling height
            PrintLine(1, TAB(10), "Shape Factor (Af/H^2) =", TAB(60), VB6.Format(RoomLength(j) * RoomWidth(j) / ((RoomHeight(j) / 2 + MinStudHeight(j) / 2) ^ 2), "0.0"))
            PrintLine(1)

            PrintLine(1, TAB(10), "Wall Surface is " & WallSurface(j))
            PrintLine(1, TAB(10), "Wall Density (kg/m3) =", TAB(60), VB6.Format(WallDensity(j), "0.0"))
            PrintLine(1, TAB(10), "Wall Conductivity (W/m.K) =", TAB(60), VB6.Format(WallConductivity(j), "0.000"))
            PrintLine(1, TAB(10), "Wall Specific Heat (J/kg.K) =", TAB(60), VB6.Format(WallSpecificHeat(j), "0"))
            PrintLine(1, TAB(10), "Wall Emissivity =", TAB(60), VB6.Format(Surface_Emissivity(2, j), "0.00"))
            PrintLine(1, TAB(10), "Wall Thickness (mm) =", TAB(60), VB6.Format(WallThickness(j), "0.0"))
            PrintLine(1, TAB(10), "SQROOT Thermal Inertia (J.m-2.s-1/2.K-1) =", TAB(60), VB6.Format(Sqrt(ThermalInertiaWall(j) * 10 ^ 6), "0"))
            PrintLine(1)

            If HaveWallSubstrate(j) = True Then
                PrintLine(1, TAB(10), "Wall Substrate is " & WallSubstrate(j))
                PrintLine(1, TAB(10), "Wall Substrate Density (kg/m3) =", TAB(60), VB6.Format(WallSubDensity(j), "0.0"))
                PrintLine(1, TAB(10), "Wall Substrate Conductivity (W/m.K) =", TAB(60), VB6.Format(WallSubConductivity(j), "0.000"))
                PrintLine(1, TAB(10), "Wall Substrate Specific Heat (J/kg.K) =", TAB(60), VB6.Format(WallSubSpecificHeat(j), "0"))
                PrintLine(1, TAB(10), "Wall Substrate Thickness (mm) =", TAB(60), VB6.Format(WallSubThickness(j), "0.0"))
                PrintLine(1)
            End If

            PrintLine(1, TAB(10), "Ceiling Surface is " & CeilingSurface(j))
            PrintLine(1, TAB(10), "Ceiling Density (kg/m3) =", TAB(60), VB6.Format(CeilingDensity(j), "0.0"))
            PrintLine(1, TAB(10), "Ceiling Conductivity (W/m.K) =", TAB(60), VB6.Format(CeilingConductivity(j), "0.000"))
            PrintLine(1, TAB(10), "Ceiling Specific Heat (J/kg.K) =", TAB(60), VB6.Format(CeilingSpecificHeat(j), "0"))
            PrintLine(1, TAB(10), "Ceiling Emissivity =", TAB(60), VB6.Format(Surface_Emissivity(1, j), "0.00"))
            PrintLine(1, TAB(10), "Ceiling Thickness (mm) =", TAB(60), VB6.Format(CeilingThickness(j), "0.0"))
            PrintLine(1, TAB(10), "SQROOT Thermal Inertia (J.m-2.s-1/2.K-1) =", TAB(60), VB6.Format(Sqrt(ThermalInertiaCeiling(j) * 10 ^ 6), "0"))

            PrintLine(1)

            If HaveCeilingSubstrate(j) = True Then
                PrintLine(1, TAB(10), "Ceiling Substrate is " & CeilingSubstrate(j))
                PrintLine(1, TAB(10), "Ceiling Substrate Density (kg/m3) =", TAB(60), VB6.Format(CeilingSubDensity(j), "0.0"))
                PrintLine(1, TAB(10), "Ceiling Substrate Conductivity (W/m.K) =", TAB(60), VB6.Format(CeilingSubConductivity(j), "0.000"))
                PrintLine(1, TAB(10), "Ceiling Substrate Specific Heat (J/kg.K) =", TAB(60), VB6.Format(CeilingSubSpecificHeat(j), "0"))
                PrintLine(1, TAB(10), "Ceiling Substrate Thickness (mm) =", TAB(60), VB6.Format(CeilingSubThickness(j), "0.0"))
                PrintLine(1)
            End If

            PrintLine(1, TAB(10), "Floor Surface is " & FloorSurface(j))
            PrintLine(1, TAB(10), "Floor Density (kg/m3) =", TAB(60), VB6.Format(FloorDensity(j), "0.0"))
            PrintLine(1, TAB(10), "Floor Conductivity (W/m.K) =", TAB(60), VB6.Format(FloorConductivity(j), "0.000"))
            PrintLine(1, TAB(10), "Floor Specific Heat (J/kg.K) =", TAB(60), VB6.Format(FloorSpecificHeat(j), "0"))
            PrintLine(1, TAB(10), "Floor Emissivity =", TAB(60), VB6.Format(Surface_Emissivity(4, j), "0.00"))
            PrintLine(1, TAB(10), "Floor Thickness = (mm) ", TAB(60), VB6.Format(FloorThickness(j), "0.0"))
            PrintLine(1, TAB(10), "SQROOT Thermal Inertia (J.m-2.s-1/2.K-1) =", TAB(60), VB6.Format(Sqrt(ThermalInertiaFloor(j) * 10 ^ 6), "0"))

            PrintLine(1)

            If HaveFloorSubstrate(j) = True Then
                PrintLine(1, TAB(10), "Floor Substrate is " & FloorSubstrate(j))
                PrintLine(1, TAB(10), "Floor Substrate Density (kg/m3) =", TAB(60), VB6.Format(FloorSubDensity(j), "0.0"))
                PrintLine(1, TAB(10), "Floor Substrate Conductivity (W/m.K) =", TAB(60), VB6.Format(FloorSubConductivity(j), "0.000"))
                PrintLine(1, TAB(10), "Floor Substrate Specific Heat (J/kg.K) =", TAB(60), VB6.Format(FloorSubSpecificHeat(j), "0"))
                PrintLine(1, TAB(10), "Floor Substrate Thickness (mm) =", TAB(60), VB6.Format(FloorSubThickness(j), "0.0"))
                PrintLine(1)
            End If

        Next j
        PrintLine(1, "====================================================================")
        PrintLine(1, "Wall Vents")
        PrintLine(1, "====================================================================")

        For Each oVent In ovents
            PrintLine(1, "Vent  " & oVent.id & " : " & oVent.description)
            If i = NumberRooms + 1 Then
                PrintLine(1, TAB(20), "From room " & oVent.fromroom & " to outside")
            Else
                PrintLine(1, TAB(20), "From room " & oVent.fromroom & " to " & oVent.toroom)
            End If
            Select Case oVent.face
                Case 0
                    PrintLine(1, TAB(20), "Front face of room " & oVent.fromroom.ToString)
                Case 1
                    PrintLine(1, TAB(20), "Right  face of room " & oVent.fromroom.ToString)
                Case 2
                    PrintLine(1, TAB(20), "Rear  face of room " & oVent.fromroom.ToString)
                Case 3
                    PrintLine(1, TAB(20), "Left  face of room " & oVent.fromroom.ToString)
            End Select
            PrintLine(1, TAB(20), "Offset (m) =", TAB(60), VB6.Format(oVent.offset, "0.000"))
            PrintLine(1, TAB(20), "Vent Width (m) =", TAB(60), VB6.Format(oVent.width, "0.000"))
            PrintLine(1, TAB(20), "Vent Height (m) =", TAB(60), VB6.Format(oVent.height, "0.000"))
            PrintLine(1, TAB(20), "Vent Sill Height (m) =", TAB(60), VB6.Format(oVent.sillheight, "0.000"))
            PrintLine(1, TAB(20), "Vent Soffit Height (m) =", TAB(60), VB6.Format(oVent.sillheight + oVent.height, "0.000"))
            PrintLine(1, TAB(20), "Opening Time (sec) =", TAB(60), VB6.Format(oVent.opentime, "0"))
            PrintLine(1, TAB(20), "Closing Time (sec) =", TAB(60), VB6.Format(oVent.closetime, "0"))
            PrintLine(1, TAB(20), "Discharge Coefficient (-) =", TAB(60), VB6.Format(oVent.cd, "0.000"))

            If oVent.spillplumemodel = 1 Then
                PrintLine(1, TAB(20), "Vent Type is 2D adhered spill plume")
            ElseIf oVent.spillplumemodel = 2 Then
                PrintLine(1, TAB(20), "Vent Type is 2D balcony spill plume")
            ElseIf oVent.spillplumemodel = 3 Then
                PrintLine(1, TAB(20), "Vent Type is 3D adhered spill plume")
            ElseIf oVent.spillplumemodel = 4 Then
                PrintLine(1, TAB(20), "Vent Type is 3D channelled balcony spill plume")
            ElseIf oVent.spillplumemodel = 5 Then
                PrintLine(1, TAB(20), "Vent Type is 3D unchannelled balcony spill plume")
            Else
            End If
            If oVent.spillplumemodel > 0 Then
                PrintLine(1, TAB(20), "Balcony Width (m) =", TAB(60), VB6.Format(oVent.spillbalconyprojection, "0.000"))
                PrintLine(1, TAB(20), "Downstand Depth (m) =", TAB(60), VB6.Format(oVent.downstand, "0.000"))
            End If

            If oVent.autoopenvent = True Then '2009.20
                PrintLine(1, TAB(20), "This vent can automatically open with the following trigger(s)")
                If oVent.triggerSD = True Then
                    PrintLine(1, TAB(25), "Smoke Detection in Room " & CStr(oVent.sdtriggerroom))
                    PrintLine(1, TAB(30), "Vent opens  " & CStr(oVent.triggerventopendelay) & " sec after detector responds for a duration of " & CStr(oVent.triggerventopenduration) & " sec.")
                End If
                If oVent.triggerHD = True Then
                    PrintLine(1, TAB(25), "Sprinkler in Room " & CStr(fireroom))
                    PrintLine(1, TAB(30), "Vent opens  " & CStr(oVent.triggerventopendelay) & " sec after sprinkler responds for a duration of " & CStr(oVent.triggerventopenduration) & " sec.")
                End If
                If oVent.triggerHRR = True Then
                    PrintLine(1, TAB(25), "When fire size reaches " & oVent.HRRthreshold & " kW + " & oVent.HRRventopendelay & " sec and remaining open for " & oVent.HRRventopenduration & " sec.")
                End If
                If oVent.triggerFO = True Then
                    PrintLine(1, TAB(25), "At flashover")
                End If
                If oVent.triggerVL = True Then
                    PrintLine(1, TAB(25), "At ventilation limit")
                End If
                If oVent.triggerHO = True Then
                    PrintLine(1, TAB(20), "This vent fitted with hold open device to automatically close on ")
                    PrintLine(1, TAB(20), "smoke detection in Room " & CStr(oVent.sdtriggerroom))
                    PrintLine(1, TAB(20), "Hold open device reliability = " & CStr(oVent.horeliability))
                End If
            End If

            If oVent.autobreakglass = True Then
                PrintLine(1, TAB(20), "Glass fracture is modelled for this vent")
                PrintLine(1, TAB(20), "Glass Thickness (mm) =", TAB(60), VB6.Format(oVent.glassthickness, "0.0"))
                PrintLine(1, TAB(20), "Glass Fracture to Fallout Time (sec) =", TAB(60), VB6.Format(oVent.glassfalloutime, "0"))
                PrintLine(1, TAB(20), "Glass Shading Depth (mm) =", TAB(60), VB6.Format(oVent.glassshading, "0.0"))
                PrintLine(1, TAB(20), "Glass Fracture Stress (MPa) =", TAB(60), VB6.Format(oVent.glassbreakingstress, "0"))
                PrintLine(1, TAB(20), "Glass Expansion Coefficient (/C) =", TAB(60), VB6.Format(oVent.glassexpansion))
                PrintLine(1, TAB(20), "Glass Conductivity (W/mK) =", TAB(60), VB6.Format(oVent.glassconductivity, "0.00"))
                PrintLine(1, TAB(20), "Glass Diffusivity (m2/s) =", TAB(60), VB6.Format(oVent.glassalpha))
                PrintLine(1, TAB(20), "Glass Modulus (MPa) =", TAB(60), VB6.Format(oVent.glassYoungsModulus, "0"))
                If oVent.glassflameflux = False Then
                    PrintLine(1, TAB(20), "Glass is heated by gas layers only.")
                Else
                    PrintLine(1, TAB(20), "Glass is heated by gas layer and flame radiation.")
                    PrintLine(1, TAB(20), "Glass to Flame Distance (m) =", TAB(60), VB6.Format(oVent.glassdistance, "0.000"))
                End If
            End If
            PrintLine(1)

        Next

        'For j = 1 To NumberRooms
        '    For i = 1 To NumberRooms + 1
        '        If NumberVents(j, i) > 0 Then
        '            If j < i Then
        '                For k = 1 To NumberVents(j, i)
        '                    'PrintLine(1, CStr(ventdescription(j, i, k)))

        '                    If i = NumberRooms + 1 Then
        '                        PrintLine(1, "From room " & j & " to outside" & ", Vent No " & CStr(k))
        '                    Else
        '                        PrintLine(1, "From room " & j & " to " & i & ", Vent No " & CStr(k))
        '                    End If
        '                    PrintLine(1, TAB(20), "Vent Width (m) =", TAB(60), VB6.Format(VentWidth(j, i, k), "0.000"))
        '                    PrintLine(1, TAB(20), "Vent Height (m) =", TAB(60), VB6.Format(VentHeight(j, i, k), "0.000"))
        '                    PrintLine(1, TAB(20), "Vent Sill Height (m) =", TAB(60), VB6.Format(VentSillHeight(j, i, k), "0.000"))
        '                    PrintLine(1, TAB(20), "Vent Soffit Height (m) =", TAB(60), VB6.Format(soffitheight(j, i, k), "0.000"))
        '                    PrintLine(1, TAB(20), "Opening Time (sec) =", TAB(60), VB6.Format(VentOpenTime(j, i, k), "0"))
        '                    PrintLine(1, TAB(20), "Closing Time (sec) =", TAB(60), VB6.Format(VentCloseTime(j, i, k), "0"))
        '                    PrintLine(1, TAB(20), "Flow Coefficient (sec) =", TAB(60), VB6.Format(VentCD(j, i, k), "0.000"))

        '                    If spillplumemodel(j, i, k) = 1 Then
        '                        PrintLine(1, TAB(20), "Vent Type is 2D adhered spill plume")
        '                    ElseIf spillplumemodel(j, i, k) = 2 Then
        '                        PrintLine(1, TAB(20), "Vent Type is 2D balcony spill plume")
        '                    ElseIf spillplumemodel(j, i, k) = 3 Then
        '                        PrintLine(1, TAB(20), "Vent Type is 3D adhered spill plume")
        '                    ElseIf spillplumemodel(j, i, k) = 4 Then
        '                        PrintLine(1, TAB(20), "Vent Type is 3D channelled balcony spill plume")
        '                        PrintLine(1, TAB(20), "Balcony Width (m) =", TAB(60), VB6.Format(spillbalconyprojection(j, i, k), "0.000"))
        '                        PrintLine(1, TAB(20), "Downstand Depth (m) =", TAB(60), VB6.Format(Downstand(j, i, k), "0.000"))
        '                    ElseIf spillplumemodel(j, i, k) = 5 Then
        '                        PrintLine(1, TAB(20), "Vent Type is 3D unchannelled balcony spill plume")
        '                        PrintLine(1, TAB(20), "Balcony Width (m) =", TAB(60), VB6.Format(spillbalconyprojection(j, i, k), "0.000"))
        '                        PrintLine(1, TAB(20), "Downstand Depth (m) =", TAB(60), VB6.Format(Downstand(j, i, k), "0.000"))
        '                    Else
        '                    End If

        '                    If AutoOpenVent(j, i, k) = True Then '2009.20
        '                        PrintLine(1, TAB(20), "This vent can automatically open with the following trigger(s)")
        '                        If trigger_device(0, j, i, k) = True Then
        '                            PrintLine(1, TAB(20), "Smoke Detection in Room " & CStr(SDtriggerroom(j, i, k)))
        '                        End If
        '                        If trigger_device(1, j, i, k) = True Then
        '                            PrintLine(1, TAB(20), "Sprinkler in Room " & CStr(fireroom))
        '                        End If
        '                        If trigger_device(2, j, i, k) = True Then
        '                            PrintLine(1, TAB(20), "When fire size reaches " & HRR_threshold(i, j, k) & " kW + " & HRR_ventopendelay(i, j, k) & " sec and remaining open for " & HRR_ventopenduration(i, j, k) & " sec.")
        '                        End If
        '                        If trigger_device(3, j, i, k) = True Then
        '                            PrintLine(1, TAB(20), "At flashover")
        '                        End If
        '                        If trigger_device(4, j, i, k) = True Then
        '                            PrintLine(1, TAB(20), "At ventilation limit")
        '                        End If

        '                    End If

        '                    If AutoBreakGlass(j, i, k) = True Then
        '                        PrintLine(1, TAB(20), "Glass fracture is modelled for this vent")
        '                        PrintLine(1, TAB(20), "Glass Thickness (mm) =", TAB(60), VB6.Format(GLASSthickness(j, i, k), "0.0"))
        '                        PrintLine(1, TAB(20), "Glass Fracture to Fallout Time (sec) =", TAB(60), VB6.Format(GLASSFalloutTime(j, i, k), "0"))
        '                        PrintLine(1, TAB(20), "Glass Shading Depth (mm) =", TAB(60), VB6.Format(GLASSshading(j, i, k), "0.0"))
        '                        PrintLine(1, TAB(20), "Glass Fracture Stress (MPa) =", TAB(60), VB6.Format(GLASSbreakingstress(j, i, k), "0"))
        '                        PrintLine(1, TAB(20), "Glass Expansion Coefficient (/C) =", TAB(60), VB6.Format(GLASSexpansion(j, i, k)))
        '                        PrintLine(1, TAB(20), "Glass Conductivity (W/mK) =", TAB(60), VB6.Format(GLASSconductivity(j, i, k), "0.00"))
        '                        PrintLine(1, TAB(20), "Glass Diffusivity (m2/s) =", TAB(60), VB6.Format(GLASSalpha(j, i, k)))
        '                        PrintLine(1, TAB(20), "Glass Modulus (MPa) =", TAB(60), VB6.Format(GlassYoungsModulus(j, i, k), "0"))
        '                        If GlassFlameFlux(j, i, k) = False Then
        '                            PrintLine(1, TAB(20), "Glass is heated by gas layers only.")
        '                        Else
        '                            PrintLine(1, TAB(20), "Glass is heated by gas layer and flame radiation.")
        '                            PrintLine(1, TAB(20), "Glass to Flame Distance (m) =", TAB(60), VB6.Format(GLASSdistance(j, i, k), "0.000"))
        '                        End If
        '                    End If
        '                    PrintLine(1)
        '                Next k
        '            End If
        '        End If
        '    Next i
        'Next j

        PrintLine(1, "====================================================================")
        PrintLine(1, "Ceiling/Floor Vents")
        PrintLine(1, "====================================================================")
        For Each oCVent In ocvents
            upperroom = oCVent.upperroom
            lowerroom = oCVent.lowerroom
            PrintLine(1, "Vent ID " & CStr(oCVent.id))
            If upperroom > NumberRooms And lowerroom <= NumberRooms Then
                PrintLine(1, "Upper room " & "outside" & " to lower room " & CStr(lowerroom))
            ElseIf upperroom <= NumberRooms And lowerroom > NumberRooms Then
                PrintLine(1, "Upper room " & CStr(upperroom) & " to lower room " & "outside")
            Else
                PrintLine(1, "Upper room " & CStr(upperroom) & " to lower room " & CStr(lowerroom))
            End If
            PrintLine(1, TAB(20), "Vent Area (m2) =", TAB(60), oCVent.area.ToString)
            PrintLine(1, TAB(20), "Opening Time (sec) =", TAB(60), oCVent.opentime.ToString)
            PrintLine(1, TAB(20), "Closing Time (sec) =", TAB(60), oCVent.closetime.ToString)
            PrintLine(1, TAB(20), "Discharge Coefficient (-) =", TAB(60), oCVent.dischargecoeff.ToString)

            If oCVent.autoopenvent = False Then
                PrintLine(1, TAB(20), "Open method = ", TAB(60), "Manual")
            Else
                PrintLine(1, TAB(20), "Open method = ", TAB(60), "Auto")
                PrintLine(1, TAB(20), "This vent can automatically open with the following trigger(s)") '2016.02
                If oCVent.triggerSD = True Then
                    PrintLine(1, TAB(25), "Smoke Detection in Room " & CStr(oCVent.sdtriggerroom))
                    PrintLine(1, TAB(30), "Vent opens  " & CStr(oCVent.triggerventopendelay) & " sec after detector responds for a duration of " & CStr(oCVent.triggerventopenduration) & " sec.")
                End If
                If oCVent.triggerHD = True Then
                    PrintLine(1, TAB(25), "Sprinkler in Room " & CStr(fireroom))
                    PrintLine(1, TAB(30), "Vent opens  " & CStr(oCVent.triggerventopendelay) & " sec after sprinkler responds for a duration of " & CStr(oCVent.triggerventopenduration) & " sec.")
                End If
                If oCVent.triggerHRR = True Then
                    PrintLine(1, TAB(25), "When fire size reaches " & oCVent.HRRthreshold & " kW + " & oCVent.HRRventopendelay & " sec and remaining open for " & oCVent.HRRventopenduration & " sec.")
                End If
                If oCVent.triggerFO = True Then
                    PrintLine(1, TAB(25), "At flashover")
                End If
                If oCVent.triggerVL = True Then
                    PrintLine(1, TAB(25), "At ventilation limit")
                End If

            End If
            PrintLine(1)

        Next

        PrintLine(1, "====================================================================")
        PrintLine(1, "Ambient Conditions")
        PrintLine(1, "====================================================================")
        PrintLine(1, "Interior Temp (C) =", TAB(60), VB6.Format(InteriorTemp - 273, "0.0"))
        PrintLine(1, "Exterior Temp (C) =", TAB(60), VB6.Format(ExteriorTemp - 273, "0.0"))
        PrintLine(1, "Relative Humidity (%) =", TAB(60), VB6.Format(RelativeHumidity * 100, "0"))
        PrintLine(1)
        PrintLine(1, "====================================================================")
        PrintLine(1, "Tenability Parameters")
        PrintLine(1, "====================================================================")
        PrintLine(1, "Monitoring Height for Visibility and FED (m) =", TAB(60), VB6.Format(MonitorHeight, "0.00"))

        If FEDCO = True Then
            PrintLine(1, "Asphyxiant gas model = ", TAB(60), "FED(CO) C/VM2")
        Else
            PrintLine(1, "Asphyxiant gas model = ", TAB(60), "FED(CO/HCN)")
            PrintLine(1, "Occupant Activity Level = ", TAB(60), Activity)
        End If

        If frmOptions1.optIlluminatedSign.Checked = False Then
            PrintLine(1, "Visibility calculations assume: ", TAB(60), "reflective signs")
        Else
            PrintLine(1, "Visibility calculations assume: ", TAB(60), "illuminated signs")
        End If

        PrintLine(1, "Egress path segments for FED calculations")
        PrintLine(1, "1. Start Time (sec)", TAB(60), VB6.Format(FEDPath(0, 0), "0"))
        PrintLine(1, "1. End Time (sec)", TAB(60), VB6.Format(FEDPath(1, 0), "0"))
        PrintLine(1, "1. Room", TAB(60), VB6.Format(FEDPath(2, 0), "0"))
        PrintLine(1, "2. Start Time (sec)", TAB(60), VB6.Format(FEDPath(0, 1), "0"))
        PrintLine(1, "2. End Time (sec)", TAB(60), VB6.Format(FEDPath(1, 1), "0"))
        PrintLine(1, "2. Room", TAB(60), VB6.Format(FEDPath(2, 1), "0"))
        PrintLine(1, "3. Start Time (sec)", TAB(60), VB6.Format(FEDPath(0, 2), "0"))
        PrintLine(1, "3. End Time (sec)", TAB(60), VB6.Format(FEDPath(1, 2), "0"))
        PrintLine(1, "3. Room", TAB(60), VB6.Format(FEDPath(2, 2), "0"))

        PrintLine(1)
        PrintLine(1, "====================================================================")

        PrintLine(1, "Sprinkler / Detector Parameters")
        PrintLine(1, "====================================================================")
        If cjModel = cjJET Then
            PrintLine(1, TAB(10), "Ceiling Jet model used is NIST JET.")
        ElseIf cjModel = cjAlpert Then
            PrintLine(1, TAB(10), "Ceiling Jet model used is Alpert's.")
        End If
        PrintLine(1, TAB(10), "Sprinkler System Reliability", TAB(60), VB6.Format(SprReliability, "0.000"))
        PrintLine(1, TAB(10), "Sprinkler Probability of Suppression", TAB(60), VB6.Format(SprSuppressionProb, "0.000"))
        PrintLine(1, TAB(10), "Sprinkler Cooling Coefficient", TAB(60), VB6.Format(SprCooling, "0.000"))

        For Each oSprinkler In oSprinklers
            PrintLine(1)
            PrintLine(1, TAB(10), "Sprinkler ID", TAB(60), VB6.Format(oSprinkler.sprid, "0"))
            PrintLine(1, TAB(10), "Room", TAB(60), VB6.Format(oSprinkler.room, "0"))

            For Each x In osprdistributions
                If x.id = oSprinkler.sprid Then
                    If x.varname = "rti" Then
                        PrintLine(1, TAB(10), "Response Time Index (m.s)^1/2 =", TAB(60), VB6.Format(x.varvalue, "0"))
                    End If
                    If x.varname = "cfactor" Then
                        PrintLine(1, TAB(10), "Sprinkler C-Factor (m/s)^1/2 = ", TAB(60), VB6.Format(x.varvalue, "0.00"))
                    End If
                    If x.varname = "sprdensity" Then
                        PrintLine(1, TAB(10), "Water Spray Density (mm/min) = ", TAB(60), VB6.Format(x.varvalue, "0.00"))
                    End If
                    If x.varname = "sprr" Then
                        PrintLine(1, TAB(10), "Radial Distance (m) = ", TAB(60), VB6.Format(x.varvalue, "0.000"))
                    End If
                    If x.varname = "sprz" Then
                        PrintLine(1, TAB(10), "Distance below ceiling (m) =", TAB(60), VB6.Format(x.varvalue, "0.000"))
                    End If
                    If x.varname = "acttemp" Then
                        PrintLine(1, TAB(10), "Actuation Temperature (deg C) = ", TAB(60), VB6.Format(x.varvalue, "0.0"))
                    End If
                End If
            Next

        Next

        PrintLine(1)
        PrintLine(1, "====================================================================")

        PrintLine(1, "Smoke Detector Parameters")
        PrintLine(1, "====================================================================")

        PrintLine(1, TAB(10), "Smoke Detection System Reliability", TAB(60), VB6.Format(SDReliability, "0.000"))
        For Each oSmokeDet In oSmokeDetects
            PrintLine(1)
            PrintLine(1, TAB(10), "Smoke Detector ID", TAB(60), VB6.Format(oSmokeDet.sdid, "0"))
            PrintLine(1, TAB(10), "Room", TAB(60), VB6.Format(oSmokeDet.room, "0"))

            For Each x In osddistributions
                If x.id = oSmokeDet.sdid Then
                    If x.varname = "OD" Then
                        PrintLine(1, TAB(10), "Smoke Optical Density for Alarm (1/m)", TAB(60), VB6.Format(x.varvalue, "0.000"))
                    End If
                    If x.varname = "charlength" Then
                        PrintLine(1, TAB(10), "Detector Characteristic Length Number (m) = ", TAB(60), VB6.Format(x.varvalue, "0.00"))
                    End If
                    If x.varname = "sdr" Then
                        PrintLine(1, TAB(10), "Radial Distance from Plume (m) = ", TAB(60), VB6.Format(x.varvalue, "0.00"))
                    End If
                    If x.varname = "sdz" Then
                        PrintLine(1, TAB(10), "Distance below Ceiling (m) = ", TAB(60), VB6.Format(x.varvalue, "0.000"))
                    End If

                End If

            Next
            If oSmokeDet.sdinside = True Then
                Print(1, TAB(10), "Detector response is based on OD inside the detector chamber.")
            Else
                Print(1, TAB(10), "Detector response is based on OD outside the detector chamber.")
            End If
        Next
        PrintLine(1)
        PrintLine(1)
        PrintLine(1, "====================================================================")
        PrintLine(1, "Mechanical Ventilation (to/from outside)")
        PrintLine(1, "====================================================================")
        If ofans.Count = 0 Then PrintLine(1, "Mechanical Ventilation not installed.")
        PrintLine(1, TAB(10), "Mech ventilation system reliability ", TAB(60), VB6.Format(FanReliability, "0.000"))
        PrintLine(1)

        For Each oFan In ofans
            PrintLine(1, TAB(10), "Fan ID = ", TAB(60), VB6.Format(oFan.fanid, "0"))
            PrintLine(1, TAB(10), "Room = ", TAB(60), VB6.Format(oFan.fanroom, "0"))
            PrintLine(1, TAB(10), "Elevation (m)= ", TAB(60), VB6.Format(oFan.fanelevation, "0.000"))
            PrintLine(1, TAB(10), "Flow rate (m3/s)= ", TAB(60), VB6.Format(oFan.fanflowrate, "0.000"))
            PrintLine(1, TAB(10), "Reliability (-)= ", TAB(60), VB6.Format(oFan.fanreliability, "0.000"))

            If oFan.fanmanual = 0 Then
                PrintLine(1, TAB(10), "Manual Start time = ", TAB(60), VB6.Format(oFan.fanstarttime, "0"))
            ElseIf oFan.fanmanual = 1 Then
                PrintLine(1, TAB(10), "Fan starts when local smoke detector operates.")
            Else
                PrintLine(1, TAB(10), "Fan starts when smoke detector in room of fire origin operates.")
            End If
            If oFan.fanextract = True Then
                PrintLine(1, TAB(10), "Fan extracts air from room")
            Else
                PrintLine(1, TAB(10), "Fan supplies air to room")
            End If
            If oFan.fancurve = True Then
                PrintLine(1, TAB(10), "Maximum cross-pressure limit (Pa)= ", TAB(60), VB6.Format(oFan.fanpressurelimit, "0.0"))
            Else
                PrintLine(1, TAB(10), "Fan curve not used.")
            End If
            PrintLine(1)
        Next

        'For j = 1 To NumberRooms
        '    If fanon(j) = False Then
        '        PrintLine(1, "Mechanical Ventilation not installed in Room " & CStr(j))
        '    Else
        '        PrintLine(1, "Mechanical Ventilation installed in Room " & CStr(j))
        '        If UseFanCurve(j) = True Then
        '            PrintLine(1, TAB(10), "Use fan curve")
        '        Else
        '            PrintLine(1, TAB(10), "Do not use fan curve")
        '        End If
        '        PrintLine(1, TAB(10), "Fan Elevation (m) = ", TAB(60), VB6.Format(FanElevation(j), "0.000"))
        '        PrintLine(1, TAB(10), "Start Time (sec) = ", TAB(60), VB6.Format(ExtractStartTime(j), "0"))
        '        PrintLine(1, TAB(10), "Maximum Cross-Fan Pressure Limit (Pa) = ", TAB(60), VB6.Format(MaxPressure(j), "0"))
        '        PrintLine(1, TAB(10), "Number of Fans = ", TAB(60), VB6.Format(NumberFans(j), "0"))
        '        If Extract(j) = True Then
        '            PrintLine(1, TAB(10), "Extract Rate per fan (m3/s) = ", TAB(60), VB6.Format(ExtractRate(j), "0.00"))
        '        Else
        '            PrintLine(1, TAB(10), "Supply Rate per fan (m3/s) = ", TAB(60), VB6.Format(ExtractRate(j), "0.000"))
        '        End If
        '    End If
        'Next j
        PrintLine(1)
        PrintLine(1, "====================================================================")
        PrintLine(1, "Description of the Fire")
        PrintLine(1, "====================================================================")
        'PrintLine(1, "Radiant Loss Fraction =", TAB(60), VB6.Format(RadiantLossFraction, "0.00"))
        'Print #1, "Underventilated Soot Yield Factor ="; Tab(60); Format$(SootFactor, "0.00")
        PrintLine(1, "CO Yield pre-flashover(g/g) =", TAB(60), VB6.Format(frmOptions1.txtpreCO.Text, "0.000"))
        If frmOptions1.optCOman.Checked = True Then
            PrintLine(1, "CO Yield post-flashover(g/g) =", TAB(60), VB6.Format(frmOptions1.txtpostCO.Text, "0.000"))
        End If

        If sootmode = True Then
            ' If frmOptions1.optSootman.Checked = True Then
            PrintLine(1, "Soot Yield pre-flashover(g/g) =", TAB(60), VB6.Format(preSoot, "0.000"))
            PrintLine(1, "Soot Yield post-flashover(g/g) =", TAB(60), VB6.Format(postSoot, "0.000"))
        Else
            PrintLine(1, "Soot Alpha Coefficient =", TAB(60), VB6.Format(SootAlpha, "0.00"))
            PrintLine(1, "Smoke Epsilon Coefficient =", TAB(60), VB6.Format(SootEpsilon, "0.00"))

        End If

        PrintLine(1, "Flame Emission Coefficient (1/m) =", TAB(60), Format(EmissionCoefficient, "0.00"))
        PrintLine(1, "Fuel - Carbon Moles", TAB(60), Format(nC, "0.00"))
        PrintLine(1, "Fuel - Hydrogen Moles", TAB(60), Format(nH, "0.00"))
        PrintLine(1, "Fuel - Oxygen Moles", TAB(60), Format(nO, "0.00"))
        PrintLine(1, "Fuel - Nitrogen Moles", TAB(60), Format(nN, "0.00"))
        PrintLine(1, "Stoichiometric air/fuel ratio", TAB(60), Format(StoichAFratio, "0.0"))

        PrintLine(1)
        If ObjectDescription.GetUpperBound(0) < NumberObjects Then Resize_Objects()
        If autopopulate = True Then
            PrintLine(1, "Burning objects are randomly positioned in room.")
        Else
            PrintLine(1, "Burning objects are manually positioned in room.")
        End If
        If Enhance = False Then
            PrintLine(1, "Enhanced burning submodel is ", TAB(60), "OFF")
        Else
            PrintLine(1, "Enhanced burning submodel is ", TAB(60), "ON")
        End If
        PrintLine(1)

        For j = 1 To NumberObjects
            PrintLine(1, "Burning Object No " & CStr(ObjectItemID(j)))
            PrintLine(1, ObjectDescription(j))
            PrintLine(1, TAB, "Located in Room ", TAB(60), VB6.Format(fireroom, "0"))
            PrintLine(1, TAB, "Energy Yield (kJ/g) =", TAB(60), VB6.Format(EnergyYield(j), "0.0"))
            PrintLine(1, TAB, "CO2 Yield (kg/kg fuel) =", TAB(60), VB6.Format(CO2Yield(j), "0.000"))
            If frmOptions1.chkHCNcalc.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                PrintLine(1, TAB, "HCN Yield (kg/kg fuel) =", TAB(60), VB6.Format(HCNuserYield(j), "0.000"))
            End If
            PrintLine(1, TAB, "H2O Yield (kg/kg fuel) =", TAB(60), VB6.Format(WaterVaporYield(j), "0.000"))

            If frmOptions1.optSootauto.Checked = True Then PrintLine(1, TAB, "Soot Yield (kg/kg fuel) =", TAB(60), VB6.Format(SootYield(j), "0.000"))
            If ObjectMLUA(2, j) > 0 Then
                PrintLine(1, TAB, "Heat Release Rate Per Unit Area (kW/m2) =", TAB(60), VB6.Format(ObjectMLUA(2, j), "0.0"))
            Else
                If ObjectMLUA(0, j) = 0 Then
                    PrintLine(1, TAB, "Mass Loss per unit area (kg/m2) =", TAB(60), VB6.Format(ObjectMLUA(1, j), "0.000"))
                Else
                    PrintLine(1, TAB, "Mass Loss per unit area (kg/m2) =", TAB(60), VB6.Format(ObjectMLUA(0, j), "0.000") & " X Qe + " & VB6.Format(ObjectMLUA(1, j), "0.000"))
                End If
            End If

            PrintLine(1, TAB, "Radiant Loss Fraction =", TAB(60), VB6.Format(ObjectRLF(j), "0.00"))
            PrintLine(1, TAB, "Fire Elevation (m) =", TAB(60), VB6.Format(FireHeight(j), "0.000"))

            If VM2 = False Then

                PrintLine(1, TAB, "Fire Object Length (m) =", TAB(60), VB6.Format(ObjLength(j), "0.000"))
                PrintLine(1, TAB, "Fire Object Width (m) =", TAB(60), VB6.Format(ObjWidth(j), "0.000"))
                PrintLine(1, TAB, "Fire Object Height (m) =", TAB(60), VB6.Format(ObjHeight(j), "0.000"))

            End If

            If autopopulate = True Then
            Else
                PrintLine(1, TAB, "Location, X-coordinate (m) =", TAB(60), VB6.Format(ObjDimX(j), "0.000"))
                PrintLine(1, TAB, "Location, Y-coordinate (m) =", TAB(60), VB6.Format(ObjDimY(j), "0.000"))
            End If

            If FireLocation(j) = 0 Then PrintLine(1, TAB, "Fire Location (for entrainment) =", TAB(60), "CENTRE")
            If FireLocation(j) = 1 Then PrintLine(1, TAB, "Fire Location (for entrainment)=", TAB(60), "WALL")
            If FireLocation(j) = 2 Then PrintLine(1, TAB, "Fire Location (for entrainment)=", TAB(60), "CORNER")

            If ObjectWindEffect(j) = 1 Then
                PrintLine(1, TAB, "Plume behaviour is ", TAB(60), "UNDISTURBED")
            ElseIf ObjectWindEffect(j) = 2 Then
                PrintLine(1, TAB, "Plume behaviour is ", TAB(60), "DISTURBED")
            End If

            If FuelResponseEffects = True Then
                If ObjectPyrolysisOption(j) = 1 Then
                    PrintLine(1, TAB, "MLR from pool fire equation")
                    PrintLine(1, TAB, "Free burn MLR (kg/s/m2) =", TAB(60), Format(ObjectPoolFBMLR(j), "0.000"))
                    PrintLine(1, TAB, "Pool fire pan diameter (m) =", TAB(60), Format(ObjectPoolDiameter(j), "0.000"))
                    PrintLine(1, TAB, "Pool fire fuel density (kg/m3) =", TAB(60), Format(ObjectPoolDensity(j), "0.0"))
                    PrintLine(1, TAB, "Pool fire fuel quantity (L) =", TAB(60), Format(ObjectPoolVolume(j), "0.00"))
                    PrintLine(1, TAB, "Pool fire ramp up time (s) =", TAB(60), Format(ObjectPoolRamp(j), "0"))
                    PrintLine(1, TAB, "Pool fire fuel vaporisation temp (C) =", TAB(60), Format(ObjectPoolVapTemp(j), "0"))

                ElseIf ObjectPyrolysisOption(j) = 2 Then
                    PrintLine(1, TAB, "MLR from wood crib equations")
                ElseIf ObjectPyrolysisOption(j) = 0 Then
                    PrintLine(1, TAB, "MLR is free burn data")
                End If
            End If

            PrintLine(1)
            If usepowerlawdesignfire = True Then
                If useT2fire = True Then
                    PrintLine(1, TAB, "Alpha T2 growth coefficient =", TAB(60), VB6.Format(AlphaT, "0.0000"))
                Else
                    PrintLine(1, TAB, "Alpha T3 growth coefficient =", TAB(60), VB6.Format(AlphaT, "0.00000"))
                    PrintLine(1, TAB, "Storage Height (m) =", TAB(60), VB6.Format(StoreHeight, "0.00"))
                End If
                PrintLine(1, TAB, "Peak HRR (kW) =", TAB(60), VB6.Format(PeakHRR, "0"))

            Else

                If FuelResponseEffects = True Then
                    If ObjectPyrolysisOption(j) = 1 Then

                    ElseIf ObjectPyrolysisOption(j) = 2 Then

                    ElseIf ObjectPyrolysisOption(j) = 0 Then
                        PrintLine(1, TAB, "Time (sec)", TAB(40), "Free burn mass loss rate (kg/s)")
                        For i = 1 To NumberDataPoints(j)
                            If i < 501 Then PrintLine(1, TAB, MLRData(1, i, j), TAB(40), Format(MLRData(2, i, j), "0.000"))
                        Next i
                    End If
                Else
                    PrintLine(1, TAB, "Time (sec)", TAB(40), "Heat Release (kW)")
                    For i = 1 To NumberDataPoints(j)
                        If i < 501 Then PrintLine(1, TAB, HeatReleaseData(1, i, j), TAB(40), Format(HeatReleaseData(2, i, j), "0"))
                    Next i
                End If
            End If

            PrintLine(1)
        Next j

        PrintLine(1, "====================================================================")
        PrintLine(1, "Postflashover Inputs")
        PrintLine(1, "====================================================================")
        If frmOptions1.optPostFlashover.Checked = False Then
            'If g_post = False Then
            PrintLine(1, "Postflashover model is OFF.")
        Else
            PrintLine(1, "Postflashover model is ON.")
            PrintLine(1, TAB(10), "FLED (MJ/m2) = ", TAB(60), VB6.Format(FLED, "0"))

            PrintLine(1, TAB(10), "Average heat of Combustion (MJ/kg) = ", TAB(60), VB6.Format(HoC_fuel, "0.0"))
            PrintLine(1, TAB(10), "Wood Crib Stick Thickness (m) = ", TAB(60), VB6.Format(Fuel_Thickness, "0.000"))
            PrintLine(1, TAB(10), "Wood Crib Stick Spacing (m) = ", TAB(60), VB6.Format(Stick_Spacing, "0.000"))
            PrintLine(1, TAB(10), "Wood Crib Height(m) = ", TAB(60), VB6.Format(Cribheight, "0.000"))
            PrintLine(1, TAB(10), "Excess Fuel Factor (-) = ", TAB(60), VB6.Format(ExcessFuelFactor, "0.00"))

        End If
        PrintLine(1)
        If RCNone = False Then
            ' If frmOptions1.optRCNone.Checked = False Then
            PrintLine(1, "====================================================================")
            PrintLine(1, "Flame Spread Inputs")
            PrintLine(1, "====================================================================")
            PrintLine(1, "This simulation includes flame spread on linings.")
            If IgnCorrelation = vbJanssens Then PrintLine(1, "Ignition data is correlated using the method of Grenier and Janssens.")
            If IgnCorrelation = vbFTP Then PrintLine(1, "Cone Calorimeter Ignition data is correlated using the Flux Time Product method.")
            If frmOptions1.optQuintiere.Checked = True Then
                PrintLine(1, "Quintiere's Room Corner Model is used.")
                PrintLine(1, "Flame length power =", TAB(60), VB6.Format(FlameLengthPower, "0.000"))
                'Print #1, "Burner flame heat flux (kW/m2) ="; Tab(60); Format(QFB, "0.0")
                'Print #1, "Heat flux ahead of flame (kW/m2) ="; Tab(60); Format(QSpread, "0.0")
                PrintLine(1, "Flame area constant =", TAB(60), VB6.Format(FlameAreaConstant, "0.0000"))
                PrintLine(1, "Burner Width (m) =", TAB(60), VB6.Format(BurnerWidth, "0.000"))
                For j = 1 To NumberRooms
                    PrintLine(1)
                    PrintLine(1, "Room " & CStr(j))
                    If WallConeDataFile(j) <> "null.txt" Then
                        PrintLine(1, "Wall Lining")
                        PrintLine(1, TAB, WallSurface(j))
                        PrintLine(1, TAB, "Cone HRR data file used =", TAB(60), VB6.Format(WallConeDataFile(j)))
                        PrintLine(1, TAB, "Heat of combustion (kJ/g) =", TAB(60), VB6.Format(WallEffectiveHeatofCombustion(j), "0.0"))
                        PrintLine(1, TAB, "Soot/smoke yield (g/g)=", TAB(60), VB6.Format(WallSootYield(j), "0.000"))
                        PrintLine(1, TAB, "CO2 yield (g/g)=", TAB(60), VB6.Format(WallCO2Yield(j), "0.000"))
                        PrintLine(1, TAB, "H20 yield (g/g)=", TAB(60), VB6.Format(WallH2OYield(j), "0.000"))
                        PrintLine(1, TAB, "Min surface temp for spread (C) =", TAB(60), VB6.Format(WallTSMin(j) - 273, "0.0"))
                        PrintLine(1, TAB, "Lateral flame spread parameter =", TAB(60), VB6.Format(WallFlameSpreadParameter(j), "0.0"))
                        PrintLine(1, TAB, "Ignition temperature (C) =", TAB(60), VB6.Format(IgTempW(j) - 273, "0.0"))
                        PrintLine(1, TAB, "Thermal inertia =", TAB(60), VB6.Format(ThermalInertiaWall(j), "0.000"))
                        PrintLine(1, TAB, "Critical Flux for Ignition (kW/m2)", TAB(60), VB6.Format(WallQCrit(j), "0.0"))
                        If IgnCorrelation = vbFTP And WallFTP(j) > 0 Then
                            PrintLine(1, TAB, "Ignition Correlation Power, n")
                            PrintLine(1, TAB, "(1=thermally thin; 0.5=thermally thick)", TAB(60), VB6.Format(1 / Walln(j), "0.00"))
                            'PrintLine(1, TAB, "Ignition Correlation Power, n", TAB(60), VB6.Format(Walln(j), "0.0"))

                            PrintLine(1, TAB, "Flux Time Product (FTP)", TAB(60), VB6.Format(WallFTP(j), "0"))
                        End If
                        If frmQuintiere.optUseOneCurve.Checked = True Then
                            PrintLine(1, TAB, "Total Energy Available (kJ/m2) =", TAB(60), VB6.Format(AreaUnderWallCurve(j), "0"))
                            PrintLine(1, TAB, "Heat of Gasification (kJ/g) =", TAB(60), VB6.Format(WallHeatofGasification(j), "0.0"))
                        End If
                    End If
                    If CeilingConeDataFile(j) <> "null.txt" Then
                        PrintLine(1, "Ceiling Lining")
                        PrintLine(1, TAB, CeilingSurface(j))
                        PrintLine(1, TAB, "Cone HRR data file used =", TAB(60), VB6.Format(CeilingConeDataFile(j)))
                        PrintLine(1, TAB, "Heat of combustion (kJ/g) =", TAB(60), VB6.Format(CeilingEffectiveHeatofCombustion(j), "0.0"))
                        PrintLine(1, TAB, "Soot/smoke yield (g/g)=", TAB(60), VB6.Format(CeilingSootYield(j), "0.000"))
                        PrintLine(1, TAB, "CO2 yield (g/g)=", TAB(60), VB6.Format(CeilingCO2Yield(j), "0.000"))
                        PrintLine(1, TAB, "H20 yield (g/g)=", TAB(60), VB6.Format(CeilingH2OYield(j), "0.000"))
                        PrintLine(1, TAB, "Ignition temperature (C) =", TAB(60), VB6.Format(IgTempC(j) - 273, "0.0"))
                        PrintLine(1, TAB, "Thermal inertia =", TAB(60), VB6.Format(ThermalInertiaCeiling(j), "0.000"))
                        PrintLine(1, TAB, "Critical Flux for Ignition (kW/m2)", TAB(60), VB6.Format(CeilingQCrit(j), "0.0"))
                        If IgnCorrelation = vbFTP And CeilingFTP(j) > 0 Then
                            PrintLine(1, TAB, "Ignition Correlation Power, n")
                            PrintLine(1, TAB, "(1=thermally thin; 0.5=thermally thick)", TAB(60), VB6.Format(1 / Ceilingn(j), "0.00"))
                            PrintLine(1, TAB, "Flux Time Product (FTP)", TAB(60), VB6.Format(CeilingFTP(j), "0"))
                        End If
                        If frmQuintiere.optUseOneCurve.Checked = True Then
                            PrintLine(1, TAB, "Total Energy Available (kJ/m2) =", TAB(60), VB6.Format(AreaUnderCeilingCurve(j), "0"))
                            PrintLine(1, TAB, "Heat of Gasification (kJ/g) =", TAB(60), VB6.Format(CeilingHeatofGasification(j), "0.0"))
                        End If
                    End If
                    If FloorConeDataFile(j) <> "null.txt" Then
                        PrintLine(1, "Floor Lining")
                        PrintLine(1, TAB, FloorSurface(j))
                        PrintLine(1, TAB, "Cone HRR data file used =", TAB(60), VB6.Format(FloorConeDataFile(j)))
                        PrintLine(1, TAB, "Heat of combustion (kJ/g) =", TAB(60), VB6.Format(FloorEffectiveHeatofCombustion(j), "0.0"))
                        PrintLine(1, TAB, "Soot/smoke yield (g/g)=", TAB(60), VB6.Format(FloorSootYield(j), "0.000"))
                        PrintLine(1, TAB, "CO2 yield (g/g)=", TAB(60), VB6.Format(FloorCO2Yield(j), "0.000"))
                        PrintLine(1, TAB, "H20 yield (g/g)=", TAB(60), VB6.Format(FloorH2OYield(j), "0.000"))
                        PrintLine(1, TAB, "Min surface temp for spread (C) =", TAB(60), VB6.Format(FloorTSMin(j) - 273, "0.0"))
                        PrintLine(1, TAB, "Lateral flame spread parameter =", TAB(60), VB6.Format(FloorFlameSpreadParameter(j), "0.0"))
                        PrintLine(1, TAB, "Ignition temperature (C) =", TAB(60), VB6.Format(IgTempF(j) - 273, "0.0"))
                        PrintLine(1, TAB, "Thermal inertia =", TAB(60), VB6.Format(ThermalInertiaFloor(j), "0.000"))
                        If IgnCorrelation = vbFTP And FloorFTP(j) > 0 Then
                            PrintLine(1, TAB, "Critical Flux for Ignition (kW/m2)", TAB(60), VB6.Format(FloorQCrit(j), "0.0"))
                            PrintLine(1, TAB, "Ignition Correlation Power")
                            PrintLine(1, TAB, "(1=thermally thin; 0.5=thermally thick)", TAB(60), VB6.Format(1 / Floorn(j), "0.00"))
                            PrintLine(1, TAB, "Flux Time Product (FTP)", TAB(60), VB6.Format(FloorFTP(j), "0"))
                        End If
                        If frmQuintiere.optUseOneCurve.Checked = True Then
                            PrintLine(1, TAB, "Total Energy Available (kJ/m2) =", TAB(60), VB6.Format(AreaUnderFloorCurve(j), "0"))
                            PrintLine(1, TAB, "Heat of Gasification (kJ/g) =", TAB(60), VB6.Format(FloorHeatofGasification(j), "0.0"))
                        End If
                    End If
                Next j
            ElseIf frmOptions1.optKarlsson.Checked = True Then
                PrintLine(1, "Karlsson's model A concurrent flow flame spread Model is used.")
                PrintLine(1, "Heat flux to wall (kW) =", TAB(60), VB6.Format(WallHeatFlux, "0.0"))
                PrintLine(1, "Heat flux to ceiling (kW) =", TAB(60), VB6.Format(CeilingHeatFlux, "0.0"))
                PrintLine(1, "Flame area constant =", TAB(60), VB6.Format(FlameAreaConstant, "0.000"))
                PrintLine(1, "Burner Width (m) =", TAB(60), VB6.Format(BurnerWidth, "0.000"))
                PrintLine(1, "Ignition temperature of wall lining (C) =", TAB(60), VB6.Format(IgTempW(fireroom) - 273, "0.0"))
                PrintLine(1, "Ignition temperature of ceiling lining (C) =", TAB(60), VB6.Format(IgTempC(fireroom) - 273, "0.0"))
                PrintLine(1, "Ignition temperature of floor covering (C) =", TAB(60), VB6.Format(IgTempF(fireroom) - 273, "0.0"))
                PrintLine(1, "Thermal inertia of wall lining =", TAB(60), VB6.Format(ThermalInertiaWall(fireroom), "0.000"))
                PrintLine(1, "Thermal inertia of ceiling lining =", TAB(60), VB6.Format(ThermalInertiaCeiling(fireroom), "0.000"))
                PrintLine(1, "Thermal inertia of floor covering =", TAB(60), VB6.Format(ThermalInertiaFloor(fireroom), "0.000"))
                PrintLine(1, "Wall cone HRR data file used =", TAB(60), VB6.Format(WallConeDataFile(fireroom)))
                PrintLine(1, "Ceiling cone HRR data file used =", TAB(60), VB6.Format(CeilingConeDataFile(fireroom)))
                PrintLine(1, "Floor cone HRR data file used =", TAB(60), VB6.Format(FloorConeDataFile(fireroom)))
                PrintLine(1, "Wall peak heat release rate (kW/sqm) =", TAB(60), VB6.Format(PeakWallHRR(1, fireroom), "0.0"))
                PrintLine(1, "Ceiling peak heat release rate (kW/sqm) =", TAB(60), VB6.Format(PeakCeilingHRR(1, fireroom), "0.0"))
                PrintLine(1, "Floor peak heat release rate (kW/sqm) =", TAB(60), VB6.Format(PeakFloorHRR(1, fireroom), "0.0"))
            End If
        End If

        PrintLine(1)
        FileClose(1)
        Exit Sub

errorhandler:
        'Err.Clear
        If Err.Number = 55 Then
            FileClose(1)
            Resume
        End If
        Exit Sub
    End Sub


    Public Sub Interpolate2(ByRef X(,) As Double, ByRef Y() As Double, ByVal n As Integer, ByVal Xint As Double, ByRef yinterp As Double)
        '****************************************************************************
        '*  From ProMath 2.0
        '*  This Subroutine Performs Linear Interpolation Within A Set Of X(),Y()
        '*  Pairs To Give The Y Value Corresponding To Xint.
        '*       X       Array Of Values Of The Independent Variable (a 2D array)
        '*       Y       Array Of Values Of The Dependent Variable Corresponding To X
        '*       N       Number Of Points To Interpolate.
        '*       Xint    The X-value For Which Estimate Of Y Is Desired.
        '*       Yinterp    The Y-value Returned To The User.
        '*
        '*  Revised Colleen Wade 12 October 1996
        '****************************************************************************

        Dim i As Short
        Dim y1, x1, x2, y2 As Double

        yinterp = 0.0#

        For i = 1 To n - 1
            If X(i, 1) <= Xint And X(i + 1, 1) >= Xint Then
                x1 = X(i, 1)
                x2 = X(i + 1, 1)
                y1 = Y(i)
                y2 = Y(i + 1)
                yinterp = (Xint - x1) * (y2 - y1) / (x2 - x1) + y1
                Exit Sub
            End If
        Next i
    End Sub

    '
    Public Sub FED_radiation(ByVal room As Integer)
        '* ======================================================
        '*  This procedure calculates the Fractional Effective
        '*  Dose causing incapacitation due to radiant heat
        '*  exposure effects.
        '*
        '* ======================================================

        Dim i As Integer
        Dim radSum, tburn As Single
        Dim rad As Double
        Dim TargetFlag As Short
        Dim emissivity As Double

        On Error GoTo toxicityhandler

        'ReDim FEDRadSum(1 To NumberRooms, NumberTimeSteps + 2) As Double

        radSum = 0

        For i = 1 To NumberTimeSteps + 1
            rad = 0
            tburn = 0
            'If i = 244 Then Stop
            Call Radiation_to_Surface(room, layerheight(room, i), uppertemp(room, i), CO2MassFraction(room, i, 1), H2OMassFraction(room, i, 1), rad, UpperVolume(room, i), OD_upper(room, i), MonitorHeight, emissivity)

            'If room = fireroom Then Debug.Print i, CO2MassFraction(room, i, 1), H2OMassFraction(room, i, 1), rad, radSum
            If tim(i, 1) >= StartOccupied And tim(i, 1) <= EndOccupied Then
                'rad = 0
                'tburn = 0
                'Call Radiation_to_Surface(ByVal room, ByVal layerheight(room, i), ByVal uppertemp(room, i), ByVal CO2MassFraction(room, i, 1), ByVal H2OMassFraction(room, i, 1), rad, UpperVolume(room, i))
                If rad > 1.7 Then tburn = Timestep / (55 * (rad - 1.7) ^ (-0.8))
                radSum = radSum + tburn
            End If

            If radSum < 1 Then
                FEDRadSum(room, i) = radSum
            Else
                FEDRadSum(room, i) = 1
            End If

            SurfaceRad(room, i) = rad

            If i = 1 Then TargetFlag = 0

            'fixed 27/2/07
            If TargetEndPoint <= FEDRadSum(fireroom, i) And TargetFlag = 0 And room = fireroom Then
                'If TargetEndPoint <= FEDRadSum(fireroom, i) And TargetEndPoint = 0 And room = fireroom Then
                frmEndPoints.lblTarget.Text = "FED Rad (incap) Exceeded " & VB6.Format(TargetEndPoint) & " at " & VB6.Format(tim(i, 1), "0.0") & " Seconds."
                TargetFlag = 1
            End If

        Next i

        Exit Sub

toxicityhandler:
        Exit Sub

    End Sub

    Public Sub FED_thermal_iso13571(ByVal room As Integer)
        '* ======================================================
        '*  This procedure calculates the Fractional Effective
        '*  Dose causing incapacitation due to radiant and convective heat
        '*  exposure effects.
        '*
        '* ======================================================

        Dim i As Integer
        Dim tburn, radSum, T As Single
        Dim rad As Double
        Dim TargetFlag As Short
        Dim emissivity As Double

        On Error GoTo toxicityhandler

        radSum = 0

        For i = 1 To NumberTimeSteps + 1
            rad = 0
            tburn = 0

            Call Radiation_to_Surface(room, layerheight(room, i), uppertemp(room, i), CO2MassFraction(room, i, 1), H2OMassFraction(room, i, 1), rad, UpperVolume(room, i), OD_upper(room, i), MonitorHeight, emissivity)
            upperemissivity(room, i) = emissivity

            If tim(i, 1) >= StartOccupied And tim(i, 1) <= EndOccupied Then

                If rad > 2.5 Then
                    tburn = Timestep / (60 * 6.9 * (rad) ^ (-1.56)) 'time to 2nd degree burns in sec, eqn 7a ISO/DIS 13571
                End If

                'T gas temperature at monitoring height
                If MonitorHeight > layerheight(room, i) Then
                    T = uppertemp(room, i) - 273
                Else
                    T = lowertemp(room, i) - 273
                End If

                tburn = tburn + Timestep / (60 * 50000000.0# * T ^ (-3.4)) 'convected heat for unclothed or lightly clothed subjects, eqn 9 ISO/DIS 13571

                radSum = radSum + tburn
            End If

            If radSum < 1 Then
                FEDRadSum(room, i) = radSum
            Else
                FEDRadSum(room, i) = 1
            End If

            SurfaceRad(room, i) = rad

            If i = 1 Then TargetFlag = 0

            'fixed 27/2/07
            If TargetEndPoint <= FEDRadSum(fireroom, i) And TargetFlag = 0 And room = fireroom Then
                frmEndPoints.lblTarget.Text = "FED thermal (incap) exceeded " & VB6.Format(TargetEndPoint) & " at " & VB6.Format(tim(i, 1), "0.0") & " Seconds."
                TargetFlag = 1
            End If

        Next i

        Exit Sub

toxicityhandler:
        Exit Sub

    End Sub
    Public Sub FED_gases_multi()

        'for the general case, allowing for CO, O2 and HCN
        'plus CO2 hyperventilation effects on CO and HCN uptake
        'allows activity level to be accounted for
        'switches between upper and lower layer as requiring for the specified monitoring height
        'switches between rooms, depending on the egress path / times specifies (max 3 path segments)
        'follows SFPE HB 4th ed

        Try
            If IsNothing(tim) Then Exit Sub
            If NumberTimeSteps < 1 Then Exit Sub

            Dim i, FEDFlag As Short
            Dim O2Sum As Single
            Dim COSum As Double
            Dim RMV2, RMV1, VCO2, COppm As Single
            Dim room As Integer
            Dim HCNsum As Double
            Dim RMVo, COHb As Single

            'Get activity level
            If frmOptions1.optAtRest.Checked = True Then Activity = "Rest"
            If frmOptions1.optLightWork.Checked = True Then Activity = "Light"
            If frmOptions1.optHeavyWork.Checked = True Then Activity = "Heavy"

            'Get values of COHb and RMVo appropriate to activity level
            Select Case Activity
                Case "Rest"
                    COHb = 40 '% COHb dose for incapacitation
                    RMVo = 8.5 'breathing rate l/min
                Case "Light"
                    COHb = 30 '% COHb dose for incapacitation
                    RMVo = 25 'breathing rate l/min
                Case "Heavy"
                    COHb = 20 '% COHb dose for incapacitation
                    RMVo = 50 'breathing rate l/min
            End Select

            O2Sum = 0
            COSum = 0
            HCNsum = 0

            For i = 1 To NumberTimeSteps + 1

                If tim(i, 1) > FEDPath(0, 0) And tim(i, 1) <= FEDPath(1, 0) Then
                    room = FEDPath(2, 0)
                    StartOccupied = FEDPath(0, 0)
                    EndOccupied = FEDPath(1, 0)
                ElseIf tim(i, 1) > FEDPath(0, 1) And tim(i, 1) <= FEDPath(1, 1) Then
                    room = FEDPath(2, 1)
                    StartOccupied = FEDPath(0, 1)
                    EndOccupied = FEDPath(1, 1)
                ElseIf tim(i, 1) > FEDPath(0, 2) And tim(i, 1) <= FEDPath(1, 2) Then
                    room = FEDPath(2, 2)
                    StartOccupied = FEDPath(0, 2)
                    EndOccupied = FEDPath(1, 2)
                End If

                If room = 0 Then Continue For
                If room > NumberRooms Then Exit Sub

                If tim(i, 1) > StartOccupied And tim(i, 1) <= EndOccupied Then
                    If layerheight(room, i) <= MonitorHeight Then
                        'exposed to upper layer
                        If CSng(O2VolumeFraction(room, i, 1)) < 0.13 Then 'ignore O2 effect where % volume > 13%
                            O2Sum = O2Sum + Timestep / 60 / (Exp(8.13 - 0.54 * (20.9 - (O2VolumeFraction(room, i, 1) * 100 + O2VolumeFraction(room, i - 1, 1) * 100) / 2)))
                        End If

                        If CO2VolumeFraction(room, i, 1) > 0.02 Then
                            RMV1 = Exp(0.2 * CO2VolumeFraction(room, i, 1) * 100)
                            RMV2 = Exp(0.2 * CO2VolumeFraction(room, i - 1, 1) * 100)
                        Else
                            RMV1 = 1
                            RMV2 = 1
                        End If
                        VCO2 = (RMV1 + RMV2) / 2

                        If CSng(COVolumeFraction(room, i, 1)) > CSng(COVolumeFraction(room, 1, 1)) Then

                            COppm = ((COVolumeFraction(room, i, 1) + COVolumeFraction(room, i - 1, 1)) / 2) * 10 ^ 6 'CO in ppm
                            COSum = COSum + Timestep / 60 * VCO2 * 0.00003317 * RMVo / COHb * (COppm) ^ 1.036

                        End If

                        'hcn
                        If CSng(HCNVolumeFraction(room, i, 1)) > CSng(HCNVolumeFraction(room, 1, 1)) Then
                            If ((CSng(HCNVolumeFraction(room, i, 1) + HCNVolumeFraction(room, i - 1, 1)) / 2) * 10 ^ 6) > 80 Then
                                HCNsum = HCNsum + Timestep / 60 * VCO2 * Exp(5.396 - 0.023 * (((HCNVolumeFraction(room, i, 1) + HCNVolumeFraction(room, i - 1, 1)) / 2) * 10 ^ 6))
                            End If
                        End If

                    Else
                        'exposed to lower layer
                        If CSng(O2VolumeFraction(room, i, 2)) < 0.13 Then 'ignore O2 effect where % volume > 13%
                            O2Sum = O2Sum + Timestep / 60 / (Exp(8.13 - 0.54 * (20.9 - (O2VolumeFraction(room, i, 2) * 100 + O2VolumeFraction(room, i - 1, 2) * 100) / 2)))
                        End If

                        If CO2VolumeFraction(room, i, 2) > 0.02 Then
                            RMV1 = Exp(0.2 * CO2VolumeFraction(room, i, 2) * 100)
                            RMV2 = Exp(0.2 * CO2VolumeFraction(room, i - 1, 2) * 100)
                        Else
                            RMV1 = 1
                            RMV2 = 1
                        End If
                        VCO2 = (RMV1 + RMV2) / 2

                    End If

                    If CSng(COVolumeFraction(room, i, 2)) > CSng(COVolumeFraction(room, 1, 2)) Then
                        COppm = ((COVolumeFraction(room, i, 2) + COVolumeFraction(room, i - 1, 2)) / 2) * 10 ^ 6 'CO in ppm
                        COSum = COSum + Timestep / 60 * VCO2 * 0.00003317 * RMVo / COHb * (COppm) ^ 1.036
                    End If

                    'hcn
                    If CSng(HCNVolumeFraction(room, i, 2)) > CSng(HCNVolumeFraction(room, 1, 2)) Then
                        If ((CSng(HCNVolumeFraction(room, i, 2) + HCNVolumeFraction(room, i - 1, 2)) / 2) * 10 ^ 6) > 80 Then
                            HCNsum = HCNsum + Timestep / 60 * VCO2 * Exp(5.396 - 0.023 * (((HCNVolumeFraction(room, i, 2) + HCNVolumeFraction(room, i - 1, 2)) / 2) * 10 ^ 6))
                        End If
                    End If

                End If

                If O2Sum + COSum + HCNsum < 1 Then
                    FEDSum(fireroom, i) = O2Sum + COSum + HCNsum
                Else
                    FEDSum(fireroom, i) = 1
                End If

                If i = 1 Then FEDFlag = 0

                If FEDEndPoint <= FEDSum(fireroom, i) And FEDFlag = 0 Then
                    Dim Message As String = "FED(gases) Exceeded " & VB6.Format(FEDEndPoint) & " at " & VB6.Format(tim(i, 1), "0.0") & " Seconds."
                    frmEndPoints.lblFED.Text = Message
                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                    FEDFlag = 1
                End If

            Next i

            If NumberTimeSteps < 1 Then Exit Sub

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in FED_gases_multi")
        End Try

    End Sub
    Public Sub FED_CO_iso13571_multi()

        'as per C/VM2 and ISO 13571
        Try

            If IsNothing(tim) Then Exit Sub
            If NumberTimeSteps < 1 Then Exit Sub

            Dim i, FEDFlag As Integer
            Dim O2Sum As Single
            Dim COSum As Double
            Dim RMV2, RMV1, VCO2 As Single
            Dim room As Integer

            O2Sum = 0
            COSum = 0

            For i = 1 To NumberTimeSteps + 1
                If i > tim.GetUpperBound(0) Then Exit For

                If tim(i, 1) > FEDPath(0, 0) And tim(i, 1) <= FEDPath(1, 0) Then
                    room = FEDPath(2, 0)
                    StartOccupied = FEDPath(0, 0)
                    EndOccupied = FEDPath(1, 0)
                ElseIf tim(i, 1) > FEDPath(0, 1) And tim(i, 1) <= FEDPath(1, 1) Then
                    room = FEDPath(2, 1)
                    StartOccupied = FEDPath(0, 1)
                    EndOccupied = FEDPath(1, 1)
                ElseIf tim(i, 1) > FEDPath(0, 2) And tim(i, 1) <= FEDPath(1, 2) Then
                    room = FEDPath(2, 2)
                    StartOccupied = FEDPath(0, 2)
                    EndOccupied = FEDPath(1, 2)
                End If

                If room = 0 Then Continue For
                If room > NumberRooms Then Exit Sub


                If tim(i, 1) > StartOccupied And tim(i, 1) <= EndOccupied Then
                    If layerheight(room, i) <= MonitorHeight Then
                        'exposed to upper layer
                        If CSng(O2VolumeFraction(room, i, 1)) < 0.13 Then 'ignore O2 effect where % volume > 13%
                            O2Sum = O2Sum + Timestep / 60 / (Exp(8.13 - 0.54 * (20.9 - (O2VolumeFraction(room, i, 1) * 100 + O2VolumeFraction(room, i - 1, 1) * 100) / 2)))
                        End If

                        If CO2VolumeFraction(room, i, 1) > 0 Then 'include hyperventilation for all CO2 as per ISO 13571:2012
                            ' VCO2 = exp(%CO2 / 5)
                            RMV1 = Exp(0.2 * CO2VolumeFraction(room, i, 1) * 100)
                            RMV2 = Exp(0.2 * CO2VolumeFraction(room, i - 1, 1) * 100)
                        Else
                            RMV1 = 1
                            RMV2 = 1
                        End If
                        VCO2 = (RMV1 + RMV2) / 2

                        If CSng(COVolumeFraction(room, i, 1)) > CSng(COVolumeFraction(room, 1, 1)) Then
                            COSum = COSum + Timestep / (60 * 35000) * VCO2 * (((COVolumeFraction(room, i, 1) + COVolumeFraction(room, i - 1, 1)) / 2) * 10 ^ 6)

                            '(COVolumeFraction(room, i, 1) + COVolumeFraction(room, i - 1, 1)) / 2) * 10 ^ 6
                            ' this is the average CO over the timestep interval converted from volume fraction to ppm
                        End If

                    Else
                        'exposed to lower layer
                        If CSng(O2VolumeFraction(room, i, 2)) < 0.13 Then 'ignore O2 effect where % volume > 13%
                            O2Sum = O2Sum + Timestep / 60 / (Exp(8.13 - 0.54 * (20.9 - (O2VolumeFraction(room, i, 2) * 100 + O2VolumeFraction(room, i - 1, 2) * 100) / 2)))
                        End If

                        If CO2VolumeFraction(room, i, 2) > 0 Then 'include hyperventilation for all CO2 as per ISO 13571:2012
                            RMV1 = Exp(0.2 * CO2VolumeFraction(room, i, 2) * 100)
                            RMV2 = Exp(0.2 * CO2VolumeFraction(room, i - 1, 2) * 100)
                        Else
                            'no hyperventilation effect
                            RMV1 = 1
                            RMV2 = 1
                        End If
                        VCO2 = (RMV1 + RMV2) / 2

                    End If

                    If CSng(COVolumeFraction(room, i, 2)) > CSng(COVolumeFraction(room, 1, 2)) Then
                        COSum = COSum + Timestep / (60 * 35000) * VCO2 * (((COVolumeFraction(room, i, 2) + COVolumeFraction(room, i - 1, 2)) / 2) * 10 ^ 6)
                    End If

                End If

                If O2Sum + COSum < 1 Then
                    FEDSum(fireroom, i) = O2Sum + COSum
                Else
                    FEDSum(fireroom, i) = 1
                End If

                If i = 1 Then FEDFlag = 0

                If FEDEndPoint <= FEDSum(fireroom, i) And FEDFlag = 0 Then
                    Dim Message As String = "FED(CO) Exceeded " & VB6.Format(FEDEndPoint) & " at " & VB6.Format(tim(i, 1), "0.0") & " Seconds."
                    frmEndPoints.lblFED.Text = Message
                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                    FEDFlag = 1
                End If

            Next i


        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in FED_CO_iso13571_multi")
        End Try

    End Sub
    Public Sub FED_thermal_iso13571_multi()
        '* ======================================================
        '*  This procedure calculates the Fractional Effective
        '*  Dose causing incapacitation due to radiant and convective heat
        '*  exposure effects.
        '*  threshold temperature below which the convective part of FED does not accumulate has been added 23/5/16
        '* ======================================================

        Try
            If IsNothing(tim) Then Exit Sub

            Dim i, room As Integer
            Dim tburn, radSum, T As Single
            Dim rad As Double
            Dim TargetFlag As Short
            Dim emissivity As Double

            radSum = 0

            For i = 1 To NumberTimeSteps + 1
                rad = 0
                tburn = 0
                If i > tim.GetUpperBound(0) Then Exit For
                If tim(i, 1) > FEDPath(0, 0) And tim(i, 1) <= FEDPath(1, 0) Then
                    room = FEDPath(2, 0)
                    StartOccupied = FEDPath(0, 0)
                    EndOccupied = FEDPath(1, 0)
                ElseIf tim(i, 1) > FEDPath(0, 1) And tim(i, 1) <= FEDPath(1, 1) Then
                    room = FEDPath(2, 1)
                    StartOccupied = FEDPath(0, 1)
                    EndOccupied = FEDPath(1, 1)
                ElseIf tim(i, 1) > FEDPath(0, 2) And tim(i, 1) <= FEDPath(1, 2) Then
                    room = FEDPath(2, 2)
                    StartOccupied = FEDPath(0, 2)
                    EndOccupied = FEDPath(1, 2)
                End If

                For jroom = 1 To NumberRooms
                    Call Radiation_to_Surface(jroom, layerheight(jroom, i), uppertemp(jroom, i), CO2MassFraction(jroom, i, 1), H2OMassFraction(jroom, i, 1), rad, UpperVolume(jroom, i), OD_upper(jroom, i), MonitorHeight, emissivity)
                    upperemissivity(jroom, i) = emissivity
                    SurfaceRad(jroom, i) = rad
                Next

                If room = 0 Then Continue For

                If room > NumberRooms Then Exit Sub

                If tim(i, 1) > StartOccupied And tim(i, 1) <= EndOccupied Then

                    If SurfaceRad(room, i) > 2.5 Then
                        tburn = Timestep / (60 * 6.9 * ((SurfaceRad(room, i) + SurfaceRad(room, i - 1)) / 2) ^ (-1.56)) 'time to 2nd degree burns in sec, eqn 7a ISO/DIS 13571
                    End If

                    'T gas temperature at monitoring height
                    If MonitorHeight > layerheight(room, i) Then
                        T = (uppertemp(room, i) + uppertemp(room, i - 1)) / 2 - 273 'gas temp in deg C
                    Else
                        T = (lowertemp(room, i) + lowertemp(room, i - 1)) / 2 - 273
                    End If

                    If T > 25 Then 'eqn only valid for T>25, if T<25 then ignore convective term
                        tburn = tburn + Timestep / (60 * 50000000.0# * T ^ (-3.4)) 'convected heat for unclothed or lightly clothed subjects, eqn 9 ISO/DIS 13571
                    End If

                    radSum = radSum + tburn
                End If


                If radSum < 1 Then
                    FEDRadSum(fireroom, i) = radSum
                Else
                    FEDRadSum(fireroom, i) = 1
                End If

                If i = 1 Then TargetFlag = 0

                'fixed 27/2/07
                If TargetEndPoint <= FEDRadSum(fireroom, i) And TargetFlag = 0 Then
                    Dim Message As String = "FED(thermal) Exceeded " & VB6.Format(TargetEndPoint) & " at " & VB6.Format(tim(i, 1), "0.0") & " Seconds."
                    frmEndPoints.lblTarget.Text = Message

                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                    TargetFlag = 1
                End If

            Next i

            Exit Sub

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in FED_thermal_iso13571_multi")
        End Try

    End Sub

    Public Sub Open_Vector_File(ByRef opendatafile As Object, ByVal thisfile As Object, ByVal totalfiles As Object)
        '*  ===================================================================
        '*  Open a cone vector file.
        '*  ===================================================================
        'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim sea, str_Renamed, data, hoc, conefile As String
        Dim temp, qmax As Single
        Dim i, Action As Short
        Dim j As Integer
        Dim uniquefiles As Short
        Dim X() As Double
        Dim Y() As Double
        Dim yinterp As Double
        Dim igtime() As Double
        Dim S_filename As String
        ReDim Preserve imax(totalfiles)
        ReDim Preserve ig_step(totalfiles)
        ReDim Preserve data_keep(2, 2000, totalfiles)
        ReDim Preserve ExtFlux(totalfiles)
        ReDim Preserve ConePeak(totalfiles)
        ReDim Preserve ConeHoC(totalfiles)
        ReDim Preserve ConeSEA(totalfiles)

        'UPGRADE_WARNING: Couldn't resolve default property of object thisfile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If thisfile = 1 Then
            ReDim imax(totalfiles)
            ReDim ig_step(totalfiles)
            ReDim data_keep(2, 2000, totalfiles)
            ReDim ExtFlux(totalfiles)
            ReDim ConeHoC(totalfiles)
            ReDim ConeSEA(totalfiles)
            ReDim ConePeak(totalfiles)
        End If

        'if file exists, open it
        'UPGRADE_WARNING: Couldn't resolve default property of object opendatafile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If opendatafile <> "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object opendatafile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Confirm_File(opendatafile, readfile, 0) = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object opendatafile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                FileOpen(1, opendatafile, OpenMode.Input)
                On Error GoTo ConfirmFileError2
10:
                'str = InputBox("What is the external heat flux for file " + Dir(opendatafile) + "?", "Input Heat Flux", 35)
                'If IsNumeric(str) = True Then
                '    ExtFlux(thisfile) = CLng(str)
                'Else
                '    If str = "" Then Err = cdlCancel: GoTo errhandler
                '    Err = 13
                '    GoTo errhandler
                'End If

                i = 1
                Do While EOF(1) = False
                    data = LineInput(1)
                    If StrConv(data, VbStrConv.Uppercase) = "FLUX" Then
                        str_Renamed = LineInput(1)
                        If IsNumeric(str_Renamed) = True Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object thisfile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            ExtFlux(thisfile) = CInt(str_Renamed)
                        Else
                            If str_Renamed = "" Then
                                'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
                                Err.Number = DialogResult.Cancel : GoTo errhandler
                            End If
                            Err.Number = 13
                            GoTo errhandler
                        End If
                    End If

                    If StrConv(data, VbStrConv.Uppercase) = "AVGHC" Then
                        hoc = LineInput(1)
                        If IsNumeric(hoc) = True Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object thisfile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            ConeHoC(thisfile) = CSng(hoc)
                        Else
                            If hoc = "" Then
                                'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
                                Err.Number = DialogResult.Cancel : GoTo errhandler
                            End If
                            Err.Number = 13
                            GoTo errhandler
                        End If
                    End If

                    If StrConv(data, VbStrConv.Uppercase) = "AVSIGMA" Then
                        sea = LineInput(1)
                        If IsNumeric(sea) = True Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object thisfile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            ConeSEA(thisfile) = Abs(CSng(sea))
                        Else
                            If sea = "" Then
                                'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
                                Err.Number = DialogResult.Cancel : GoTo errhandler
                            End If
                            Err.Number = 13
                            GoTo errhandler
                        End If
                    End If

                    If StrConv(data, VbStrConv.Uppercase) = "TIME" Then
                        'throw next 3 lines away
                        data = LineInput(1)
                        data = LineInput(1)
                        data = LineInput(1)
                        'keep the numbers
                        Do Until (data = "VARIABLE") Or (EOF(1) = True)
                            data = LineInput(1)
                            On Error Resume Next
                            'UPGRADE_WARNING: Couldn't resolve default property of object thisfile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If IsNumeric(data) = True Then data_keep(1, i, thisfile) = CSng(data)
                            i = i + 1
                            If i > 2000 Then Exit Do
                        Loop
                        Err.Clear()
                        'UPGRADE_WARNING: Couldn't resolve default property of object thisfile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        imax(thisfile) = i - 2
                    End If

                    i = 1
                    qmax = 0
                    If StrConv(data, VbStrConv.Uppercase) = "R.H.R." Or StrConv(data, VbStrConv.Uppercase) = "HRR/A" Or StrConv(data, VbStrConv.Uppercase) = "HRR" Then
                        'throw next 2 lines away
                        data = LineInput(1)
                        data = LineInput(1)
                        'keep the numbers
                        If StrConv(data, VbStrConv.Uppercase) = "W/M^2" Then
                            Do Until data = "VARIABLE"
                                data = LineInput(1)
                                On Error Resume Next
                                'UPGRADE_WARNING: Couldn't resolve default property of object thisfile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                data_keep(2, i, thisfile) = CDbl(data) / 1000 'data in watts
                                If CDbl(data) / 1000 > qmax Then qmax = CDbl(data) / 1000
                                i = i + 1
                                If i > 2000 Then Exit Do
                            Loop
                        Else
                            Do Until data = "VARIABLE"
                                data = LineInput(1)
                                On Error Resume Next
                                If IsNumeric(data) = True Then
                                    'UPGRADE_WARNING: Couldn't resolve default property of object thisfile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    data_keep(2, i, thisfile) = CSng(data) 'data in kW
                                    If CDbl(data) > qmax Then qmax = CSng(data)
                                End If
                                i = i + 1
                                If i > 2000 Then Exit Do
                            Loop
                        End If
                        'UPGRADE_WARNING: Couldn't resolve default property of object thisfile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        ConePeak(thisfile) = qmax
                        Err.Clear()
                    End If

                    'i = 1
                    'find the SEA data
                    'If StrConv(data, vbUpperCase) = "S.E.A." Then
                    'throw next 2 lines away
                    '    Line Input #1, data
                    '    Line Input #1, data
                    'keep the numbers
                    'If StrConv(data, vbUpperCase) = "W/M^2" Then
                    '         Do Until data$ = "VARIABLE"
                    '             Line Input #1, data
                    '             On Error Resume Next
                    '             data_keep(3, i, thisfile) = data  's.e.a data in m2/kg
                    'If data / 1000 > qmax Then qmax = data / 1000
                    '            i = i + 1
                    '             If i > 2000 Then Exit Do
                    '           Loop
                    'Else
                    '    Do Until data$ = "VARIABLE"
                    '        Line Input #1, data
                    '        On Error Resume Next
                    '        If IsNumeric(data) = True Then
                    '            data_keep(2, i, thisfile) = data 'data in kW
                    '            If data > qmax Then qmax = data
                    '        End If
                    '        i = i + 1
                    '        If i > 2000 Then Exit Do
                    '    Loop
                    'End If
                    'ConePeak(thisfile) = qmax
                    '      Err.Clear
                    ' End If
                Loop
                FileClose(1)
            End If
        End If


        'determine time step at which hrr exceeds 30 kw/m2
        'UPGRADE_WARNING: Couldn't resolve default property of object thisfile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 1 To imax(thisfile)
            'UPGRADE_WARNING: Couldn't resolve default property of object thisfile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If data_keep(2, i, thisfile) >= IgCriteria Then
                'UPGRADE_WARNING: Couldn't resolve default property of object thisfile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ig_step(thisfile) = i
                Exit For
            Else

            End If
        Next i

        'UPGRADE_WARNING: Couldn't resolve default property of object totalfiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object thisfile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If thisfile < totalfiles Then Exit Sub
        S_filename = InputBox("Enter a name to save the file as.", "Create Cone Data File", "conedata")
        'conefile = Left$(opendatafile, (Len(opendatafile) - 3)) + "txt"
        conefile = UserPersonalDataFolder & gcs_folder_ext & "\cone\" & S_filename & ".txt"

        'store required data in text file
        If Confirm_File(conefile, replacefile, 0) >= 0 Then
            FileOpen(1, conefile, OpenMode.Output)
            uniquefiles = 0
            str_Renamed = InputBox("Enter a name to describe the material.", "Material Description", "Material Description")

            WriteLine(1, str_Renamed)
            'UPGRADE_WARNING: Couldn't resolve default property of object totalfiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            WriteLine(1, "Number of HRR Curves", VB6.Format(totalfiles, "00"))
            ReDim igtime(totalfiles)
            'UPGRADE_WARNING: Couldn't resolve default property of object totalfiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For i = 1 To totalfiles
                'get ignition time
                ReDim X(imax(i))
                ReDim Y(imax(i))
                For j = 1 To imax(i)
                    X(j) = data_keep(2, j, i) 'hrr
                    Y(j) = data_keep(1, j, i) 'time
                Next j
                Call Interpolate_D(X, Y, imax(i), IgCriteria, yinterp)
                igtime(i) = yinterp
                If i = 1 Then
                    uniquefiles = uniquefiles + 1
                    WriteLine(1, "Heat Flux", ExtFlux(i))
                    WriteLine(1, "Number of HRR Data Pairs", imax(i) - ig_step(i) + 2)
                    WriteLine(1, "sec,kw/m2")
                    WriteLine(1, 0, IgCriteria)
                    For j = ig_step(i) To imax(i)
                        'Write #1, CSng(Format((data_keep(1, j, i) - data_keep(1, ig_step(i), i)), "0.0")), CSng(Format(data_keep(2, j, i), "0.0"))
                        If j > 0 Then
                            WriteLine(1, CSng(VB6.Format(data_keep(1, j, i) - igtime(i), "0.0")), CSng(VB6.Format(data_keep(2, j, i), "0.0")))
                        End If
                    Next j
                ElseIf i > 1 And ExtFlux(i) = ExtFlux(i - 1) Then
                    'do nothing
                Else
                    uniquefiles = uniquefiles + 1
                    WriteLine(1, "Heat Flux", ExtFlux(i))
                    WriteLine(1, "Number of HRR Data Pairs", imax(i) - ig_step(i) + 2)
                    WriteLine(1, "sec,kw/m2")
                    WriteLine(1, 0, IgCriteria)
                    For j = ig_step(i) To imax(i)
                        'Write #1, CSng(Format((data_keep(1, j, i) - data_keep(1, ig_step(i), i)), "0.0")), CSng(Format(data_keep(2, j, i), "0.0"))
                        If j > 0 Then
                            WriteLine(1, CSng(VB6.Format(data_keep(1, j, i) - igtime(i), "0.0")), CSng(VB6.Format(data_keep(2, j, i), "0.0")))
                        End If
                    Next j
                End If
            Next i
            WriteLine(1, "Ignition Data")
            WriteLine(1, "Number of Pairs", totalfiles)
            WriteLine(1, "flux kw/m2,ignition time sec, peak hrr kw/m2")
            'UPGRADE_WARNING: Couldn't resolve default property of object totalfiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For i = 1 To totalfiles
                'Write #1, ExtFlux(i), CSng(Format(data_keep(1, ig_step(i), i), "0.0")), CSng(Format(ConePeak(i), "0.0"))
                WriteLine(1, ExtFlux(i), CSng(VB6.Format(igtime(i), "0.0")), CSng(VB6.Format(ConePeak(i), "0.0")))
            Next i
            WriteLine(1, "Flame Spread Parameter", 0)
            WriteLine(1, "Min Surface Temp For Spread", 0)
            temp = 0
            'UPGRADE_WARNING: Couldn't resolve default property of object totalfiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For i = 1 To totalfiles
                temp = ConeHoC(i) + temp
            Next i
            'UPGRADE_WARNING: Couldn't resolve default property of object totalfiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            temp = temp / totalfiles
            WriteLine(1, "Effective Heat of Combustion", VB6.Format(temp, "0.0"))
            temp = 0
            'UPGRADE_WARNING: Couldn't resolve default property of object totalfiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For i = 1 To totalfiles
                temp = ConeSEA(i) + temp
            Next i
            'UPGRADE_WARNING: Couldn't resolve default property of object totalfiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            temp = temp / totalfiles
            WriteLine(1, "Avg Specific Extinction Area", VB6.Format(temp, "0.0"))
            Seek(1, 1)
            WriteLine(1, str_Renamed)
            WriteLine(1, "Number of HRR Curves", VB6.Format(uniquefiles, "00"))
            FileClose(1)
            MsgBox("The template file " & conefile & " has been created, containing the necessary cone calorimeter data for the material. Edit this file as required.", MsgBoxStyle.OkOnly)
        End If
        Erase data_keep
        Erase imax
        Erase ig_step
        Erase X
        Erase Y
        Erase igtime
        Exit Sub

errhandler:
        'User pressed Cancel button
        'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'System.Windows.Forms.Cursor.Current = Default_Renamed
        If Err.Number = 13 Then 'type mismatch
            MsgBox("Please enter a number greater than zero.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
            Err.Clear()
            GoTo 10
            'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
        ElseIf Err.Number = DialogResult.Cancel Then
            Err.Clear()
            Exit Sub
        Else
            MsgBox("Error - operation cancelled", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
        End If
        FileClose(1)
        Err.Clear()
        Exit Sub

ConfirmFileError2:
        Action = File_Errors(Err.Number)
        Select Case Action
            Case 1
                Resume Next
            Case Else
                'Screen.MousePointer = DEFAULT
                FileClose(1)
                MsgBox("Error - material cone data file not created", MsgBoxStyle.OkOnly)
                Erase data_keep
                Erase imax
                Erase ig_step
                Exit Sub
        End Select

    End Sub

    Public Sub Interpolate_D(ByRef X() As Double, ByRef Y() As Double, ByVal n As Integer, ByVal Xint As Double, ByRef yinterp As Double)
        '****************************************************************************
        '*  From ProMath 2.0
        '*  This Subroutine Performs Linear Interpolation Within A Set Of X(),Y()
        '*  Pairs To Give The Y Value Corresponding To Xint.
        '*       X       Array Of Values Of The Independent Variable
        '*       Y       Array Of Values Of The Dependent Variable Corresponding To X
        '*       N       Number Of Points To Interpolate.
        '*       Xint    The X-value For Which Estimate Of Y Is Desired.
        '*       Yinterp    The Y-value Returned To The User.
        '*
        '*  Revised Colleen Wade 12 October 1996
        '****************************************************************************

        Dim i As Integer
        Dim y1, x1, x2, y2 As Double

        yinterp = 0

        For i = 1 To n - 1
            If X(i) <= Xint And X(i + 1) >= Xint Then
                x1 = X(i)
                x2 = X(i + 1)
                y1 = Y(i)
                y2 = Y(i + 1)
                yinterp = (Xint - x1) * (y2 - y1) / (x2 - x1) + y1
                Exit Sub
            End If
        Next i

    End Sub

    Public Sub Graph_Data_3D(ByVal id As Short, ByVal Title As String, ByVal datatobeplotted(,,) As Double, ByRef DataShift As Double, ByRef DataMultiplier As Double)
        '*  ====================================================================
        '*  This function takes data for a variable from a three-dimensional array
        '*  and displays it in a graph
        '*  ====================================================================

        Dim j As Integer
        Dim room As Integer
        Dim ydata(0 To NumberTimeSteps) As Double

        'if no data exists
        If NumberTimeSteps < 1 Then
            MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Sub
        End If

        Try

            frmPlot.Chart1.Series.Clear()
            room = 1
            For room = 1 To NumberRooms

                frmPlot.Chart1.Series.Add("Room " & room)
                frmPlot.Chart1.Series("Room " & room).ChartType = SeriesChartType.FastLine

                For j = 1 To NumberTimeSteps

                    ydata(j) = datatobeplotted(room, id, j) * DataMultiplier + DataShift 'data to be plotted
                    frmPlot.Chart1.Series("Room " & room).Points.AddXY(tim(j, 1), ydata(j))
                Next

            Next room
            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            'frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            frmPlot.Chart1.Titles("Title1").Text = Title
            ' Disable X axis margin

            frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Graph_Data_3D")
        End Try


    End Sub
    Public Sub Graph_Data_NHL(ByVal id As Short, ByVal Title As String, ByVal datatobeplotted(,,) As Double, ByRef DataShift As Double, ByRef DataMultiplier As Double)
        '*  ====================================================================
        '*  This function takes data for a variable from a three-dimensional array
        '*  and displays it in a graph
        '*  ====================================================================

        Dim j As Integer
        Dim room As Integer
        Dim ydata(0 To NumberTimeSteps) As Double

        'if no data exists
        If NumberTimeSteps < 1 Then
            MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Sub
        End If

        Try

            frmPlot.Chart1.Series.Clear()
            room = fireroom
            For room = 1 To NumberRooms

                frmPlot.Chart1.Series.Add("Room " & room)
                frmPlot.Chart1.Series("Room " & room).ChartType = SeriesChartType.FastLine

                If Not roomcolor(room - 1).IsEmpty Then
                    frmPlot.Chart1.Series("Room " & room).Color = roomcolor(room - 1) '=line color
                End If

                For j = 1 To NumberTimeSteps

                    ydata(j) = datatobeplotted(id, room, j) * DataMultiplier + DataShift 'data to be plotted
                    frmPlot.Chart1.Series("Room " & room).Points.AddXY(tim(j, 1), ydata(j))
                Next

            Next room
            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            'frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            frmPlot.Chart1.Titles("Title1").Text = Title
            ' Disable X axis margin

            frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Graph_Data_NHL")
        End Try


    End Sub
    Public Sub Graph_Data_NHLmulti(ByVal id As Short, ByVal Title As String, ByVal datatobeplotted(,,) As Double, ByRef DataShift As Double, ByRef DataMultiplier As Double, ByRef XMultiplier As Double)
        '*  ====================================================================
        '*  This function takes data for a variable from a three-dimensional array
        '*  and displays it in a graph
        '*  ====================================================================

        Dim j As Integer
        Dim room As Integer
        Dim ydata(0 To NumberTimeSteps) As Double
        Dim Desc1 As String = "series 1"

        'if no data exists
        If NumberTimeSteps < 1 Then
            MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Sub
        End If

        Try

            frmPlot.Chart1.Series.Clear()
            room = fireroom

            For id = 0 To 4
                If id = 1 Then Desc1 = "ceiling"
                If id = 2 Then Desc1 = "upper wall"
                If id = 3 Then Desc1 = "lower wall"
                If id = 4 Then Desc1 = "floor"
                If id = 0 Then Desc1 = "average"

                frmPlot.Chart1.Series.Add(Desc1)
                frmPlot.Chart1.Series(Desc1).ChartType = SeriesChartType.FastLine

                If Not roomcolor(room - 1).IsEmpty Then
                    frmPlot.Chart1.Series(Desc1).Color = roomcolor(id - 1) '=line color
                End If

                For j = 1 To NumberTimeSteps

                    ydata(j) = datatobeplotted(id, room, j) * DataMultiplier + DataShift 'data to be plotted
                    frmPlot.Chart1.Series(Desc1).Points.AddXY(tim(j, 1) * XMultiplier, ydata(j))

                Next

            Next id

            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            'frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (minutes)"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            frmPlot.Chart1.Titles("Title1").Text = Title
            ' Disable X axis margin

            frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Graph_Data_NHL")
        End Try


    End Sub
    Public Sub Graph_Data_3D_CJET(ByVal id As Short, ByVal Title As String, ByVal datatobeplotted(,,) As Double, ByRef DataShift As Double, ByRef DataMultiplier As Double)
        '*  ====================================================================
        '*  This function takes data for a variable from a three-dimensional array
        '*  and displays it in a graph
        '*  ====================================================================

        Dim j As Integer = 0
        Dim room As Integer = 0
        Dim ydata(0 To NumberTimeSteps) As Double

        'if no data exists
        If NumberTimeSteps < 1 Then
            MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Sub
        End If

        Try

            frmPlot.Chart1.Series.Clear()

            For i = 1 To NumSprinklers

                frmPlot.Chart1.Series.Add("Spr " & i.ToString)
                frmPlot.Chart1.Series("Spr " & i.ToString).ChartType = SeriesChartType.FastLine

                For j = 1 To NumberTimeSteps

                    ydata(j) = datatobeplotted(j, id, i - 1) * DataMultiplier + DataShift 'data to be plotted
                    frmPlot.Chart1.Series("Spr " & i).Points.AddXY(tim(j, 1), ydata(j))
                    'If tim(j, 1) = 225 Then Stop
                Next

            Next

            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            'frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            frmPlot.Chart1.Titles("Title1").Text = Title
            ' Disable X axis margin

            frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Graph_Data_3D")
        End Try


    End Sub

    Public Sub Interpolate_hrr(ByVal step_Renamed As Integer, ByRef X() As Double, ByVal surface As String, ByVal Xint As Single, ByRef yinterp As Double)
        '****************************************************************************
        '*  From ProMath 2.0
        '*  This Subroutine Performs Linear Interpolation Within A Set Of X(),Y()
        '*  Pairs To Give The Y Value Corresponding To Xint.
        '*       X       Array Of Values Of The Independent Variable (a 2D array)
        '*       Y       Array Of Values Of The Dependent Variable Corresponding To X
        '*       N       Number Of Points To Interpolate.
        '*       Xint    The X-value For Which Estimate Of Y Is Desired.
        '*       Yinterp    The Y-value Returned To The User.
        '*
        '****************************************************************************
        'unused code

        Dim i As Short
        Dim y1, x1, x2, y2 As Double
        Dim Y() As Double
        ReDim Y(CurveNumber_W(fireroom))

        For i = 1 To CurveNumber_W(fireroom)
            If surface = "wall" Then
                If step_Renamed > ConeNumber_W(i, fireroom) Then step_Renamed = ConeNumber_W(i, fireroom)
                Y(i) = ConeXW(fireroom, i, step_Renamed)
            ElseIf surface = "ceiling" Then
                If step_Renamed > ConeNumber_C(i, fireroom) Then step_Renamed = ConeNumber_C(i, fireroom)
                Y(i) = ConeXC(fireroom, i, step_Renamed)
            Else
            End If
            If CurveNumber_W(fireroom) = 1 Then yinterp = Y(i) : Exit Sub
        Next i

        yinterp = 0.0#

        For i = 1 To CurveNumber_W(fireroom) - 1
            If X(i) <= Xint And X(i + 1) >= Xint Then
                x1 = X(i)
                x2 = X(i + 1)
                y1 = Y(i)
                y2 = Y(i + 1)
                yinterp = (Xint - x1) * (y2 - y1) / (x2 - x1) + y1
                Exit Sub
            ElseIf Xint < X(i) Then
                yinterp = Y(i)
                Exit Sub
            ElseIf Xint > X(i + 1) Then
                yinterp = Y(i + 1)
                Exit Sub
            Else
            End If
        Next i

    End Sub





    Public Sub Graph_Data20(ByRef graph As System.Windows.Forms.Control, ByRef YAxisTitle As String, ByRef datatobeplotted() As Double, ByRef DataShift As Single, ByRef multiplier As Single, ByRef Style As Short, ByRef ymax As Single)
        '		'*  ====================================================================
        '		'*  This function takes data for a variable from a one-dimensional array
        '		'*  and displays it in a run-time graph
        '		'*
        '		'*  Revised 9 September 1996 Colleen Wade
        '		'*  ====================================================================

        '		Static NRooms, i, factor As Short
        '		Static j, k As Short

        '		On Error GoTo graphhandler1

        '		'if no data exists
        '		If NumberTimeSteps < 1 Then Exit Sub
        '		'If stepcount < 2 Then i = 1: Exit Sub

        '		With graph
        '			'.NumSets = NumberRooms
        '			'.AutoInc = 0
        '			'.ThickLines = 1
        '			If Style = 2 Then
        '				'UPGRADE_WARNING: Couldn't resolve default property of object graph.YAxisMax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '				.YAxisMax = ymax
        '			End If
        '			'UPGRADE_WARNING: Couldn't resolve default property of object graph.LineStats. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '			.LineStats = 0
        '		End With
        '		If stepcount = 1 Then
        '			factor = 1
        '			i = 1
        '			'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '			If graph.NumSets > 1 Then
        '				NRooms = NumberRooms
        '			Else
        '				NRooms = 1
        '			End If
        '		End If
        'here1: 

        '		Do While (NRooms * stepcount / factor) > 3800
        '			factor = factor * 2
        '			i = 1
        '		Loop 

        '		If Int(stepcount / factor) < 2 Then Exit Sub
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.NumPoints = Int(stepcount / factor)
        '		k = i
        '		With graph
        '			For j = 1 To NRooms
        '				i = k
        '				'UPGRADE_WARNING: Couldn't resolve default property of object graph.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '				.ThisSet = j
        '				'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '				While i <= .NumPoints
        '					'UPGRADE_WARNING: Couldn't resolve default property of object graph.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '					.ThisPoint = i
        '					'UPGRADE_WARNING: Couldn't resolve default property of object graph.Graphdata. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '					.Graphdata = (multiplier * datatobeplotted(j, factor * i) + DataShift)
        '					'UPGRADE_WARNING: Couldn't resolve default property of object graph.XPosData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '					.XPosData = (tim(factor * i, 1))
        '					i = i + 1
        '				End While
        '			Next j
        '			'UPGRADE_ISSUE: Control method graph.DrawMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        '			.DrawMode = 3

        '		End With
        '		Exit Sub

        'graphhandler1: 
        '		GoTo here1

    End Sub

    Public Sub Graph_Data21(ByRef graph As System.Windows.Forms.Control, ByRef YAxisTitle As String, ByRef datatobeplotted() As Double, ByRef DataShift As Single, ByRef multiplier As Single, ByRef Style As Short, ByRef ymax As Single)
        '		'*  ====================================================================
        '		'*  This function takes data for a variable from a one-dimensional array
        '		'*  and displays it in a run-time graph
        '		'*
        '		'*  Revised 9 September 1996 Colleen Wade
        '		'*  ====================================================================

        '		Static NRooms, i, factor As Short
        '		Static j, k As Short

        '		On Error GoTo graphhandler1

        '		'if no data exists
        '		If NumberTimeSteps < 1 Then Exit Sub
        '		'If stepcount < 2 Then i = 1: Exit Sub

        '		With graph
        '			'UPGRADE_WARNING: Couldn't resolve default property of object graph.LineStats. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '			.LineStats = 0
        '			'.NumSets = NumberRooms
        '			'.AutoInc = 0
        '			'.ThickLines = 1
        '			If Style = 2 Then
        '				'UPGRADE_WARNING: Couldn't resolve default property of object graph.YAxisMax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '				.YAxisMax = ymax
        '			End If
        '		End With
        '		If stepcount = 1 Then
        '			factor = 1
        '			i = 1
        '			'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '			If graph.NumSets > 1 Then
        '				NRooms = NumberRooms
        '			Else
        '				NRooms = 1
        '			End If
        '		End If
        'here1: 

        '		Do While (NRooms * stepcount / factor) > 3800
        '			factor = factor * 2
        '			i = 1
        '		Loop 

        '		If Int(stepcount / factor) < 2 Then Exit Sub
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.NumPoints = Int(stepcount / factor)
        '		k = i
        '		With graph
        '			For j = 1 To NRooms
        '				i = k
        '				'UPGRADE_WARNING: Couldn't resolve default property of object graph.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '				.ThisSet = j
        '				'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '				While i <= .NumPoints
        '					'UPGRADE_WARNING: Couldn't resolve default property of object graph.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '					.ThisPoint = i
        '					'UPGRADE_WARNING: Couldn't resolve default property of object graph.Graphdata. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '					.Graphdata = (multiplier * datatobeplotted(j, factor * i) + DataShift)
        '					'UPGRADE_WARNING: Couldn't resolve default property of object graph.XPosData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '					.XPosData = (tim(factor * i, 1))
        '					i = i + 1
        '				End While
        '			Next j
        '			'UPGRADE_ISSUE: Control method graph.DrawMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        '			.DrawMode = 3
        '		End With
        '		Exit Sub

        'graphhandler1: 
        '		GoTo here1

    End Sub

    Public Sub Graph_Data22(ByRef graph As System.Windows.Forms.Control, ByRef YAxisTitle As String, ByRef datatobeplotted() As Double, ByRef DataShift As Single, ByRef multiplier As Single, ByRef Style As Short, ByRef ymax As Single)
        '		'*  ====================================================================
        '		'*  This function takes data for a variable from a one-dimensional array
        '		'*  and displays it in a run-time graph
        '		'*
        '		'*  Revised 9 September 1996 Colleen Wade
        '		'*  ====================================================================

        '		Static NRooms, i, factor As Short
        '		Static j, k As Short

        '		On Error GoTo graphhandler1

        '		'if no data exists
        '		If NumberTimeSteps < 1 Then Exit Sub
        '		'If stepcount < 2 Then i = 1: Exit Sub

        '		With graph
        '			'UPGRADE_WARNING: Couldn't resolve default property of object graph.LineStats. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '			.LineStats = 0
        '			'.NumSets = NumberRooms
        '			'.AutoInc = 0
        '			'.ThickLines = 1
        '			If Style = 2 Then
        '				'UPGRADE_WARNING: Couldn't resolve default property of object graph.YAxisMax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '				.YAxisMax = ymax
        '			End If
        '		End With
        '		If stepcount = 1 Then
        '			factor = 1
        '			i = 1
        '			'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '			If graph.NumSets > 1 Then
        '				NRooms = NumberRooms
        '			Else
        '				NRooms = 1
        '			End If
        '		End If
        'here1: 


        '		Do While (NRooms * stepcount / factor) > 3800
        '			factor = factor * 2
        '			i = 1
        '		Loop 

        '		If Int(stepcount / factor) < 2 Then Exit Sub
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.NumPoints = Int(stepcount / factor)
        '		k = i
        '		With graph
        '			For j = 1 To NRooms
        '				i = k
        '				'UPGRADE_WARNING: Couldn't resolve default property of object graph.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '				.ThisSet = j
        '				'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '				While i <= .NumPoints
        '					'UPGRADE_WARNING: Couldn't resolve default property of object graph.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '					.ThisPoint = i
        '					'UPGRADE_WARNING: Couldn't resolve default property of object graph.Graphdata. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '					.Graphdata = (multiplier * datatobeplotted(j, factor * i) + DataShift)
        '					'UPGRADE_WARNING: Couldn't resolve default property of object graph.XPosData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '					.XPosData = (tim(factor * i, 1))
        '					i = i + 1
        '				End While
        '			Next j
        '			'UPGRADE_ISSUE: Control method graph.DrawMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        '			.DrawMode = 3
        '		End With
        '		Exit Sub

        'graphhandler1: 
        '		GoTo here1


    End Sub

    Public Sub GRaph_Data23(ByRef graph As System.Windows.Forms.Control, ByRef YAxisTitle As String, ByRef datatobeplotted() As Double, ByRef DataShift As Single, ByRef multiplier As Single, ByRef Style As Short, ByRef ymax As Single)
        '		'*  ====================================================================
        '		'*  This function takes data for a variable from a one-dimensional array
        '		'*  and displays it in a run-time graph
        '		'*
        '		'*  Revised 9 September 1996 Colleen Wade
        '		'*  ====================================================================

        '		Static NRooms, i, factor As Short
        '		Static j, k As Short

        '		On Error GoTo graphhandler1

        '		'if no data exists
        '		If NumberTimeSteps < 1 Then Exit Sub
        '		'If stepcount < 2 Then i = 1: Exit Sub

        '		With graph
        '			'UPGRADE_WARNING: Couldn't resolve default property of object graph.LineStats. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '			.LineStats = 0
        '			'.NumSets = NumberRooms
        '			'.AutoInc = 0
        '			'.ThickLines = 1
        '			If Style = 2 Then
        '				'UPGRADE_WARNING: Couldn't resolve default property of object graph.YAxisMax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '				.YAxisMax = ymax
        '			End If
        '		End With
        '		If stepcount = 1 Then
        '			factor = 1
        '			i = 1
        '			'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '			If graph.NumSets > 1 Then
        '				NRooms = NumberRooms
        '			Else
        '				NRooms = 1
        '			End If
        '		End If
        'here1: 


        '		Do While (NRooms * stepcount / factor) > 3800
        '			factor = factor * 2
        '			i = 1
        '		Loop 

        '		If Int(stepcount / factor) < 2 Then Exit Sub
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.NumPoints = Int(stepcount / factor)
        '		k = i
        '		With graph
        '			For j = 1 To NRooms
        '				i = k
        '				'UPGRADE_WARNING: Couldn't resolve default property of object graph.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '				.ThisSet = j
        '				'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '				While i <= .NumPoints
        '					'UPGRADE_WARNING: Couldn't resolve default property of object graph.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '					.ThisPoint = i
        '					'UPGRADE_WARNING: Couldn't resolve default property of object graph.Graphdata. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '					.Graphdata = (multiplier * datatobeplotted(j, factor * i) + DataShift)
        '					'UPGRADE_WARNING: Couldn't resolve default property of object graph.XPosData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '					.XPosData = (tim(factor * i, 1))
        '					i = i + 1
        '				End While
        '			Next j
        '			'UPGRADE_ISSUE: Control method graph.DrawMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        '			.DrawMode = 3
        '		End With
        '		Exit Sub

        'graphhandler1: 
        '		GoTo here1


    End Sub

    Function COSH(ByVal X As Double) As Double

        ' Hyperbolic cosine function

        COSH = 0.5 * (Exp(X) + Exp(-X))

    End Function

    Function ACOSH(ByVal X As Double) As Double

        ' Inverse hyperbolic cosine function

        If (Abs(X) < 1.0#) Then
            ACOSH = 0.0#
        Else
            ACOSH = Log(X + Sqrt(X * X - 1.0#))
        End If

    End Function


    Public Sub Graph_ArrayBL(ByRef graph As System.Windows.Forms.Control, ByRef comboGraphIndex As Short, ByRef Graphdata() As Double, ByRef Shift As Single, ByRef multiplier As Single)
        '		'===================================================================
        '		'   this procedure allows the user to switch to display another graph
        '		'===================================================================
        '		Dim NRooms, factor, k As Short
        '		Dim i As Integer
        '		Dim room As Short
        '		Dim j As Integer
        '		ReDim Preserve Graphdata(MaxNumberRooms, stepcount)
        '		On Error GoTo errorhandler

        '		If stepcount <= 2 Then Exit Sub
        '		'If stepcount = 2 Then
        '		'room = frmOptions1.lstGraphRoom.ListIndex + 1
        '		i = 1
        '		'End If
        '		j = i
        '		For room = 1 To NumberRooms
        '			i = j
        '			Select Case comboGraphIndex
        '				Case 1
        '					While i <= stepcount
        '						Graphdata(room, i) = layerheight(room, i)
        '						Shift = 0
        '						multiplier = 1
        '						i = i + 1
        '					End While
        '				Case 2
        '					While i <= stepcount
        '						Graphdata(room, i) = uppertemp(room, i)
        '						Shift = -273
        '						multiplier = 1
        '						i = i + 1
        '					End While
        '				Case 3
        '					While i <= stepcount
        '						Graphdata(room, i) = lowertemp(room, i)
        '						Shift = -273
        '						multiplier = 1
        '						i = i + 1
        '					End While
        '				Case 4
        '					While i <= stepcount
        '						Graphdata(room, i) = HeatRelease(room, i, 2)
        '						Shift = 0
        '						multiplier = 1
        '						i = i + 1
        '					End While
        '				Case 5
        '					While i <= stepcount
        '						'Graphdata(room, i) = massplumeflow(i, fireroom)
        '						Graphdata(room, i) = massplumeflow(i, room)
        '						Shift = 0
        '						multiplier = 1
        '						i = i + 1
        '					End While
        '				Case 6
        '					While i <= stepcount
        '						Graphdata(room, i) = FlowToUpper(room, i)
        '						Shift = 0
        '						multiplier = 1
        '						i = i + 1
        '					End While
        '				Case 7
        '					While i <= stepcount
        '						Graphdata(room, i) = FlowToLower(room, i)
        '						Shift = 0
        '						multiplier = 1
        '						i = i + 1
        '					End While
        '				Case 8
        '					While i <= stepcount
        '						Graphdata(room, i) = UFlowToOutside(room, i)
        '						Shift = 0
        '						multiplier = 1
        '						i = i + 1
        '					End While
        '				Case 9
        '					While i <= stepcount
        '						Graphdata(room, i) = O2MassFraction(room, i, 1)
        '						Shift = 0
        '						multiplier = 100 * MolecularWeightAir / MolecularWeightO2
        '						i = i + 1
        '					End While
        '				Case 10
        '					While i <= stepcount
        '						Graphdata(room, i) = O2MassFraction(room, i, 2)
        '						Shift = 0
        '						multiplier = 100 * MolecularWeightAir / MolecularWeightO2
        '						i = i + 1
        '					End While
        '				Case 11
        '					While i <= stepcount
        '						Graphdata(room, i) = CO2MassFraction(room, i, 1)
        '						Shift = 0
        '						multiplier = 100 * MolecularWeightAir / MolecularWeightCO2
        '						i = i + 1
        '					End While
        '				Case 12
        '					While i <= stepcount
        '						Graphdata(room, i) = CO2MassFraction(room, i, 2)
        '						Shift = 0
        '						multiplier = 100 * MolecularWeightAir / MolecularWeightCO2
        '						i = i + 1
        '					End While
        '				Case 13
        '					While i <= stepcount
        '						Graphdata(room, i) = COMassFraction(room, i, 1)
        '						Shift = 0
        '						multiplier = 100 * MolecularWeightAir / MolecularWeightCO
        '						i = i + 1
        '					End While
        '				Case 14
        '					While i <= stepcount
        '						Graphdata(room, i) = COMassFraction(room, i, 2)
        '						Shift = 0
        '						multiplier = 100 * MolecularWeightAir / MolecularWeightCO
        '						i = i + 1
        '					End While
        '				Case 15
        '					While i < stepcount
        '						Graphdata(room, i) = OD_upper(room, i)
        '						Shift = 0
        '						multiplier = 1
        '						i = i + 1
        '						ReDim Preserve Graphdata(MaxNumberRooms, stepcount - 1)
        '					End While
        '				Case 16
        '					While i < stepcount
        '						Graphdata(room, i) = OD_lower(room, i)
        '						Shift = 0
        '						multiplier = 1
        '						i = i + 1
        '						ReDim Preserve Graphdata(MaxNumberRooms, stepcount - 1)
        '					End While
        '				Case 17
        '					While i < stepcount
        '						Graphdata(room, i) = Visibility(room, i)
        '						Shift = 0
        '						multiplier = 1
        '						i = i + 1
        '						ReDim Preserve Graphdata(MaxNumberRooms, stepcount - 1)
        '					End While
        '				Case 18
        '					While i <= stepcount
        '						Graphdata(room, i) = RoomPressure(room, i)
        '						Shift = 0
        '						multiplier = 1
        '						i = i + 1
        '					End While
        '			End Select
        '		Next room

        '		With graph
        '			'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '			.NumSets = NumberRooms
        '			If frmOptions1.optThickLines.Checked = True Then
        '				'UPGRADE_WARNING: Couldn't resolve default property of object graph.ThickLines. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '				.ThickLines = 1
        '			Else
        '				'UPGRADE_WARNING: Couldn't resolve default property of object graph.ThickLines. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '				.ThickLines = 0
        '			End If
        '			'UPGRADE_WARNING: Couldn't resolve default property of object graph.GraphStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '			.GraphStyle = 4
        '		End With
        '		'If stepcount = 1 Then
        '		factor = 1
        '		i = 1
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		If graph.NumSets > 1 Then
        '			NRooms = NumberRooms
        '		Else
        '			NRooms = 1
        '		End If
        '		'End If

        'here1: 

        '		Do While (NRooms * stepcount / factor) > 3800
        '			factor = factor * 2
        '			i = 1
        '		Loop 

        '		If Int(stepcount / factor) < 2 Then Exit Sub
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.NumPoints = Int(stepcount / factor)
        '		If UBound(Graphdata, 2) < stepcount Then
        '			'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '			graph.NumPoints = graph.NumPoints - 1
        '		End If
        '		k = i
        '		With graph
        '			For j = 1 To NRooms
        '				i = k
        '				'UPGRADE_WARNING: Couldn't resolve default property of object graph.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '				.ThisSet = j
        '				'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '				While i <= .NumPoints
        '					'UPGRADE_WARNING: Couldn't resolve default property of object graph.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '					.ThisPoint = i
        '					'UPGRADE_WARNING: Couldn't resolve default property of object graph.Graphdata. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '					.Graphdata = (multiplier * Graphdata(j, factor * i) + Shift)
        '					'UPGRADE_WARNING: Couldn't resolve default property of object graph.XPosData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '					.XPosData = (tim(factor * i, 1))
        '					i = i + 1
        '				End While
        '			Next j
        '			'UPGRADE_ISSUE: Control method graph.DrawMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        '			.DrawMode = 3
        '		End With
        '		Exit Sub

        'errorhandler: 

    End Sub

    Public Function JET_CJmax(ByVal Q As Double, ByVal room As Integer, ByVal tim As Double, ByVal layerz As Double, ByVal UT As Double, ByVal LT As Double, ByVal LinkTemp As Double) As Double
        '*  =====================================================================
        '*      Find the maximum ceiling jet temperature in the presence of a hot layer
        '*
        '*       Based on SFPE handbook 2nd ed, 2-37, ceiling jet flows
        '*  =====================================================================

        Dim Qi1, beta, CT, Qi2 As Double
        Dim Zi2, H2, Zi1, H1 As Single
        Dim alpha As Single
        Dim epsilon As Single
        Dim delTo As Double

        On Error GoTo errhandler

        If FireLocation(1) = 0 Then
            Q = Q
        ElseIf FireLocation(1) = 1 Then
            Q = 2 * Q 'wall fire
        Else
            Q = 4 * Q 'corner fire
        End If

        beta = 0.9555 'velocity to temperature ratio of gaussian profile half widths
        CT = 9.115
        alpha = 0.44

        If room = fireroom Then
            'get heat release rate for object 1 at this time
            'Q = Get_HRR(1, tim, 0, 0, 0, 0)
            If Q = 0 Then
                JET_CJmax = UT
                Exit Function
            End If

            'plume centreline temperature
            epsilon = UT / LT
            Zi1 = layerz - FireHeight(1)
            If Zi1 < 0 Then Exit Function
            H1 = RoomHeight(room) - FireHeight(1)
            'Qi1 = (1 - RadiantLossFraction) * Q / ((Atm_Pressure) / (Gas_Constant / MW_air) * SpecificHeat_air * Sqr(G) * Zi1 ^ (5 / 2))
            Qi1 = (1 - NewRadiantLossFraction(1)) * Q / ((Atm_Pressure) / (Gas_Constant / MW_air) * SpecificHeat_air * Sqrt(G) * Zi1 ^ (5 / 2))
            If ((1 + CT * Qi1 ^ (2 / 3)) / (epsilon * CT) - 1 / CT) > 0 Then
                Qi2 = ((1 + CT * Qi1 ^ (2 / 3)) / (epsilon * CT) - 1 / CT) ^ (3 / 2)
                Zi2 = (epsilon * Qi1 * CT) / (Qi2 ^ (1 / 3) * ((epsilon - 1) * (beta ^ 2 + 1) + epsilon * CT * Qi2 ^ (2 / 3))) ^ (2 / 5) * Zi1
            Else
                Qi2 = 0
                Zi2 = 0
            End If

            'dimensionalise again
            Qi2 = Qi2 * (Atm_Pressure) / (Gas_Constant / MW_air) * SpecificHeat_air * Sqrt(G) * Zi2 ^ (5 / 2)
            H2 = H1 - Zi1 + Zi2

            delTo = (0.188) ^ (-4 / 3)
            'UPGRADE_WARNING: Couldn't resolve default property of object JET_CJmax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            JET_CJmax = UT + delTo * UT * (Qi2 / ((Atm_Pressure) / (Gas_Constant / MW_air) * SpecificHeat_air * Sqrt(G) * H2 ^ (5 / 2))) ^ (2 / 3)

        Else
            'not the fireroom
            'UPGRADE_WARNING: Couldn't resolve default property of object JET_CJmax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            JET_CJmax = UT
        End If
        Exit Function

errhandler:
        If Err.Number = 6 Then
            Err.Clear()
            Resume Next
        Else
            MsgBox(ErrorToString(Err.Number) & " in subroutine JET_CJmax")
        End If
    End Function

    Public Sub Graph_Ign_Data(ByRef label As String, ByRef caption As String, ByRef IgnPoints As Short, ByRef YAxisTitle As String, ByRef xdata() As Double, ByRef ydata() As Double, ByRef DataShift As Single, ByRef multiplier As Single, ByRef Style As Short)

        ' Dim i As Integer
        Dim j As Integer

        If IgnPoints < 2 Then
            MsgBox("There is insufficient data to plot.", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Try

            frmPlot.Chart1.Series.Clear()
            frmPlot.Chart1.Series.Add("1")
            frmPlot.Chart1.Series("1").ChartType = SeriesChartType.Point

            For j = 1 To IgnPoints
                frmPlot.Chart1.Series("1").Points.AddXY(xdata(j), ydata(j))
            Next

            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = YAxisTitle
            'frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelAutoFitMaxFontSize = 12
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.000"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = Math.Ceiling(xdata.Max)
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Minimum = 0
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Minimum = 0
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.LabelStyle.Format = "0"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Maximum = ydata.Max

            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "External Flux (kW/m²)" & label
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False

            frmPlot.Chart1.Legends("Legend1").Enabled = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            frmPlot.Chart1.Titles("Title1").Text = "Derived Material Properties"
            frmPlot.Chart1.DataManipulator.FinancialFormula(FinancialFormula.Forecasting, "linear,,true,true", "1:Y,Range", "2:Y,Range")

            frmPlot.Chart1.Series("2").ChartType = SeriesChartType.Line

            frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Graph_Ign_Data")
        End Try



        'Dim chdata(0 To IgnPoints - 1, 0 To 1) As Double

        ' With frmGraph.AxMSChart1
        '.chartType = MSChart20Lib.VtChChartType.VtChChartType2dXY


        '.Visible = True
        '.TitleText = "Derived Material Properties"
        '.Plot.UniformAxis = False

        '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).AxisTitle.Text = "External Flux (kW/m²)" & label
        '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).AxisTitle.Text = YAxisTitle

        '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).ValueScale.Auto = False
        '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).ValueScale.Minimum = 0
        '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).ValueScale.Maximum = ydata.Max
        '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).ValueScale.MajorDivision = 10

        '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).ValueScale.Auto = False
        '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).ValueScale.Minimum = 0
        '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).ValueScale.Maximum = Math.Ceiling(xdata.Max)
        '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).ValueScale.MajorDivision = CInt(0.1 * xdata.Max)

        '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).AxisTitle.TextLayout.Orientation = MSChart20Lib.VtOrientation.VtOrientationHorizontal

        '.Plot.SeriesCollection(1).SeriesMarker.Auto = False
        '.Plot.SeriesCollection(1).DataPoints(-1).Marker.Style = MSChart20Lib.VtMarkerStyle.VtMarkerStyleFilledCircle
        '.Plot.SeriesCollection(1).DataPoints(-1).Marker.Size = 150
        '.Plot.SeriesCollection(1).DataPoints(-1).Marker.Visible = True

        ' .Plot.DataSeriesInRow = False
        ' .Plot.SeriesCollection(1).ShowLine = False
        '.Plot.SeriesCollection(1).StatLine.Flag = MSChart20Lib.VtChStats.VtChStatsRegression
        '.Legend.Location.Visible = False

        'chdata(0, 0) = "ignition data"
        'chdata(0, 1) = "ignition data"
        'For i = 1 To IgnPoints
        '    chdata(i - 1, 0) = xdata(i)
        '    chdata(i - 1, 1) = ydata(i)
        'Next i

        '.ChartData = chdata

        'End With
        'frmGraph.Text = label
        'frmGraph.Show()

        '		'UPGRADE_WARNING: Couldn't resolve default property of object frmGraphs.Show. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		frmGraphs.Show()
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.AutoInc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.AutoInc = 1
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.LeftTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.LeftTitle = YAxisTitle
        '		'graph.GraphTitle = Left$(Description, 80)
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.GraphTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.GraphTitle = "Derived Material Properties"
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.ThickLines. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.ThickLines = 1
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.FontStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.FontStyle = 0
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.Ticks. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.Ticks = 1
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.FontUse. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.FontUse = GraphLib.FontUseConstants.gphAllText
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.FontFamily. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.FontFamily = GraphLib.FontFamilyConstants.gphSwiss
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.BottomTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.BottomTitle = "External Flux (kW/m²)" & label
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.GraphStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.GraphStyle = 1
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.LineStats. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.LineStats = 8
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.DataReset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.DataReset = 9


        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.NumPoints = maxpoints
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.NumSets = 1

        '		If Style = 2 Then
        '			'UPGRADE_WARNING: Couldn't resolve default property of object graph.YAxisMax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '			graph.YAxisMax = ymax
        '			'UPGRADE_WARNING: Couldn't resolve default property of object graph.YAxisTicks. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '			graph.YAxisTicks = 9
        '		End If

        'here: 
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.Graphdata. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.Graphdata = (multiplier * ydata(1) + DataShift)
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		For i = 2 To graph.NumPoints
        '			'UPGRADE_WARNING: Couldn't resolve default property of object graph.Graphdata. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '			graph.Graphdata = (multiplier * ydata(i) + DataShift)
        '		Next i
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.XPosData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.XPosData = (xdata(1))
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		For i = 2 To graph.NumPoints
        '			'UPGRADE_WARNING: Couldn't resolve default property of object graph.XPosData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '			graph.XPosData = (xdata(i))
        '		Next i

        '		'UPGRADE_ISSUE: Control method graph.DrawMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        '		graph.DrawMode = 2
        '		Exit Sub

        'graphhandler: 
        '		'UPGRADE_WARNING: Couldn't resolve default property of object graph.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '		graph.NumPoints = graph.NumPoints - 1
        '		GoTo here
        '		Exit Sub

    End Sub

    Public Sub Flame_Spread_Props_graph(ByRef room As Integer, ByRef id As Integer, ByRef high As Single, ByRef ConeDataFile As Object, ByRef emissivity As Double, ByRef ThermalInertia As Double, ByRef IgnitionTemp As Double, ByRef EffectiveHeatofCombustion As Double, ByRef HeatofGasification As Double, ByRef AreaUnderCurve As Double, ByRef FTP As Double, ByRef QCritical As Double, ByRef Nmax As Single)
        '========================================================================
        '   find flame spread properties
        '   4 April 1998
        '========================================================================
        Dim radflux(20) As Double
        Dim IgnitionTime(20) As Double
        Dim Peak(20) As Double
        Dim IgnPoints As Short
        Dim slope, IO, slopemax As Double
        Dim Ignition(20) As Double
        Dim Yintercept, RMax, R2 As Double
        Dim n As Single
        Dim i As Short
        Dim caption As String = ""

        AreaUnderCurve = 0
        IgnitionTemp = 0
        ThermalInertia = 0

        If ConeDataFile = "null.txt" Then
            Exit Sub
        End If

        'read data from cone
        Call Read_Cone_Data(room, id, high, ConeDataFile, radflux, IgnitionTime, IgnPoints, Peak, IgnitionTemp, ThermalInertia, HeatofGasification, AreaUnderCurve, QCritical)

        If frmQuintiere.optUseOneCurve.Checked = True Then
            'normalise the cone data
            Call Normalise_Data(id, room)
        End If
        If AreaUnderCurve <= 0 Then
            'find the area under the cone hrr curve if the user has not specified the value
            Call Integrate_HRR(room, id, 0, high, IO)
            AreaUnderCurve = IO
            If IO = 0 Then MsgBox("Problem integrating the area under the cone data curve.")
        End If

        Dim label As String
        If IgnPoints > 1 Then
            'determine ignition temperature and thermal inertia from a piloted ignition
            'correlation using the method of Marc Janssens

            RMax = 0
            Nmax = 0

            If IgnCorrelation = vbJanssens Then
                For n = 0.547 To 1 Step 0.001
                    For i = 1 To IgnPoints
                        If IgnitionTime(i) > 0 Then
                            Ignition(i) = (1 / IgnitionTime(i)) ^ n
                        End If
                    Next

                    'fit a linear regression line to the data, plotting vs external flux
                    Call RegressionL(radflux, Ignition, IgnPoints, Yintercept, slope, R2)

                    If R2 * R2 > RMax Then
                        RMax = R2 * R2
                        Nmax = n
                        slopemax = slope
                    End If
                Next
            Else 'FTP method
                For n = 1 To 0.5 Step -0.01
                    For i = 1 To IgnPoints
                        If IgnitionTime(i) > 0 Then
                            Ignition(i) = (1 / IgnitionTime(i)) ^ n
                        End If
                    Next

                    'fit a linear regression line to the data, plotting vs external flux
                    Call RegressionL(radflux, Ignition, IgnPoints, Yintercept, slope, R2)

                    If CDbl(VB6.Format(R2 * R2, "0.00")) > RMax Then
                        'If R2 * R2 > Format(RMax, "0.00") Then
                        RMax = CDbl(VB6.Format(R2 * R2, "0.00"))
                        Nmax = n
                        slopemax = slope
                    End If
                Next
                FTP = (1 / slopemax) ^ (1 / Nmax)
            End If

            For i = 1 To IgnPoints
                If IgnitionTime(i) > 0 Then
                    Ignition(i) = (1 / IgnitionTime(i)) ^ Nmax
                End If
            Next

            Call Ignition_Correlation(radflux, Ignition, IgnitionTime, IgnPoints, emissivity, ThermalInertia, IgnitionTemp, QCritical)

            'plot
            label = ""
            If id = 1 Then
                Call Graph_Ign_Data(label, caption, IgnPoints, "wall (1/t_ign)^n", radflux, Ignition, 0, 1, 4)
            ElseIf id = 2 Then
                Call Graph_Ign_Data(label, caption, IgnPoints, "ceiling (1/t_ign)^n", radflux, Ignition, 0, 1, 4)
            ElseIf id = 3 Then
                Call Graph_Ign_Data(label, caption, IgnPoints, "floor (1/t_ign)^n", radflux, Ignition, 0, 1, 4)
            End If


            'determine effective heat of gasification using Quintiere's method
            Call EHoG_Correlation(radflux, Peak, IgnPoints, EffectiveHeatofCombustion, HeatofGasification)

            Nmax = 1 / Nmax
            label = " (Heat of Gasification = " & VB6.Format(HeatofGasification, "0.0") & " kJ/g)"
            If id = 4 Then
                Call Graph_Ign_Data(label, caption, IgnPoints, "Peak wall HRR (kW/m²)", radflux, Peak, 0, 1, 4)
            ElseIf id = 5 Then
                Call Graph_Ign_Data(label, caption, IgnPoints, "Peak ceiling HRR (kW/m²)", radflux, Peak, 0, 1, 4)
            ElseIf id = 6 Then
                Call Graph_Ign_Data(label, caption, IgnPoints, "Peak floor HRR (kW/m²)", radflux, Peak, 0, 1, 4)
            End If

        End If
    End Sub
    Public Sub Flame_Spread_Props_graph2(ByRef room As Integer, ByRef id As Integer, ByRef high As Single, ByRef ConeDataFile As Object, ByRef emissivity As Double, ByRef ThermalInertia As Double, ByRef IgnitionTemp As Double, ByRef EffectiveHeatofCombustion As Double, ByRef HeatofGasification As Double, ByRef AreaUnderCurve As Double, ByRef FTP As Double, ByRef QCritical As Double, ByRef Nmax As Single)
        '========================================================================
        '   find flame spread properties
        '   Nov 2012
        '========================================================================
        Dim radflux(20) As Double
        Dim IgnitionTime(20) As Double
        Dim Peak(20) As Double
        Dim IgnPoints As Short
        Dim slope, IO, slopemax As Double
        Dim Ignition(20) As Double
        Dim Yintercept, RMax, R2 As Double
        Dim n As Single
        Dim i As Short
        Dim caption As String = ""

        AreaUnderCurve = 0
        IgnitionTemp = 0
        ThermalInertia = 0

        If ConeDataFile = "null.txt" Then
            Exit Sub
        End If

        'read data from cone
        Call Read_Cone_Data(room, id, high, ConeDataFile, radflux, IgnitionTime, IgnPoints, Peak, IgnitionTemp, ThermalInertia, HeatofGasification, AreaUnderCurve, QCritical)

        If frmQuintiere.optUseOneCurve.Checked = True Then
            'normalise the cone data
            Call Normalise_Data(id, room)
        End If
        If AreaUnderCurve <= 0 Then
            'find the area under the cone hrr curve if the user has not specified the value
            Call Integrate_HRR(room, id, 0, high, IO)
            AreaUnderCurve = IO
            If IO = 0 Then MsgBox("Problem integrating the area under the cone data curve.")
        End If

        Dim label As String
        If IgnPoints > 1 Then
            'determine ignition temperature and thermal inertia from a piloted ignition
            'correlation using the method of Marc Janssens

            RMax = 0
            Nmax = 0

            If IgnCorrelation = vbJanssens Then

            Else 'FTP method
                For n = 1 To 0.5 Step -0.01
                    For i = 1 To IgnPoints
                        If IgnitionTime(i) > 0 Then
                            Ignition(i) = 1 / (IgnitionTime(i) ^ (1 / n))
                        End If
                    Next

                    'fit a linear regression line to the data, plotting vs external flux
                    Call RegressionL(Ignition, radflux, IgnPoints, Yintercept, slope, R2)

                    If CDbl(VB6.Format(R2 * R2, "0.00")) > RMax Then

                        RMax = CDbl(VB6.Format(R2 * R2, "0.00"))
                        Nmax = n
                        slopemax = slope
                    End If
                Next
                'FTP = (1 / slopemax) ^ (1 / Nmax)
                FTP = (slopemax) ^ (Nmax)
            End If

            For i = 1 To IgnPoints
                If IgnitionTime(i) > 0 Then
                    'Ignition(i) = (1 / IgnitionTime(i)) ^ Nmax
                    Ignition(i) = 1 / (IgnitionTime(i) ^ (1 / Nmax))
                End If
            Next

            Call Ignition_Correlation(radflux, Ignition, IgnitionTime, IgnPoints, emissivity, ThermalInertia, IgnitionTemp, QCritical)

            'plot
            label = ""
            'If id = 1 Then
            '    Call Graph_Ign_Data(label, caption, IgnPoints, "wall (1/t_ign)^n", radflux, Ignition, 0, 1, 4)
            'ElseIf id = 2 Then
            '    Call Graph_Ign_Data(label, caption, IgnPoints, "ceiling (1/t_ign)^n", radflux, Ignition, 0, 1, 4)
            'ElseIf id = 3 Then
            '    Call Graph_Ign_Data(label, caption, IgnPoints, "floor (1/t_ign)^n", radflux, Ignition, 0, 1, 4)
            'End If


            'determine effective heat of gasification using Quintiere's method
            Call EHoG_Correlation(radflux, Peak, IgnPoints, EffectiveHeatofCombustion, HeatofGasification)

            Nmax = 1 / Nmax
            label = " (Heat of Gasification = " & VB6.Format(HeatofGasification, "0.0") & " kJ/g)"
            If id = 4 Then
                Call Graph_Ign_Data(label, caption, IgnPoints, "Peak wall HRR (kW/m²)", radflux, Peak, 0, 1, 4)
            ElseIf id = 5 Then
                Call Graph_Ign_Data(label, caption, IgnPoints, "Peak ceiling HRR (kW/m²)", radflux, Peak, 0, 1, 4)
            ElseIf id = 6 Then
                Call Graph_Ign_Data(label, caption, IgnPoints, "Peak floor HRR (kW/m²)", radflux, Peak, 0, 1, 4)
            End If

        End If
    End Sub

    Public Sub Ventflow_Log(ByRef FirstTime As Boolean, ByRef tim As Double, ByRef i As Integer, ByRef j As Integer, ByRef m As Integer, ByRef nslab As Integer, ByRef xmslab() As Double, ByRef dirs12() As Double, ByRef yvelev() As Double)
        '===========================================================================
        ' save vent flow data in log file
        '===========================================================================

        On Error GoTo more

        Dim count As Integer
        Dim logfile, Txt As String
        Dim F7dot2, E12dot4, F8dot4, F6dot2 As String
        Dim F7dot1, F7dot3, F10dot6, F12dot2 As String
        Dim F8dot2, I10 As String

        'logfile = UserAppDataFolder & gcs_folder_ext & "\data\ventlog.csv"
        logfile = RiskDataDirectory & "ventlog.csv"

        'define a format string
        I10 = "@@@@@@@@@@"


        'If tim - Int(tim) = 0 Then
        If FirstTime = True Then
            FileOpen(1, logfile, OpenMode.Output)
            WriteLine(1, DataFile)
            WriteLine(1, "time(s)," & "from-room," & "to-room," & "vent#," & "#slabs," & "ventflow(kg/s) for slabs 1 2 3 etc from 1st room to 2nd room")
            WriteLine(1, "," & "," & "," & "," & "," & "elevations (m) from floor")
            FirstTime = False
        Else
            'If tim > tlast Then
            FileOpen(1, logfile, OpenMode.Append)
            'Else
            '    Exit Sub
            'End If
        End If
        Txt = (VB6.Format(tim, "E12dot4") & "," & VB6.Format(i, I10) & "," & VB6.Format(j, I10) & "," & VB6.Format(m, I10) & "," & VB6.Format(nslab, I10)) & ","
        For count = 1 To nslab
            Txt = Txt & VB6.Format(dirs12(count) * xmslab(count), "#.000") & ","
        Next count
        WriteLine(1, Txt)
        Txt = "," & "," & "," & "," & ","
        For count = 1 To nslab + 1
            Txt = Txt & VB6.Format(yvelev(count), "#.000") & ","
        Next count
        WriteLine(1, Txt)
        FileClose(1)

        'End If
        Exit Sub



more:
        If Err.Number <> 32755 Then MsgBox(ErrorToString(Err.Number))
        Exit Sub

    End Sub


    Public Function JET5(ByVal room As Integer, ByVal tim As Double, ByVal layerz As Double, ByVal UT As Double, ByVal LT As Double, ByVal LinkTemp As Double, ByRef CJTemp As Double, ByRef CJMax As Double) As Double
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

        On Error GoTo errhandler

        beta = 0.9555 'velocity to temperature ratio of gaussian profile half widths
        CT = 9.115
        alpha = 0.44

        If room = fireroom Then
            'get total heat release rate for object 1 at this time (q in kW)
            Q = Get_HRR(1, tim, 0, 0, 0, 0, 0)

            'use method of reflection to account for fires located in corner or against wall
            If FireLocation(1) = 2 Then Q = 4 * Q 'corner fire
            If FireLocation(1) = 1 Then Q = 2 * Q 'wall fire

            Zi1 = layerz - FireHeight(1) 'Fireheight(1) = the height of the base of the fire above the floor
            H1 = RoomHeight(room) - FireHeight(1) 'uses height of object 1
            epsilon = UT / LT
            If epsilon < 1 Then Exit Function ' shouldn't happen!

            If Zi1 <= 0.01 Then Zi1 = 0.001 'in the limit - calculate based on the layer being just above the fire height

            'calculate non-dimensional Qi1
            If Zi1 > 0 Then
                Qi1 = (1 - ObjectRLF(1)) * Q / ((Atm_Pressure) / (Gas_Constant / MW_air) * SpecificHeat_air * Sqrt(G) * Zi1 ^ (5 / 2))
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

            If RadialDistance(room) > ro Then
                delCeilingJetTemp = C / (RadialDistance(room) ^ gamma_2)
            Else
                delCeilingJetTemp = C / (ro ^ gamma_2)
            End If

            If RadialDistance(room) > 0.15 * H1 Then
                velocity = 0.195 * ((1 - ObjectRLF(1)) * Q) ^ (1 / 3) * Sqrt(H1) / (RadialDistance(room) ^ (5 / 6))
            Else
                velocity = 0.96 * ((1 - ObjectRLF(1)) * Q / H1) ^ (1 / 3)
            End If
            velocity = velocity * (1 - 0.25 * (1 - Exp(-yl / yj)))

            CJMax = delCeilingJetTemp + LT 'the maximum ceiling jet temperature at specified radial distance in K

            'variation with depth of link beneath ceiling - only the region from the ceiling surface to the
            'depth at which the maximum temperature occurs is considered, ceiling jet at greater depths are
            'assumed to be at the maximum temperature
            'LAVENT NPFA 204 Appendix B
            delta = 0.1 * H1 * (RadialDistance(room) / H1) ^ 0.9

            'adjustment for location of link below ceiling
            If RadialDistance(room) / H1 >= 0.2 Then
                If SprinkDistance(room) >= 0.23 * delta Then
                    ratio_v = COSH((0.23 / 0.77) * ACOSH(Sqrt(2)) * (SprinkDistance(room) / 0.23 / delta - 1)) ^ (-2)
                    ratio_t = ratio_v
                Else
                    ratio_v = 8 / 7 * (SprinkDistance(room) / 0.23 / delta) ^ (1 / 7) * (1 - SprinkDistance(room) / 0.23 / delta / 8)
                    If (delCeilingJetTemp + LT - UT) <> 0 Then
                        psi = (CeilingTemp(room, stepcount) - UT) / (delCeilingJetTemp + LT - UT)
                    Else
                        psi = 0
                    End If
                    ratio_t = psi + 2 * ((1 - psi) * (SprinkDistance(room) / 0.23 / delta)) - ((1 - psi) * (SprinkDistance(room) / 0.23 / delta) ^ 2)
                End If
            Else
                ratio_v = 1
                ratio_t = 1
            End If

            velocity = velocity * ratio_v 'velocity in the ceiling jet at the position of the sprinkler

            delCeilingJetTemp = ratio_t * (delCeilingJetTemp + LT - UT) 'this is the rise (above upper layer) in the ceiling jet temp at the location of sensor link
            delCeilingJetTemp = delCeilingJetTemp + UT - LT 'excess above lower temp

            If (SprinkDistance(room) > yj) And (RadialDistance(room) / H1 >= 0.2) Then 'not in ceiling jet
                If RoomHeight(room) - SprinkDistance(room) < layerz Then 'not in the ceiling jet, but in lower layer
                    delCeilingJetTemp = 0
                End If
            End If

            CJTemp = LT + delCeilingJetTemp 'the maximum ceiling jet temperature at specified radial distance and distance below ceiling in K

            If CJMax < UT And 0.23 * delta < yl Then CJMax = UT 'if max location is in the upper layer
            If CJMax < LT And 0.23 * delta > yl Then CJMax = LT 'if max location is in the lower layer
            If CJTemp < UT And SprinkDistance(room) < yl And SprinkDistance(room) > 0.23 * delta Then CJTemp = UT 'if max location is in the upper layer
            If CJTemp < LT And SprinkDistance(room) > yl And SprinkDistance(room) > 0.23 * delta Then CJTemp = LT 'if max location is in the lower layer

            'If velocity > 0 Then
            '    If DetectorType(room) > 1 Then
            '        JET5 = Sqrt(velocity) / RTI(room) * (delCeilingJetTemp + LT - LinkTemp - cfactor(room) / Sqrt(velocity) * (LinkTemp - InteriorTemp))
            '    Else
            '        JET5 = Sqrt(velocity) / RTI(room) * (delCeilingJetTemp + LT - LinkTemp)
            '    End If
            'Else
            '    JET5 = 0
            'End If

            JET5 = 0 'we don't solve for the link temp using this routine anymore

            Exit Function

        Else 'not the fireroom

        End If
        'Next SprinkDistance
        Exit Function

errhandler:
        If Err.Number = 6 Then
            Err.Clear()
            Resume Next
        Else
            MsgBox(ErrorToString(Err.Number) & " in subroutine JET5")
        End If

    End Function


    Public Function Mass_Plume_nextroom(ByVal tim As Double, ByVal Z1 As Double, ByVal qmax As Double, ByVal UT As Double, ByVal LT As Double) As Double
        '*  ===================================================================
        '*  This function returns the values of the mass flow of air entrained
        '*  from the lower layer into the fire plume.
        '*
        '*  Arguments passed to the function are:
        '*      tim = time (sec)
        '*      Z1 = layer height above the floor (m)
        '*      QMax = max heat release that than be supported by air supply
        '*
        '*  Functions or subprocedures called:
        '*      Get_HRR
        '*      MassLoss_Object
        '*      Fire_Diameter
        '*      Virtual_Source
        '*  ===================================================================

        'this routine only used when floor is burning. 

        Dim Region, zo, diameter As Double
        Dim L As Single
        Dim i As Integer
        Dim Q As Single
        Dim MLoss, total As Double
        Dim ZF, massplume As Double
        'Static post, rcnone As Boolean
        Static post As Boolean

        If gb_first_time_vent = True Then 'If tim = 0 Then
            post = frmOptions1.optPostFlashover.Checked
            RCNone = frmOptions1.optRCNone.Checked
        End If

        'z1 cannot be negative
        If Z1 <= 0 Then
            Mass_Plume_nextroom = 0
            Exit Function
        End If

        total = 0
        Q = qmax

        'find height of layer above the base of the fire
        ZF = Z1
        If ZF < 0 Then ZF = 0

        'fire in centre of room
        Select Case PlumeModel

            Case "General"
                'general correlation from strong plume theory for buoyant plume
                'and Delichatsios's correlations for the flaming regions

                If ObjectMLUA(2, 1) > 0 Then
                    'hrrua
                    diameter = 2 * Sqrt((Q / ObjectMLUA(2, 1)) / PI)
                Else

                    'get mass loss rate for object
                    Dim mlua As Single = ObjectMLUA(0, 1) * (Target(fireroom, stepcount - 1) - Target(fireroom, 1)) + ObjectMLUA(1, 1)
                    If ObjectMLUA(0, 1) > 0 Or ObjectMLUA(1, 1) > 0 Then
                    Else
                        'mlua = MLUnitArea 'not specified for object
                        mlua = ObjectMLUA(2, 1) / EnergyYield(1) / 1000 'uses  object 1
                    End If

                    'find fire diameter for each object
                    diameter = Fire_Diameter(MLoss, mlua)

                End If

                'call function to get virtual source
                zo = Virtual_Source(Q, diameter)

                'call function to get flame height
                L = Flame_Height(diameter, Q) 'for a fire in centre of room

                If Z1 - FireHeight(i) - zo >= L Then
                    'buoyant plume
                    'massplume = EntrainmentCoefficient * (1 - RadiantLossFraction) ^ (1 / 3) * Q ^ (1 / 3) * (ZF - zo) ^ (5 / 3)
                    massplume = EntrainmentCoefficient * (1 - NewRadiantLossFraction(1)) ^ (1 / 3) * Q ^ (1 / 3) * (ZF - zo) ^ (5 / 3)
                ElseIf ZF - zo >= 5.16 * diameter Then
                    'flame
                    massplume = 0.18 * ((ZF - zo)) ^ (5 / 2)
                ElseIf ZF - zo >= diameter Then
                    massplume = 0.93 * ((ZF - zo) / diameter) ^ (3 / 2) * diameter ^ (5 / 2)
                ElseIf ZF - zo > 0 Then
                    massplume = 0.86 * ((ZF - zo) / diameter) ^ (1 / 2) * diameter ^ (5 / 2)
                Else
                    massplume = 0
                End If

                If ZF <= 0 Then
                    massplume = 0
                End If

            Case "McCaffrey"
                'McCaffrey's correlation

                If Q > 0 Then
                    Region = (ZF) / (Q ^ (2 / 5))
                Else
                    massplume = 0
                End If

                If ZF >= 0 And Q > 0 Then
                    If Region < 0.08 Then
                        'flaming region
                        massplume = 0.011 * Q * (Region) ^ 0.566
                    ElseIf Region >= 0.2 Then
                        'buoyant plume
                        massplume = 0.124 * Q * (Region) ^ 1.895
                    Else
                        'intermittent region
                        massplume = 0.026 * Q * (Region) ^ 0.909
                    End If
                Else
                    massplume = 0
                End If
        End Select

        total = total + massplume

        If UT > LT Then
            If (1 - NewRadiantLossFraction(1)) * Q / (SpecificHeat_air * (UT - LT)) < total Then
                total = (1 - NewRadiantLossFraction(1)) * Q / (SpecificHeat_air * (UT - LT))
            End If
        ElseIf UT = LT Then
            total = 0
        End If

        Mass_Plume_nextroom = total

    End Function

    Public Sub ISO_energy(ByRef energy As Double, ByRef frr As Double)
        '===========================================
        ' determine what FRR period gives a specified area under the energy curve
        ' 10/10/2002
        ' send energy, return equiv FRR time
        '==============================================
        Dim T, isotemp, E, Etot As Double

        T = 0
        E = 0
        Etot = 0

        Do While energy > Etot And T <= 21600 '6 hour upper limit
            isotemp = 345 * Log(8 * (T / 60) + 1) / Log(10.0#) + InteriorTemp 'K
            E = StefanBoltzmann * isotemp ^ 4 'W/m2
            Etot = Etot + E / 1000 'kJ/m2
            T = T + 1 'increment by 1 sec
        Loop

        frr = (T - 1) / 60 'min in furnace test

    End Sub

    Public Sub View_ventlog(ByRef filename As Object)
        '*  ====================================================================
        '*  View the vent flow data in a rich text box
        '*  5/11/2002
        '*  ====================================================================

        Dim j, i, k As Short
        Dim dummy As String = ""
        Dim temp As String = ""
        Dim Line1 As String
        On Error GoTo errorhandler

        'UPGRADE_WARNING: Couldn't resolve default property of object filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileOpen(1, filename, OpenMode.Input)
        FileOpen(2, temp, OpenMode.Output)
        'read in preamble
        dummy = LineInput(1)

        dummy = LineInput(1)

        dummy = LineInput(1)


        Print(2, "Time (sec)", TAB(20), "From Room", TAB(20), "To Room", TAB(20), "Vent#")

        Input(1, dummy)
        Input(1, dummy)
        Input(1, dummy)
        Input(1, dummy)


        FileClose(2)
        FileClose(1)

        Exit Sub

        'UPGRADE_WARNING: Couldn't resolve default property of object filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileOpen(1, filename, OpenMode.Output)
        PrintLine(1)
        PrintLine(1, VB6.Format(Now, "dddd,mmmm dd,yyyy,hh:mm AM/PM"))
        PrintLine(1, "Input Filename : " & DataFile)
        PrintLine(1)
        PrintLine(1, "BRANZFIRE Multi-Compartment Fire Model (Ver " & Version & ")")

        PrintLine(1)
        PrintLine(1, Description)
        PrintLine(1)
        PrintLine(1, "====================================================================")
        PrintLine(1, "Wall Vent Flow Data")
        PrintLine(1, "====================================================================")


        For j = 1 To NumberRooms
            For i = 1 To NumberRooms + 1
                If NumberVents(j, i) > 0 Then
                    If j < i Then
                        For k = 1 To NumberVents(j, i)
                            If i = NumberRooms + 1 Then
                                PrintLine(1, "From room " & j & " to outside" & ", Vent No " & CStr(k))
                            Else
                                PrintLine(1, "From room " & j & " to " & i & ", Vent No " & CStr(k))
                            End If
                            PrintLine(1, TAB(20), "Vent Width (m) =", TAB(60), VB6.Format(VentWidth(j, i, k), "0.000"))
                            PrintLine(1, TAB(20), "Vent Height (m) =", TAB(60), VB6.Format(VentHeight(j, i, k), "0.000"))
                            PrintLine(1, TAB(20), "Vent Sill Height (m) =", TAB(60), VB6.Format(VentSillHeight(j, i, k), "0.000"))
                            PrintLine(1, TAB(20), "Vent Soffit Height (m) =", TAB(60), VB6.Format(soffitheight(j, i, k), "0.000"))
                            PrintLine(1, TAB(20), "Opening Time (sec) =", TAB(60), VB6.Format(VentOpenTime(j, i, k), "0"))
                            PrintLine(1, TAB(20), "Closing Time (sec) =", TAB(60), VB6.Format(VentCloseTime(j, i, k), "0"))
                            If AutoBreakGlass(j, i, k) = True Then
                                PrintLine(1, TAB(20), "Glass fracture is modelled for this vent")
                                PrintLine(1, TAB(20), "Glass Thickness (mm) =", TAB(60), VB6.Format(GLASSthickness(j, i, k), "0.0"))
                                PrintLine(1, TAB(20), "Glass Fracture to Fallout Time (sec) =", TAB(60), VB6.Format(GLASSFalloutTime(j, i, k), "0"))
                                PrintLine(1, TAB(20), "Glass Shading Depth (mm) =", TAB(60), VB6.Format(GLASSshading(j, i, k), "0.0"))
                                PrintLine(1, TAB(20), "Glass Fracture Stress (MPa) =", TAB(60), VB6.Format(GLASSbreakingstress(j, i, k), "0"))
                                PrintLine(1, TAB(20), "Glass Expansion Coefficient (/C) =", TAB(60), VB6.Format(GLASSexpansion(j, i, k)))
                                PrintLine(1, TAB(20), "Glass Conductivity (W/mK) =", TAB(60), VB6.Format(GLASSconductivity(j, i, k), "0.00"))
                                PrintLine(1, TAB(20), "Glass Diffusivity (m2/s) =", TAB(60), VB6.Format(GLASSalpha(j, i, k)))
                                PrintLine(1, TAB(20), "Glass Modulus (W/mK) =", TAB(60), VB6.Format(GlassYoungsModulus(j, i, k), "0"))
                                If GlassFlameFlux(j, i, k) = False Then
                                    PrintLine(1, TAB(20), "Glass is heated by gas layers only.")
                                Else
                                    PrintLine(1, TAB(20), "Glass is heated by gas layer and flame radiation.")
                                    PrintLine(1, TAB(20), "Glass to Flame Distance (m) =", TAB(60), VB6.Format(GLASSdistance(j, i, k), "0.000"))
                                End If
                            End If
                            PrintLine(1)
                        Next k
                    End If
                End If
            Next i
        Next j
        PrintLine(1, "====================================================================")
        PrintLine(1, "Description of Ceiling/Floor Vents")
        PrintLine(1, "====================================================================")
        For j = 1 To NumberRooms + 1
            For i = 1 To NumberRooms + 1
                If NumberCVents(j, i) > 0 Then
                    For k = 1 To NumberCVents(j, i)
                        If j <> NumberRooms + 1 Then
                            If i = NumberRooms + 1 Then
                                PrintLine(1, "Upper room " & j & " to lower outside" & ", Vent No " & CStr(k))
                            Else
                                PrintLine(1, "Upper room " & j & " to lower room " & i & ", Vent No " & CStr(k))
                            End If
                        Else
                            If i = NumberRooms + 1 Then
                            Else
                                PrintLine(1, "Upper outside " & " to lower room" & i & ", Vent No " & CStr(k))
                            End If
                        End If
                        PrintLine(1, TAB(20), "Vent Area (m2) =", TAB(60), VB6.Format(CVentArea(j, i, k), "0.00"))
                        PrintLine(1, TAB(20), "Opening Time (sec) =", TAB(60), VB6.Format(CVentOpenTime(j, i, k), "0"))
                        PrintLine(1, TAB(20), "Closing Time (sec) =", TAB(60), VB6.Format(CVentCloseTime(j, i, k), "0"))
                        PrintLine(1)
                    Next k
                End If
            Next i
        Next j


        PrintLine(1)
        FileClose(1)
        Exit Sub

errorhandler:
        'Err.Clear
        If Err.Number = 55 Then
            FileClose(1)
            Resume
        End If
        Exit Sub

    End Sub

    Public Sub Ventflow_log2(ByRef FirstTime As Boolean, ByRef tim As Double, ByRef i As Integer, ByRef j As Integer, ByRef m As Integer, ByRef nslab As Integer, ByRef xmslab() As Double, ByRef dirs12() As Double, ByRef yvelev() As Double, ByRef vententrain(,,) As Double)
        '===========================================================================
        ' save vent flow data in log file
        ' 28 January 2008
        '===========================================================================

        On Error GoTo more

        Dim count As Integer
        Dim logfile, Txt As String
        Dim F7dot2, E12dot4, F8dot4, F6dot2 As String
        Dim F7dot1, F7dot3, F10dot6, F12dot2 As String
        Dim I10, F8dot2, fname As String
        Dim strTempPath As String

        'define a format string
        I10 = "@@@@@@@@@@"

        If FirstTime = False Then
            If (tim Mod DisplayInterval = 0) Then
            Else
                Exit Sub
            End If
            FileOpen(1, WVentlogfile, OpenMode.Append)
        Else
            fname = CStr(Int(Rnd() * 100000))
            If Val(fname) < 1 Then Exit Sub
            'strTempPath = getTempPathName() 'get name of windows temp folder
            strTempPath = ProjectDirectory 'get name of windows temp folder

            'WVentlogfile = strTempPath & "BF" & Str(CDbl(fname)) & ".tmp"
            WVentlogfile = strTempPath & "wallventflows" & ".txt"
            FileOpen(1, WVentlogfile, OpenMode.Output)
            PrintLine(1, DataFile)
            PrintLine(1, "Created: ", VB6.Format(Now, "dddd,mmmm dd,yyyy,hh:mm AM/PM"))
            PrintLine(1)
            PrintLine(1, "==================================")
            PrintLine(1, " Wall Vent Flows")
            PrintLine(1, "==================================")
            PrintLine(1)
            PrintLine(1, "Time(s)", TAB(10), "from-room", TAB(20), "to-room", TAB(30), "vent#", TAB(40), "#slabs", TAB(50), "elevation (m)", TAB(68), "ventflow(kg/s)", TAB(84), "entrain (kg/s)")
            FirstTime = False
        End If

        Print(1, VB6.Format(tim), TAB(10), VB6.Format(i), TAB(20), VB6.Format(j), TAB(30), VB6.Format(m), TAB(40), VB6.Format(nslab))
        For count = nslab To 1 Step -1
            If count = nslab And j <= NumberRooms And i <= NumberRooms Then
                If dirs12(count) >= 0 Then
                    PrintLine(1, TAB(50), VB6.Format(yvelev(count), "0.000") & " to " & VB6.Format(yvelev(count + 1), "0.000"), TAB(68), VB6.Format(dirs12(count) * xmslab(count), "0.0000"), TAB(84), VB6.Format(vententrain(i, j, m), "0.0000"))
                Else
                    PrintLine(1, TAB(50), VB6.Format(yvelev(count), "0.000") & " to " & VB6.Format(yvelev(count + 1), "0.000"), TAB(68), VB6.Format(dirs12(count) * xmslab(count), "0.0000"), TAB(84), VB6.Format(vententrain(j, i, m), "0.0000"))
                End If
            Else
                PrintLine(1, TAB(50), VB6.Format(yvelev(count), "0.000") & " to " & VB6.Format(yvelev(count + 1), "0.000"), TAB(68), VB6.Format(dirs12(count) * xmslab(count), "0.0000"))
            End If
        Next count
        PrintLine(1)

        FileClose(1)

        Exit Sub

more:
        If Err.Number <> 32755 Then MsgBox(ErrorToString(Err.Number))
        Exit Sub

    End Sub
    Public Sub Ventflow_log_csv(ByRef FirstTime3 As Boolean, ByRef tim As Double, ByRef i As Integer, ByRef j As Integer, ByRef m As Integer, ByRef nslab As Integer, ByRef xmslab() As Double, ByRef dirs12() As Double, ByRef yvelev() As Double, ByRef vententrain(,,) As Double)
        '===========================================================================
        ' save vent flow data in log file
        '===========================================================================

        On Error GoTo more

        Dim count As Integer
        Dim Txt As String
        Dim F7dot2, E12dot4, F8dot4, F6dot2 As String
        Dim F7dot1, F7dot3, F10dot6, F12dot2 As String
        Dim I10, F8dot2, fname As String
        Dim strTempPath As String

        'define a format string
        I10 = "@@@@@@@@@@"

        If FirstTime3 = False Then
            If (tim Mod DisplayInterval = 0) Then
            Else
                Exit Sub
            End If
            FileOpen(2, WVentlogfile2, OpenMode.Append)
        Else
            fname = CStr(Int(Rnd() * 100000))
            If Val(fname) < 1 Then Exit Sub
            strTempPath = ProjectDirectory 'get name of windows temp folder

            WVentlogfile2 = strTempPath & "wallventflows" & ".csv"
            FileOpen(2, WVentlogfile2, OpenMode.Output)

            PrintLine(2, "Time(s)", TAB(10), "from-room", TAB(20), "to-room", TAB(30), "vent#", TAB(40), "#slabs", TAB(50), "elevation (m)", TAB(68), "ventflow(kg/s)", TAB(84), "entrain (kg/s)")
            FirstTime3 = False
        End If

        Dim stuff As String = CStr(tim) & ","

        For count = 1 To nslab

            stuff = stuff & yvelev(count) & "," & yvelev(count + 1) & "," & dirs12(count) * xmslab(count)

            stuff = stuff & ","
        Next count


        PrintLine(2, stuff)

        FileClose(2)

        Exit Sub

more:
        If Err.Number <> 32755 Then MsgBox(ErrorToString(Err.Number))
        Exit Sub

    End Sub

    Public Sub Ventflow_log3(ByRef FirstTime2 As Boolean, ByRef tim As Double, ByRef i As Integer, ByRef j As Integer, ByRef m As Integer, ByRef r1u As Double, ByRef r1l As Double, ByRef r2u As Double, ByRef r2l As Double, ByRef XMVent() As Double)
        '===========================================================================
        ' save ceiling vent flow data in log file
        '===========================================================================

        On Error GoTo more

        Dim count As Integer
        Dim Txt As String
        Dim F7dot2, E12dot4, F8dot4, F6dot2 As String
        Dim F7dot1, F7dot3, F10dot6, F12dot2 As String
        Dim I10, F8dot2, fname As String
        'logfile = App.Path + "\data\cventlog.txt"
        Dim strTempPath As String
        'define a format string
        I10 = "@@@@@@@@@@"


        If FirstTime2 = True Then
            fname = CStr(Int(Rnd() * 100000))
            If Val(fname) < 1 Then Exit Sub
            strTempPath = getTempPathName() 'get name of windows temp folder
            CVentlogfile = strTempPath & "BF" & Str(CDbl(fname)) & ".tmp"
            FileOpen(1, CVentlogfile, OpenMode.Output)
            PrintLine(1, DataFile)
            PrintLine(1, "Created: ", VB6.Format(Now, "dddd,mmmm dd,yyyy,hh:mm AM/PM"))
            PrintLine(1)
            PrintLine(1, "==================================")
            PrintLine(1, " Ceiling Vent Flows")
            PrintLine(1, "==================================")
            PrintLine(1, "UR-UL = upper room, upper layer")
            PrintLine(1, "UR-LL = upper room, lower layer")
            PrintLine(1, "LR-UL = lower room, upper layer")
            PrintLine(1, "LR-LL = lower room, lower layer")
            PrintLine(1, "Vent flow (kg/s) +ve = deposited into the layer")
            PrintLine(1, "Vent flow (kg/s) -ve = removed from the layer")
            PrintLine(1, "Flow->U is the vent flow component to the upper space")
            PrintLine(1, "Flow->L is the vent flow component to the lower space")
            PrintLine(1)
            PrintLine(1, "Time(s)", TAB(10), "U room", TAB(20), "L room", TAB(30), "vent#", TAB(40), "UR-UL", TAB(50), "UR-LL", TAB(60), "LR-UL", TAB(70), "LR-LL", TAB(80), "Flow->U", TAB(90), "Flow->L")
            FirstTime2 = False
        Else
            'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            If (tim Mod DisplayInterval = 0) Then
            Else
                Exit Sub
            End If
            FileOpen(1, CVentlogfile, OpenMode.Append)
        End If
        'Txt = (Format(tim, E12dot4) & "," & Format(i, I10) & "," & Format(j, I10) & "," & Format(m, I10) & "," & Format(nslab, I10)) & ","
        Print(1, VB6.Format(tim), TAB(10), VB6.Format(i), TAB(20), VB6.Format(j), TAB(30), VB6.Format(m), TAB(40), VB6.Format(r1u, "0.000"), TAB(50), VB6.Format(r1l, "0.000"), TAB(60), VB6.Format(r2u, "0.000"), TAB(70), VB6.Format(r2l, "0.000"), TAB(80), VB6.Format(XMVent(1), "0.000"), TAB(90), VB6.Format(XMVent(2), "0.000"))
        PrintLine(1)

        FileClose(1)

        Exit Sub

more:
        If Err.Number <> 32755 Then MsgBox(ErrorToString(Err.Number))
        Exit Sub


    End Sub

    Public Sub Graph_GER(ByRef thisroom As Short, ByRef Title As String, ByRef datatobeplotted() As Double, ByRef DataShift As Double, ByRef DataMultiplier As Double, ByRef GraphSty As Short, ByRef MaxYValue As Double)
        '*  ====================================================================
        '*  This function takes data for a variable from a two-dimensional array
        '*  and displays it in a graph
        '*
        '*  ====================================================================
        Dim j As Integer
        Dim room As Integer
        Dim ydata(0 To NumberTimeSteps) As Double

        'if no data exists
        If NumberTimeSteps < 1 Then
            MsgBox("There is no data to plot, please run the simulation first.", vbExclamation)
            Exit Sub
        End If
        Try

            frmPlot.Chart1.Series.Clear()
            room = 1
            For room = 1 To NumberRooms

                frmPlot.Chart1.Series.Add("Room " & room)

                frmPlot.Chart1.Series("Room " & room).ChartType = SeriesChartType.FastLine

                For j = 1 To NumberTimeSteps
                    ydata(j) = datatobeplotted(j) * DataMultiplier + DataShift 'data to be plotted
                    frmPlot.Chart1.Series("Room " & room).Points.AddXY(tim(j, 1) / timeunit, ydata(j))
                Next

            Next room
            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            'frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            If timeunit = 60 Then
                frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (min)"
            Else
                frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            End If
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            'frmPlot.Chart1.Titles("Title1").Text = Title
            ' Disable X axis margin

            frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in Graph_GER")
        End Try


    End Sub

    Public Sub Chemistry()
        '======================================
        ' calculating water vapor yield using
        ' chemical composition of fuel
        ' c wade 20/3/2003
        '======================================

        Dim h As Short
        Dim nSoot, nCO2, nHCN, nH2O, nCO As Double


        nCO = (COYield * MolecularWeightFuel / MolecularWeightCO)

        For h = 1 To NumberObjects
            nCO2 = (CO2Yield(h) * MolecularWeightFuel / MolecularWeightCO2)
            nSoot = (SootYield(h) * MolecularWeightFuel / MolecularWeightC)
            'moles of HCN in product
            If nN > 0 Then
                nHCN = nC - nCO2 - nCO - nSoot
            Else
                nHCN = 0
            End If
            If nHCN > 0 Then
                If frmOptions1.chkHCNcalc.Checked = CHECKED Then
                    HCNYield(h) = nHCN * MolecularWeightHCN / MolecularWeightFuel
                Else
                    HCNYield(h) = HCNuserYield(h)
                End If
                nH2O = (nH - nHCN) / 2
            Else
                'If frmOptions1.chkHCNcalc.Checked = CHECKED Then
                '    HCNYield(h) = 0
                'Else
                '    HCNYield(h) = HCNuserYield(h)
                'End If
                HCNYield(h) = 0
                nH2O = (nH) / 2
            End If
            'moles of water
            If nH2O > 0 Then
                WaterVaporYield(h) = nH2O * MolecularWeightH2O / MolecularWeightFuel
            Else
                WaterVaporYield(h) = 0
            End If
            'moles of air required
            'nair = (2 * nCO2 + nH2O + nCO - nO) / 2
        Next h

    End Sub

    Public Function Wall_flow(ByVal density As Double, ByVal X As Double, ByVal room As Integer, ByVal gastemp As Double, ByVal surfacetemp As Double, ByVal layer As Double, ByVal depth As Double) As Double
        '=================================================
        ' natural convection wall flow
        ' Jaluria 1-116 SFPE Handbook first ed.
        ' C Wade 20/3/2003
        '=================================================

        Dim L, Pr, GR, V, U As Double
        Dim P, DEL As Double
        Dim m, vent As Short

        On Error GoTo errorhandler

        If Abs(gastemp - surfacetemp) <= 0 Then Exit Function

        'determine perimeter dimension to use in the calculation
        For m = 1 To NumberRooms + 1
            P = 2 * (RoomWidth(room) + RoomLength(room))
            For vent = 1 To NumberVents(room, m)
                'If layer > VentSillHeight(room, m, vent) + VentHeight(room, m, vent) Then
                If VentOpenTime(room, m, vent) < X Then 'if vent open
                    If VentCloseTime(room, m, vent) > VentOpenTime(room, m, vent) And VentCloseTime(room, m, vent) > X Then
                        P = P - VentWidth(room, m, vent)
                    ElseIf VentCloseTime(room, m, vent) = VentOpenTime(room, m, vent) Then
                        P = P - VentWidth(room, m, vent)
                    Else
                        P = P
                    End If
                End If
                'End If
            Next vent
        Next m

        Pr = 0.69 'prandtl no

        'kinematic viscocity
        V = 7.18 * 10 ^ (-10) * ((gastemp + surfacetemp) / 2) ^ (7 / 4)

        L = depth

        GR = G * L ^ 3 * Abs(gastemp - surfacetemp) / (V ^ 2 * gastemp) 'beta=1/gastemp

        If GR > 5 * 10 ^ 9 Then
            'turbulent
            U = 1.185 * V / L * Sqrt(GR) * (1 + 0.494 * Pr ^ (2 / 3)) ^ (-1 / 2)
            DEL = 0.565 * L * GR ^ (-0.1) * Pr ^ (-8 / 15) * (1 + 0.494 * Pr ^ (2 / 3)) ^ (1 / 10)
            Wall_flow = 0.1463 * P * density * DEL * U 'kg/s
        Else
            'laminar
            U = 5.17 * V * (Pr + 20 / 21) ^ (-0.5) * Sqrt(G / gastemp / V ^ 2 * Abs(gastemp - surfacetemp)) * Sqrt(depth)
            DEL = 3.93 * Pr ^ (-0.5) * (Pr + 20 / 21) ^ 0.25 * (G / gastemp / V ^ 2 * Abs(gastemp - surfacetemp)) ^ (-1 / 4) * depth ^ 0.25
            Wall_flow = P * density * DEL * U / 12 'kg/s
        End If

        Exit Function

errorhandler:

        MsgBox(ErrorToString() & " in Module Main.bas Subroutine = wall_flow")
        Wall_flow = 0
        flagstop = 1
    End Function

    Public Function Wall_flow_momentum(ByVal density As Double, ByVal X As Double, ByVal room As Integer, ByVal gastemp As Double, ByVal surfacetemp As Double, ByVal layer As Double, ByVal depth As Double) As Double
        '=================================================
        ' natural convection wall flow
        ' Jaluria 1-116 SFPE Handbook first ed.
        ' C Wade 20/3/2003
        '=================================================

        Dim L, Pr, GR, V, U As Double
        Dim P, DEL As Double
        Dim m, vent As Short

        On Error GoTo errorhandler

        If Abs(gastemp - surfacetemp) = 0 Then Exit Function

        For m = 1 To NumberRooms + 1
            P = 2 * (RoomWidth(room) + RoomLength(room))
            For vent = 1 To NumberVents(room, m)
                If layer > VentSillHeight(room, m, vent) + VentHeight(room, m, vent) Then
                    If VentOpenTime(room, m, vent) < X Then
                        If VentCloseTime(room, m, vent) > VentOpenTime(room, m, vent) And VentCloseTime(room, m, vent) > X Then
                            P = P - VentWidth(room, m, vent)
                        ElseIf VentCloseTime(room, m, vent) = VentOpenTime(room, m, vent) Then
                            P = P - VentWidth(room, m, vent)
                        End If
                    End If
                End If
            Next vent
        Next m

        Pr = 0.69 'prandtl no

        'kinematic viscocity
        V = 7.18 * 10 ^ (-10) * ((gastemp + surfacetemp) / 2) ^ (7 / 4)

        L = depth

        GR = G * L ^ 3 * Abs(gastemp - surfacetemp) / (V ^ 2 * gastemp) 'beta=1/gastemp

        If GR > 5 * 10 ^ 9 Then
            'turbulent
            U = 1.185 * V / L * Sqrt(GR) * (1 + 0.494 * Pr ^ (2 / 3)) ^ (-1 / 2)
            DEL = 0.565 * L * GR ^ (-0.1) * Pr ^ (-8 / 15) * (1 + 0.494 * Pr ^ (2 / 3)) ^ (1 / 10)
            Wall_flow_momentum = 0.0523 * P * density * DEL * U ^ 2
        Else
            'laminar
            U = 5.17 * V * (Pr + 20 / 21) ^ (-0.5) * Sqrt(G / gastemp / V ^ 2 * Abs(gastemp - surfacetemp)) * Sqrt(depth)
            DEL = 3.93 * Pr ^ (-0.5) * (Pr + 20 / 21) ^ 0.25 * (G / gastemp / V ^ 2 * Abs(gastemp - surfacetemp)) ^ (-1 / 4) * depth ^ 0.25
            Wall_flow_momentum = P * density * DEL * U ^ 2 / 105
        End If
        Exit Function

errorhandler:

        MsgBox(ErrorToString() & " in Module Main.bas Subroutine = wall_flow_momentum")
        'UPGRADE_WARNING: Couldn't resolve default property of object wall_flow_momentum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Wall_flow_momentum = 0


    End Function

    Public Function Get_CO_yield(ByVal GER As Double, ByVal comode As Boolean, ByVal fo As Boolean) As Single

        'Dim dummy As Double
        'Dim total1, total, total2 As Single
        'Dim i As Short

        If comode = True Then 'manual entry from user
            'If GER < 1 Then
            If fo = False Then

                Get_CO_yield = preCO
            Else
                Get_CO_yield = postCO
            End If
        Else 'default
            If GER < 0.6 Then
                Get_CO_yield = preCO
            ElseIf GER > 1.1 Then
                Get_CO_yield = 0.2 'mulholland / postflashover
            Else
                Get_CO_yield = preCO + (GER - 0.6) * (0.2 - preCO) / (1.1 - 0.6)
            End If
        End If

    End Function

    Public Sub Device_Defaults(ByRef i As Short)
        'assign default values to sprinklers and heat detectors
        'these should be assigned to global constants

        ''sprinklers and heat detectors
        ''standard response sprinkler
        'RTI(i) = gcs_RTI_cspr 'Response Time Index 5 mm bulb
        'RadialDistance(i) = 4.3
        'cfactor(i) = gcs_cfactor 'conduction factor
        'ActuationTemp(i) = 68 + 273
        'WaterSprayDensity(i) = 0
        'SprinkDistance(i) = gcs_sprdist

        'smoke detectors
        SmokeOD(i) = 0.14
        SDdelay(i) = 15
        DetSensitivity(i) = 2.5
        SDRadialDist(i) = 0
        SDdepth(i) = 0.025

        DetectorType(i) = 0

    End Sub
    Public Function Chardepth(ByVal idr As Integer, ByRef wFLED As Double)

        Dim j As Integer
        Dim crittemp, wDensity, wHC As Double
        Dim proportion As Double

        Try
            'idr = 1 'current room
            crittemp = chartemp + 273 'char temp
            j = stepcount 'current timestep
            'wDensity = 450 'density of wood kg/m3
            wDensity = WallDensity(idr)
            'wHC = 15 'heat of combustion of wood MJ/kg
            wHC = WallEffectiveHeatofCombustion(idr)
            proportion = CLTwallpercent / 100 'percentage should relate to the entire wall area including vents

            'define variables
            Dim depth As Single = 0
            Dim maxtemp As Double = 0
            Dim NumberwallNodes As Integer
            NumberwallNodes = UBound(UWallNode, 2)

            Dim mydepth As Double
            Dim X(0 To NumberwallNodes) As Double
            Dim Y(0 To NumberwallNodes) As Double
            Dim T(0 To NumberwallNodes) As Double

            For curve = 1 To NumberwallNodes
                X(NumberwallNodes - curve + 1) = UWallNode(idr, curve, j) 'descending order
                If HaveWallSubstrate(idr) = True Then
                    depth = (curve - 1) * WallThickness(idr) / (Wallnodes - 1)
                    If depth > WallThickness(idr) Then
                        depth = WallThickness(idr) + (curve - Wallnodes) * WallSubThickness(idr) / (Wallnodes - 1)
                    End If
                Else
                    depth = (curve - 1) * WallThickness(idr) / (NumberwallNodes - 1)
                End If
                Y(NumberwallNodes - curve + 1) = depth
                T(NumberwallNodes - curve + 1) = j * Timestep

            Next
            'char depth by interpolation
            Interpolate_D(X, Y, NumberwallNodes, crittemp, mydepth)
            Chardepth = mydepth
            ' If mydepth > 0 Then Stop
            'calculate the volume of CLT

            'wall vent area
            'For j = idr To NumberRooms + 1
            '    For m = 1 To NumberVents(idr, j)
            '        Dim avent As Double = VentHeight(idr, j, m) * VentWidth(idr, j, m)
            '    Next
            'Next

            Dim w_area As Double = 2 * (RoomLength(idr) + RoomWidth(idr)) * RoomHeight(idr) * proportion
            Dim volume As Double = w_area * Chardepth / 1000 'm3
            Dim wMass As Double = volume * wDensity 'kg
            wFLED = wMass * wHC / RoomFloorArea(idr) 'MJ/m2

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in chardepth()")
        End Try
    End Function
    Public Function Chardepth_ceil(ByVal idr As Integer, ByRef cFLED As Double)

        Dim j As Integer
        Dim crittemp, cDensity, cHC As Double
        Dim proportion As Double

        Try
            'idr = 1 'current room
            crittemp = chartemp + 273 'char temp
            j = stepcount 'current timestep
            cDensity = CeilingDensity(idr)
            cHC = CeilingEffectiveHeatofCombustion(idr)
            proportion = CLTceilingpercent / 100

            'define variables
            Dim depth As Single = 0
            Dim maxtemp As Double = 0
            Dim NumberceilingNodes As Integer
            NumberceilingNodes = UBound(CeilingNode, 2)

            Dim mydepth As Double
            Dim X(0 To NumberceilingNodes) As Double
            Dim Y(0 To NumberceilingNodes) As Double
            Dim T(0 To NumberceilingNodes) As Double

            For curve = 1 To NumberceilingNodes
                X(NumberceilingNodes - curve + 1) = CeilingNode(idr, curve, j)  'descending order
                If HaveCeilingSubstrate(idr) = True Then
                    depth = (curve - 1) * CeilingThickness(idr) / (Ceilingnodes - 1)
                    If depth > CeilingThickness(idr) Then
                        depth = CeilingThickness(idr) + (curve - Ceilingnodes) * CeilingSubThickness(idr) / (Ceilingnodes - 1)
                    End If
                Else
                    depth = (curve - 1) * CeilingThickness(idr) / (NumberceilingNodes - 1)
                End If
                Y(NumberceilingNodes - curve + 1) = depth
                T(NumberceilingNodes - curve + 1) = j * Timestep

            Next
            'char depth by interpolation
            Interpolate_D(X, Y, NumberceilingNodes, crittemp, mydepth)
            Chardepth_ceil = mydepth

            'calculate the volume of CLT
            Dim c_area As Double = RoomLength(idr) * RoomWidth(idr) * proportion
            Dim volume As Double = c_area * Chardepth_ceil / 1000 'm3
            Dim cMass As Double = volume * cDensity 'kg
            cFLED = cMass * cHC / RoomFloorArea(idr) 'MJ/m2

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in chardepth_ceil()")
        End Try
    End Function
    Function RightJustified(ByVal ColumnValue, ByVal ColumnWidth)
        If ColumnWidth - Len(ColumnValue) >= 0 Then
            RightJustified = Space(ColumnWidth - Len(ColumnValue)) & ColumnValue
        Else
            RightJustified = Space(0) & ColumnValue
        End If
    End Function
    Function LeftJustified(ByVal ColumnValue, ByVal ColumnWidth)
        If ColumnWidth - Len(ColumnValue) >= 0 Then
            LeftJustified = ColumnValue & Space(ColumnWidth - Len(ColumnValue))
        Else
            LeftJustified = ColumnValue & Space(0)
        End If
    End Function
    Public Sub Create_excel()
        '===========================================================
        '   Saves results directly in an excel file
        '===========================================================

        Dim s As String
        Dim room As Integer
        Dim k, j, count As Integer
        Dim oExcel As Object
        Dim oBook As Object
        Dim oSheet As Object

        Try

            'Dim SaveBox As New SaveFileDialog()

            ''dialogbox title
            'SaveBox.Title = "Save Results to Excel File"

            ''Set filters
            'SaveBox.Filter = "All Files (*.*)|*.*|xlsx Files (*.xlsx)|*.xlsx|xls Files (*.xls)|*.xls"

            ''Specify default filter
            'SaveBox.FilterIndex = 2
            'SaveBox.RestoreDirectory = True

            'default filename extension
            'SaveBox.DefaultExt = "xlsx"

            'default filename
            ' DataDirectory = ProjectDirectory
            '   If DataDirectory = "" Then DataDirectory = UserPersonalDataFolder & gcs_folder_ext & "\" & "data\"
            'SaveBox.InitialDirectory = DataDirectory

            'default filename
            'If Len(DataFile) > 0 Then
            '        SaveBox.FileName = Mid(DataFile, 1, Len(DataFile) - 4)
            '    Else
            '        MsgBox("No data, you need to load iteration first")
            '        Exit Sub

            '    End If

            'Dim excelfile As String = Mid(DataFile, 1, Len(DataFile) - 4)

            count = 0

            '========================================

            'Create an array
            Dim DataArray(0 To (NumberTimeSteps * Timestep / ExcelInterval + 1), 0 To 65) As Object

            'On Error GoTo excelerrorhandler
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            'Start a new workbook in Excel
            oExcel = CreateObject("Excel.Application")

            oBook = oExcel.Workbooks.Add

            'Add headers to the worksheet on row 1
            oSheet = oBook.Worksheets(1)

            oSheet.name = "Room 1"

            'assign values to excel cells
            'define a format string
            's = "Scientific"
            s = "0.000E+00"

            Dim rowcount As Short

            For room = 1 To NumberRooms
                'i = 1
                'For k = 0 To 10
                DataArray(0, 0) = "Time (sec)"
                DataArray(0, 1) = "Layer (m)"
                DataArray(0, 2) = "Upper Layer Temp (C)"
                DataArray(0, 3) = "HRR (kW)"
                DataArray(0, 4) = "Mass Loss Rate (kg/s)"
                DataArray(0, 5) = "Plume (kg/s)"
                DataArray(0, 6) = "Vent Fire (kW)"
                DataArray(0, 7) = "CO2 Upper(%)"
                DataArray(0, 8) = "CO Upper (ppm)"
                DataArray(0, 9) = "O2 Upper (%)"
                DataArray(0, 10) = "CO2 Lower(%)"
                DataArray(0, 11) = "CO Lower(ppm)"
                DataArray(0, 12) = "O2 Lower (%)"
                DataArray(0, 13) = "FED gases(inc)"
                DataArray(0, 14) = "Upper Wall Temp (C)"
                DataArray(0, 15) = "Ceiling Temp (C)"
                DataArray(0, 16) = "Rad on Floor (kW/m2)"
                DataArray(0, 17) = "Lower Layer Temp (C)"
                DataArray(0, 18) = "Lower Wall Temp (C)"
                DataArray(0, 19) = "Floor Temp (C)"
                DataArray(0, 20) = "Y Pyrolysis Front (m)"
                DataArray(0, 21) = "X Pyrolysis Front (m)"
                DataArray(0, 22) = "Z Pyrolysis Front (m)"
                DataArray(0, 23) = "Upward Velocity (m/s)"
                DataArray(0, 24) = "Lateral Velocity (m/s)"
                DataArray(0, 25) = "Pressure (Pa)"
                DataArray(0, 26) = "Visibility (m)"
                DataArray(0, 27) = "Vent Flow to Upper Layer (kg/s)"
                DataArray(0, 28) = "Vent Flow to Lower Layer (kg/s)"
                DataArray(0, 29) = "Rad on Target (kW/m2)"
                DataArray(0, 30) = "FED thermal(inc)"
                DataArray(0, 31) = "OD upper (1/m)"
                DataArray(0, 32) = "OD lower (1/m)"
                DataArray(0, 33) = "k upper (1/m)"
                DataArray(0, 34) = "k lower (1/m)"
                DataArray(0, 35) = "Vent Flow to Outside (m3/s)"
                DataArray(0, 36) = "HCN Upper (ppm)"
                DataArray(0, 37) = "HCN Lower (ppm)"
                DataArray(0, 38) = "SPR (m2/kg)"
                DataArray(0, 39) = "Unexposed Upper Wall Temp (C)"
                DataArray(0, 40) = "Unexposed Lower Wall Temp (C)"
                DataArray(0, 41) = "Unexposed Ceiling Temp (C)"
                DataArray(0, 42) = "Unexposed Floor Temp (C)"

                'Next
                If room = fireroom Then
                    DataArray(0, 43) = "Ceiling Jet Temp(C)"
                    DataArray(0, 44) = "Max Ceil Jet Temp(C)"
                    DataArray(0, 45) = "GER"
                    DataArray(0, 46) = "Smoke Detect OD-out (1/m)"
                    DataArray(0, 47) = "Smoke Detect OD-in (1/m)"
                    DataArray(0, 48) = "Link Temp (C)"

                    DataArray(0, 54) = "MLR free burn (kg/s)"
                    DataArray(0, 55) = "MLR ventilation effect (kg/s)"
                    DataArray(0, 56) = "MLR thermal effect (kg/s)"
                    DataArray(0, 57) = "MLR total (kg/s)"
                    DataArray(0, 58) = "Burning rate (kg/s)"
                    DataArray(0, 59) = "Fuel area shrinkage ratio (-)"
                    DataArray(0, 60) = "Unconstrained HRR (kW)"
                End If

                DataArray(0, 49) = "Normalised Heat Load (s1/2K)"
                DataArray(0, 50) = "Net ceiling flux (kW/m2)"
                DataArray(0, 51) = "Net upper wall flux (kW/m2)"
                DataArray(0, 52) = "Net lower Wall flux (kW/m2)"
                DataArray(0, 53) = "Net floor flux (kW/m2)"

                If useCLTmodel = True Then
                    DataArray(0, 61) = "Incident ceiling radiant flux (kW/m2)"
                    DataArray(0, 62) = "Incident upper wall radiant flux (kW/m2)"
                    DataArray(0, 63) = "Incident lower Wall radiant flux (kW/m2)"
                    DataArray(0, 64) = "Incident floor radiant flux (kW/m2)"
                End If

                DataArray(0, 65) = "Total fuel mass loss (kg)"

                If NumberTimeSteps > 0 Then

                    Do While NumberTimeSteps * Timestep / ExcelInterval * NumberRooms * 59 > 32000 'the maximum number of data points able to be plotted in excel chart
                        ExcelInterval = ExcelInterval * 2
                    Loop

                    k = 2 'row
                    For j = 1 To NumberTimeSteps + 1
                        count = count + 1
                        If System.Math.Round(Int(tim(j, 1) / ExcelInterval) - tim(j, 1) / ExcelInterval, 4) = 0 Then
                            MDIFrmMain.ToolStripStatusLabel3.Text = "Saving to Excel ... Please Wait - " & Format(count / (NumberRooms * NumberTimeSteps) * 100, "0") & "%"
                            DataArray(k - 1, 0) = Format(tim(j, 1), "general number")
                            DataArray(k - 1, 1) = Format(layerheight(room, j), s)
                            DataArray(k - 1, 2) = Format(uppertemp(room, j) - 273, s)
                            DataArray(k - 1, 3) = Format(HeatRelease(room, j, 2), s)
                            DataArray(k - 1, 4) = Format(FuelMassLossRate(j, 1), s)
                            DataArray(k - 1, 5) = Format(massplumeflow(j, room), s)
                            DataArray(k - 1, 6) = Format(ventfire(room, j), s)
                            DataArray(k - 1, 7) = Format(CO2VolumeFraction(room, j, 1) * 100, s)
                            DataArray(k - 1, 8) = Format(COVolumeFraction(room, j, 1) * 1000000, s)
                            DataArray(k - 1, 9) = Format(O2VolumeFraction(room, j, 1) * 100, s)
                            DataArray(k - 1, 10) = Format(CO2VolumeFraction(room, j, 2) * 100, s)
                            DataArray(k - 1, 11) = Format(COVolumeFraction(room, j, 2) * 1000000, s)
                            DataArray(k - 1, 12) = Format(O2VolumeFraction(room, j, 2) * 100, s)
                            DataArray(k - 1, 13) = Format(FEDSum(room, j), s)
                            DataArray(k - 1, 14) = Format(Upperwalltemp(room, j) - 273, s)
                            DataArray(k - 1, 15) = Format(CeilingTemp(room, j) - 273, s)
                            DataArray(k - 1, 16) = Format(Target(room, j), s)
                            DataArray(k - 1, 17) = Format(lowertemp(room, j) - 273, s)
                            DataArray(k - 1, 18) = Format(LowerWallTemp(room, j) - 273, s)
                            DataArray(k - 1, 19) = Format(FloorTemp(room, j) - 273, s)
                            DataArray(k - 1, 20) = Format(Y_pyrolysis(room, j), s)
                            DataArray(k - 1, 21) = Format(X_pyrolysis(room, j), s)
                            DataArray(k - 1, 22) = Format(Z_pyrolysis(room, j), s)
                            DataArray(k - 1, 23) = Format(FlameVelocity(room, 1, j), s)
                            DataArray(k - 1, 24) = Format(FlameVelocity(room, 2, j), s)
                            DataArray(k - 1, 25) = Format(RoomPressure(room, j), s)
                            DataArray(k - 1, 26) = Format(Visibility(room, j), s)
                            DataArray(k - 1, 27) = Format(FlowToUpper(room, j), s)
                            DataArray(k - 1, 28) = Format(FlowToLower(room, j), s)
                            DataArray(k - 1, 29) = Format(SurfaceRad(room, j), s)
                            DataArray(k - 1, 30) = Format(FEDRadSum(room, j), s)
                            DataArray(k - 1, 31) = Format(OD_upper(room, j), s)
                            DataArray(k - 1, 32) = Format(OD_lower(room, j), s)
                            DataArray(k - 1, 33) = Format(2.3 * OD_upper(room, j), s)
                            DataArray(k - 1, 34) = Format(2.3 * OD_lower(room, j), s)
                            DataArray(k - 1, 35) = Format(UFlowToOutside(room, j), s)
                            DataArray(k - 1, 36) = Format(HCNVolumeFraction(room, j, 1) * 1000000, s)
                            DataArray(k - 1, 37) = Format(HCNVolumeFraction(room, j, 2) * 1000000, s)
                            DataArray(k - 1, 38) = Format(SPR(room, j), s)
                            DataArray(k - 1, 39) = Format(UnexposedUpperwalltemp(room, j) - 273, s)
                            DataArray(k - 1, 40) = Format(UnexposedLowerwalltemp(room, j) - 273, s)
                            DataArray(k - 1, 41) = Format(UnexposedCeilingtemp(room, j) - 273, s)
                            DataArray(k - 1, 42) = Format(UnexposedFloortemp(room, j) - 273, s)

                            DataArray(k - 1, 49) = Format(NHL(1, room, j), s)
                            DataArray(k - 1, 50) = Format(QCeiling(room, j), s)
                            DataArray(k - 1, 51) = Format(QUpperWall(room, j), s)
                            DataArray(k - 1, 52) = Format(QLowerWall(room, j), s)
                            DataArray(k - 1, 53) = Format(QFloor(room, j), s)

                            If room = fireroom Then
                                DataArray(k - 1, 43) = Format(CJetTemp(j, 1, 0) - 273, s)
                                DataArray(k - 1, 44) = Format(CJetTemp(j, 2, 0) - 273, s)
                                DataArray(k - 1, 45) = Format(GlobalER(j), s)

                                If NumSmokeDetectors > 0 Then
                                    DataArray(k - 1, 46) = Format(OD_outsideSD(1, j), s) 'for SD1 only
                                    DataArray(k - 1, 47) = Format(OD_insideSD(1, j), s) 'for SD1 only
                                End If

                                DataArray(k - 1, 48) = Format(LinkTemp(room, j) - 273, s)

                                DataArray(k - 1, 54) = Format(FuelBurningRate(0, room, j), s)
                                DataArray(k - 1, 55) = Format(FuelBurningRate(1, room, j), s)
                                DataArray(k - 1, 56) = Format(FuelBurningRate(2, room, j), s)
                                DataArray(k - 1, 57) = Format(FuelBurningRate(3, room, j), s)
                                DataArray(k - 1, 58) = Format(FuelBurningRate(4, room, j), s)
                                DataArray(k - 1, 59) = Format(FuelBurningRate(5, room, j), s)
                                DataArray(k - 1, 60) = Format(HeatRelease(room, j, 1), s)

                                If useCLTmodel = True Then
                                    DataArray(k - 1, 61) = Format(QCeilingAST(room, 0, j), s)
                                    DataArray(k - 1, 62) = Format(QUpperWallAST(room, 0, j), s)
                                    DataArray(k - 1, 63) = Format(QLowerWallAST(room, 0, j), s)
                                    DataArray(k - 1, 64) = Format(QFloorAST(room, 0, j), s)
                                End If
                                DataArray(k - 1, 65) = Format(TotalFuel(j), s)

                            End If

                            Application.DoEvents()
                            k = k + 1
                        End If
                    Next j

                    rowcount = k

                End If

                'Transfer the array to the worksheet starting at cell A1
                oSheet.Range("A1").Resize(rowcount - 1, 66).Value = DataArray

                If room < NumberRooms Then
                    oBook.Worksheets.add()
                    oSheet = oBook.ActiveSheet
                    oSheet.Name = "Room " & CStr(room + 1)
                End If
            Next

            oBook.Worksheets.add()
            oSheet = oBook.ActiveSheet
            oSheet.Name = "Outside"

            DataArray(0, 0) = "Time (sec)"
            DataArray(0, 1) = "Vent Fire (kW)"

            If NumberTimeSteps > 0 Then
                k = 2 'row
                For j = 1 To NumberTimeSteps + 1
                    If Int(tim(j, 1) / ExcelInterval) - tim(j, 1) / ExcelInterval = 0 Then
                        DataArray(k - 1, 0) = Format(tim(j, 1), s)
                        DataArray(k - 1, 1) = Format(ventfire(room, j), s)
                        k = k + 1
                    End If
                Next j
            End If

            'Transfer the array to the worksheet starting at cell A1
            oSheet.Range("A1").Resize(k - 1, 2).Value = DataArray

            MDIFrmMain.ToolStripStatusLabel4.Text = "Saving Excel Charts... Please Wait" = "Saving Excel Charts... Please Wait"
            'If frmprintvar.chkLH.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,B:B", "Layer Height (m)", "B2", "A2:A" & CStr(rowcount), "B2:B" & CStr(rowcount))
            'If frmprintvar.chkUT.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,C:C", "Upper Layer Temp (C)", "C2", "A2:A" & CStr(rowcount), "C2:C" & CStr(rowcount))
            'If frmprintvar.chkLT.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,R:R", "Lower Layer Temp (C)", "R2", "A2:A" & CStr(rowcount), "R2:R" & CStr(rowcount))
            'If frmprintvar.chkHRR.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,D:D", "Heat Release (kW)", "D2", "A2:A" & CStr(rowcount), "D2:D" & CStr(rowcount))
            'If frmprintvar.chkMassLoss.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,E:E", "Mass Loss Rate (kg.s-1)", "E2", "A2:A" & CStr(rowcount), "E2:E" & CStr(rowcount))
            'If frmprintvar.chkPlume.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,F:F", "Plume Flow (kg.s-1)", "F2", "A2:A" & CStr(rowcount), "F2:F" & CStr(rowcount))
            'If frmprintvar.chkFED.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,N:N", "FED gases", "N2", "A2:A" & CStr(rowcount), "N2:N" & CStr(rowcount))
            'If frmprintvar.chkVisi.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,AA:AA", "Visibility (m)", "AA2", "A2:A" & CStr(rowcount), "AA2:AA" & CStr(rowcount))
            'If frmprintvar.chkFEDrad.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,AE:AE", "FED thermal", "AE2", "A2:A" & CStr(rowcount), "AE2:AE" & CStr(rowcount))
            'If frmprintvar.chkODupper.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,AF:AF", "Upper Layer OD (m-1)", "AF2", "A2:A" & CStr(rowcount), "AF2:AF" & CStr(rowcount))
            'If frmprintvar.chkODlower.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,AG:AG", "Lower Layer OD (m-1)", "AG2", "A2:A" & CStr(rowcount), "AG2:AG" & CStr(rowcount))
            'If frmprintvar.chkODsmoke.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,AU:AU", "Smoke Detector OD-OUTSIDE (m-1)", "AU2", "A2:A" & CStr(rowcount), "AU2:AU" & CStr(rowcount))
            'If frmprintvar.chkODsmoke.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,AV:AV", "Smoke Detector OD-INSIDE (m-1)", "AV2", "A2:A" & CStr(rowcount), "AV2:AV" & CStr(rowcount))
            'If frmprintvar.chkLCO.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,L:L", "Lower Layer CO (ppm)", "L2", "A2:A" & CStr(rowcount), "L2:L" & CStr(rowcount))
            'If frmprintvar.chkUCO.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,I:I", "Upper Layer CO (ppm)", "I2", "A2:A" & CStr(rowcount), "I2:I" & CStr(rowcount))
            'If frmprintvar.chkPressure.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,Z:Z", "Pressure (Pa)", "Z2", "A2:A" & CStr(rowcount), "Z2:Z" & CStr(rowcount))

            'save the worksheet
            'On Error Resume Next
            Dim fname As String = FileSystem.Dir(basefile)
            fname = Mid(fname, 1, Len(fname) - 4) & "_results"

            oExcel.DisplayAlerts = False

            'Save the Workbook and Quit Excel
            oBook.SaveAs(RiskDataDirectory & fname)
            oBook.Close(SaveChanges:=False)
            oExcel.DisplayAlerts = True
            oExcel.Quit()



            If Err.Number = 1004 Then
                'MsgBox(SaveBox.FileName & " is already Open. Please close and then try again.", MsgBoxStyle.OkOnly)
                Err.Clear()
            Else
                'MsgBox("Data saved in " & DataDirectory & fname & ".xlsx", MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
            End If

            oExcel.Application.Quit()
            'release the objects
            oExcel = Nothing
            oBook = Nothing
            oSheet = Nothing

            MDIFrmMain.ToolStripStatusLabel3.Text = ""
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Arrow
            Exit Sub


        Catch ex As Exception
            'oExcel.Application.Quit()
            'release the objects
            oExcel = Nothing
            oBook = Nothing
            oSheet = Nothing
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in module1.vb create_excel")
        End Try
    End Sub

End Module