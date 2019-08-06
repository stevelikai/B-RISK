Option Strict Off
Option Explicit On
Imports System.Collections.Generic
Imports System.Math
Imports CenterSpace.NMath.Core
Imports CenterSpace.NMath.Analysis
Module DIFFEQNS

       Function DMin1(ByVal x1 As Double, ByVal x2 As Double) As Double
        '**********************************************************
        '*  Function to determine the minimum of two values.
        '*
        '*  Revised: Colleen Wade 24 July 1995
        '**********************************************************

        If x1 > x2 Then
            DMin1 = x2
        Else
            DMin1 = x1
        End If
    End Function

    Function Sign(ByRef x1 As Double, ByRef x2 As Double) As Double
        '//////////////////////////////////////////////////////////////////////
        '*  From Promath 2.0
        '//////////////////////////////////////////////////////////////////////

        If x2 <> 0.0# Then
            Sign = Abs(x1) * x2 / Abs(x2)
        Else
            Sign = Abs(x1)
        End If

    End Function


    Public Sub Layer_Height(ByVal upper_vol As Double, ByRef layerz As Double, ByVal room As Short)
        '====================================================================
        '   finds the height of the layer interface given an upper layer volume
        '
        '   17/6/97 C A Wade
        '====================================================================
        'If upper_vol < 0.0001 * RoomVolume(room) Then upper_vol = 0.0001 * RoomVolume(room)

        'find the height of the smoke layer above the floor based on upper layer volume
        If CeilingSlope(room) = False Or RoomHeight(room) = MinStudHeight(room) Then
            'for a flat ceiling
            layerz = (RoomVolume(room) - upper_vol) / RoomFloorArea(room)

            'If layerz > 0.999 * RoomHeight(room) Then layerz = 0.999 * RoomHeight(room)
        ElseIf CeilingSlope(room) = True Then
            'sloping ceiling
            If upper_vol > RoomFloorArea(room) * (RoomHeight(room) - MinStudHeight(room)) / 2 Then
                'layer height is below the eaves level
                layerz = (RoomVolume(room) - upper_vol) / RoomFloorArea(room)
            Else
                'layer height is above the eaves level - checked okay 26/4/2010
                layerz = RoomHeight(room) - Sqrt((2 * upper_vol * (RoomHeight(room) - MinStudHeight(room))) / (RoomFloorArea(room)))
            End If
        End If
        If layerz < 0.005 Then
            layerz = 0.005 'minimum lower layer is 5 mm, forced
        End If

    End Sub

    Public Sub heat_losses_stiff(ByVal room As Integer, ByVal X As Double, ByVal Z As Double, ByVal UT As Double, ByVal LT As Double, ByVal UTW As Double, ByVal TC As Double, ByVal LTW As Double, ByVal FT As Double, ByVal Q As Double, ByRef upperabsorb As Double, ByRef lowerabsorb As Double, ByRef matc(,) As Double, ByVal YCO2 As Double, ByVal YH2O As Double, ByVal YCO2Lower As Double, ByVal YH2Olower As Double, ByVal volume As Double)
        '*  =====================================================================
        '*  This function returns the value of the heat losses by convection
        '*  and radiation from the hot gas layer to the ceiling and that part of
        '*  the walls in contact with the upper gas layer. An explicit
        '*  finite-difference scheme is used.
        '*
        '*  Arguments passed to the function are:
        '*      Z  = layer height above the fire (m)
        '*      UT = upper gas layer temp (K)
        '*      LT = lower gas layer (K)
        '*      UTW = upper wall temp (K)
        '*      TC = ceiling temp (K)
        '*      Q = total heat release rate (kW)
        '*
        '*  Called by: derivs
        '*  Calls: FourWallRad, Upper_Layer_Absorb
        '*
        '*  Revised: Colleen Wade 13 July 1995
        '*  =====================================================================
        'called once per timestep

        'Dim InsideConvCoeff As Double
        Dim InsideConvCoeff1 As Double
        Dim InsideConvCoeff2 As Double
        Dim InsideConvCoeff3 As Double
        Dim InsideConvCoeff4 As Double
        Dim VentAreaUpper, VentAreaLower As Double
        Dim prd(4, 1) As Double
        Dim A3, A1, A2, A4 As Double
        Dim UpperAbsorbRad, LowerAbsorbRad As Double
        Dim lowerwallconvect, ceilingconvect, upperwallconvect, floorconvect As Double

        If UT < InteriorTemp - 1 And UT < ExteriorTemp - 1 Then
            UT = uppertemp(room, stepcount)
        End If

        'get radiant heat fluxes to room surfaces
        Call FourWallRad(room, Q, TC, UTW, LTW, FT, UT, LT, Z, prd, A1, A2, A3, A4, matc, YCO2, YH2O, YCO2Lower, YH2Olower, volume)
        'A1 to A4 includes both surface and vent areas

        'get vent areas (that are open)
        Call Get_Vent_Area(room, X, VentAreaUpper, VentAreaLower, Z)

        'get radiant heat absorbed by the upper layer
        UpperAbsorbRad = Upper_Layer_Absorb(room, VentAreaUpper, VentAreaLower, Z, prd, A1, A2, A3, A4, UT, LT, TC, UTW, LTW, FT) 'kW

        'get convective component for upper layer
        'find convection components
        InsideConvCoeff1 = Get_Convection_Coefficient2(A1, UT, TC, CStr(HORIZONTAL))
        ceilingconvect = InsideConvCoeff1 / 1000 * (TC - UT) * A1

        InsideConvCoeff2 = Get_Convection_Coefficient2(A2, UT, UTW, CStr(VERTICAL))
        upperwallconvect = InsideConvCoeff2 / 1000 * (UTW - UT) * (A2 - VentAreaUpper)

        'get total heat added to upper layer
        upperabsorb = UpperAbsorbRad + ceilingconvect + upperwallconvect

        'get radiant heat absorbed by the lower layer
        LowerAbsorbRad = Lower_Layer_Absorb(room, VentAreaUpper, VentAreaLower, Z, prd, A1, A2, A3, A4, UT, LT, TC, UTW, LTW, FT) 'kW

        'get convective component for lower layer
        InsideConvCoeff3 = Get_Convection_Coefficient2(A3, LT, LTW, CStr(VERTICAL))
        lowerwallconvect = InsideConvCoeff3 / 1000 * (LTW - LT) * (A3 - VentAreaLower)

        InsideConvCoeff4 = Get_Convection_Coefficient2(A4, LT, FT, CStr(HORIZONTAL))
        floorconvect = InsideConvCoeff4 / 1000 * (FT - LT) * A4

        'get total heat added to lower layer
        lowerabsorb = LowerAbsorbRad + floorconvect + lowerwallconvect 'kW

        'find the heat flux to the ceiling (out of ceiling +ve)
        'add convective and radiative parts
        If A1 > 0 Then
            ceilingconvect = ceilingconvect / A1 'kW/m2
            QCeiling(room, stepcount) = ceilingconvect + prd(1, 1) 'kW/m2
            QCeilingAST(room, 3, stepcount) = -prd(1, 1) 'net radiant flux
            QCeilingAST(room, 0, stepcount) = matc(1, 1) 'incident radiant fluxes
            QCeilingAST(room, 1, stepcount) = InsideConvCoeff1 'W/m2K
            'If stepcount = 20000 Then Stop
        Else
            QCeiling(room, stepcount) = 0
        End If

        'find the heat flux to the upper wall (out of wall +ve)
        If (A2 - VentAreaUpper) > 0 Then
            upperwallconvect = upperwallconvect / (A2 - VentAreaUpper)
            QUpperWall(room, stepcount) = upperwallconvect + prd(2, 1) 'kW/m2
            QUpperWallAST(room, 3, stepcount) = -prd(2, 1)
            QUpperWallAST(room, 0, stepcount) = matc(2, 1)
            QUpperWallAST(room, 1, stepcount) = InsideConvCoeff2
        Else
            QUpperWall(room, stepcount) = 0
        End If

        'find the heat flux to the floor  (out of floor +ve)
        If A4 > 0 Then
            floorconvect = floorconvect / A4
            QFloor(room, stepcount) = floorconvect + prd(4, 1) 'kW/m2
            QFloorAST(room, 3, stepcount) = -prd(4, 1)
            QFloorAST(room, 0, stepcount) = matc(4, 1)
            QFloorAST(room, 1, stepcount) = InsideConvCoeff4
        Else
            QFloor(room, stepcount) = 0
        End If

        'find the heat flux to the lower walls  (out of wall +ve)
        If (A3 - VentAreaLower) > 0 Then
            lowerwallconvect = lowerwallconvect / (A3 - VentAreaLower)
            QLowerWall(room, stepcount) = lowerwallconvect + prd(3, 1) 'kW/m2
            QLowerWallAST(room, 3, stepcount) = -prd(3, 1)
            QLowerWallAST(room, 0, stepcount) = matc(3, 1)
            QLowerWallAST(room, 1, stepcount) = InsideConvCoeff3
        Else
            QLowerWall(room, stepcount) = 0
        End If

        'total heat into the room surfaces
        TotalLosses(room, stepcount) = -(QCeiling(room, stepcount) * A1 + QUpperWall(room, stepcount) * (A2 - VentAreaUpper) + QLowerWall(room, stepcount) * (A3 - VentAreaLower) + QFloor(room, stepcount) * A4) 'kW

        If calcFRR = True Then
            'normalised heat load
            Dim sum As Double = 0
            Dim A As Double = 0

            If QCeiling(room, stepcount) < 0 Then
                NHL(3, room, stepcount) = NHL(3, room, stepcount - 1) - Timestep * QCeiling(room, stepcount) / Sqrt(ThermalInertiaCeiling(room))
                sum = -A1 * QCeiling(room, stepcount) / Sqrt(ThermalInertiaCeiling(room))
                A = A1
            Else
                NHL(3, room, stepcount) = NHL(3, room, stepcount - 1)
            End If
            If QUpperWall(room, stepcount) < 0 Then
                NHL(4, room, stepcount) = NHL(4, room, stepcount - 1) - Timestep * QUpperWall(room, stepcount) / Sqrt(ThermalInertiaWall(room))
                sum = sum - (A2 - VentAreaUpper) * QUpperWall(room, stepcount) / Sqrt(ThermalInertiaWall(room))
                A = A + (A2 - VentAreaUpper)
            Else
                NHL(4, room, stepcount) = NHL(4, room, stepcount - 1)
            End If
            If QLowerWall(room, stepcount) < 0 Then
                NHL(5, room, stepcount) = NHL(5, room, stepcount - 1) - Timestep * QLowerWall(room, stepcount) / Sqrt(ThermalInertiaWall(room))
                sum = sum - (A3 - VentAreaLower) * QLowerWall(room, stepcount) / Sqrt(ThermalInertiaWall(room))
                A = A + (A3 - VentAreaLower)
            Else
                NHL(5, room, stepcount) = NHL(5, room, stepcount - 1)
            End If
            If QFloor(room, stepcount) < 0 Then
                NHL(6, room, stepcount) = NHL(6, room, stepcount - 1) - Timestep * QFloor(room, stepcount) / Sqrt(ThermalInertiaFloor(room))
                sum = sum - A4 * QFloor(room, stepcount) / Sqrt(ThermalInertiaFloor(room))
                A = A + A4
            Else
                NHL(6, room, stepcount) = NHL(6, room, stepcount - 1)
            End If

            If A > 0 Then
                NHL(0, room, stepcount) = sum / A 's^-1/2 K area weighted NHL
                NHL(1, room, stepcount) = NHL(1, room, stepcount - 1) + Timestep * (NHL(0, room, stepcount) + NHL(0, room, stepcount - 1)) / 2
            Else
                NHL(1, room, stepcount) = NHL(1, room, stepcount - 1)
            End If

            Call frmInputs.calcAST2(stepcount, room)


        End If
    End Sub


    Public Function SMax1(ByVal x1 As Single, ByVal x2 As Single) As Double
        '*******************************************************************
        '*  Function to return the maximum of two values.
        '*
        '*  Created: Colleen Wade 24 July 1995
        '*******************************************************************

        If x1 > x2 Then SMax1 = x1 Else SMax1 = x2

    End Function
    Function DMax1(ByVal x1 As Double, ByVal x2 As Double) As Double
        '*******************************************************************
        '*  Function to return the maximum of two values.
        '*
        '*  Created: Colleen Wade 24 July 1995
        '*******************************************************************

        If x1 > x2 Then DMax1 = x1 Else DMax1 = x2

    End Function
    Public Sub hrr_estimate_fuelresponse(ByVal room As Integer, ByVal Mass_Upper As Double, ByRef massplume As Double, ByRef heatreleaselimit As Double, ByVal X As Double, ByVal layerheight As Double, ByVal uppertemp As Double, ByVal lowertemp As Double, ByVal O2Upper As Double, ByVal O2Lower As Double, ByVal mw_upper As Double, ByVal mw_lower As Double, ByVal TUHC As Double, ByVal hc As Double, ByVal incidentflux As Double)
        'routine only called for fireroom
        'iterates between mplume and ventilation limited burning
        Dim heatreleased, Qplume, burningrate As Double
        Dim count As Integer
        Dim mrate() As Double
        ReDim mrate(NumberObjects)

        'determine theoretical heat release rate
        'hrrmax = Composite_HRR(X)
        Dim i As Integer = stepcount

        Dim idg As Integer = 1
        Dim S As Double = 15.1 'stoichiometric air to fuel ratio for heptane

        'well ventilated, theoretical
        'HeatRelease(fireroom, i, 1) = EnergyYield(1) * fuelmassloss 'kW
        'HeatRelease(fireroom, i, 2) = EnergyYield(1) * FuelBurningRate(3, fireroom, i) 'kW


        'redo 
        Call mass_rate_withfuelresponse(X, mrate, 0, 0, 0, O2Lower, lowertemp, massplume, incidentflux, burningrate)

        heatreleaselimit = 1000 * EnergyYield(1) * burningrate 'kW

        count = 1
        heatreleased = HeatRelease(fireroom, i, 1) 'max theoretical hrr from the fuel
        'Qplume = HeatRelease(fireroom, i, 2)
        Qplume = heatreleaselimit

        If heatreleased > 0 Then
            Do While Abs(heatreleased - heatreleaselimit) / heatreleased > 0.001

                'fire is ventilation-limited
                heatreleased = heatreleaselimit
                If heatreleased = 0 Then Exit Do

                'Mass flow in the plume
                massplume = Mass_Plume_2012(X, layerheight, heatreleased, uppertemp, lowertemp)

                'recalculate oxygen limit using new plume flow
                'HeatRelease(fireroom, i, 2) = O2_limit_cfast(fireroom, Mass_Upper, heatreleased, mplume, uppertemp(fireroom, i), O2MassFraction(fireroom, i, 1), O2MassFraction(fireroom, i, 2), mw_upper, mw_lower, Qplume)
                Call mass_rate_withfuelresponse(X, mrate, 0, 0, 0, O2Lower, lowertemp, massplume, incidentflux, burningrate)

                heatreleaselimit = 1000 * EnergyYield(1) * burningrate 'kW

                count = count + 1
                If count > 50 Then
                    Exit Do
                End If
            Loop
        End If

        massplume = Mass_Plume_2012(X, layerheight, heatreleaselimit, uppertemp, lowertemp)

    End Sub
    Public Sub hrr_estimate(ByVal room As Integer, ByVal Mass_Upper As Double, ByRef massplume As Double, ByRef heatreleaselimit As Double, ByVal X As Double, ByVal layerheight As Double, ByVal uppertemp As Double, ByVal lowertemp As Double, ByVal O2Upper As Double, ByVal O2Lower As Double, ByVal mw_upper As Double, ByVal mw_lower As Double, ByVal TUHC As Double, ByVal hc As Double)
        'routine only called for fireroom
        Dim heatreleased, hrrmax, Qplume As Double
        Dim count As Integer

        'determine theoretical heat release rate
        'hrrmax = Composite_HRR(X)
        If TUHC > 0.00001 Then

            '%%% cw here 16082014
            'hrrmax = Composite_HRR(X) + TUHC * Mass_Upper * hc * 1000 * Timestep 'could burn upper layer tuhc in 1 second
            hrrmax = Composite_HRR(X)
        Else
            hrrmax = Composite_HRR(X)
        End If

        'determine mass flow in the plume based on theoretical HRR
        'but this shouldn't include the Q due to unburned fuel in the upper layer ???
        massplume = Mass_Plume_2012(X, layerheight, hrrmax, uppertemp, lowertemp)
        'massplume = massplumeflow(stepcount, 1)

        'determine oxygen limited heat release rate
        heatreleaselimit = O2_limit_cfast(room, Mass_Upper, hrrmax, massplume, uppertemp, O2Upper, O2Lower, mw_upper, mw_lower, Qplume)
        'If heatreleaselimit > 10 Then Stop
        count = 1
        'If hrrmax < 0.000001 Then hrrmax = 0.000001
        heatreleased = hrrmax

        'If heatreleased > 0 Then
        If heatreleased > 0.1 Then
            'Do While Abs(heatreleased - Qplume) / heatreleased > 0.001
            Do While Abs(heatreleased - heatreleaselimit) / heatreleased > 0.001
                'Do While Abs(heatreleased - heatreleaselimit) > 1
                'fire is ventilation-limited
                heatreleased = heatreleaselimit
                'heatreleased = Qplume
                If heatreleased = 0 Then Exit Do
                'Mass flow in the plume
                'massplume = Mass_Plume(ByVal X, ByVal layerheight, ByVal heatreleased, ByVal uppertemp, ByVal lowertemp)
                massplume = Mass_Plume_2012(X, layerheight, Qplume, uppertemp, lowertemp)

                'recalculate oxygen limit using new plume flow
                heatreleaselimit = O2_limit_cfast(room, Mass_Upper, heatreleased, massplume, uppertemp, O2Upper, O2Lower, mw_upper, mw_lower, Qplume)
                count = count + 1
                If count > 50 Then
                    Exit Do
                End If
            Loop
            'If X = 160 Then Stop
            massplume = Mass_Plume_2012(X, layerheight, Qplume, uppertemp, lowertemp)

        End If
    End Sub



    Public Sub heat_losses_stiff2(ByVal room As Integer, ByVal X As Double, ByVal Z As Double, ByVal UT As Double, ByVal LT As Double, ByVal UTW As Double, ByVal TC As Double, ByVal LTW As Double, ByVal FT As Double, ByVal Q As Double, ByRef upperabsorb As Double, ByRef lowerabsorb As Double, ByRef matc(,) As Double, ByVal YCO2 As Double, ByVal YH2O As Double, ByVal YCO2Lower As Double, ByVal YH2Olower As Double, ByVal volume As Double)
        '*  =====================================================================
        '*  This function returns the value of the heat losses by convection
        '*  and radiation from the hot gas layer to the ceiling and that part of
        '*  the walls in contact with the upper gas layer. An explicit
        '*  finite-difference scheme is used.
        '*
        '*  Arguments passed to the function are:
        '*      Z  = layer height above the fire (m)
        '*      UT = upper gas layer temp (K)
        '*      LT = lower gas layer (K)
        '*      UTW = upper wall temp (K)
        '*      TC = ceiling temp (K)
        '*      Q = total heat release rate (kW)
        '*
        '*  Called by: derivs
        '*  Calls: FourWallRad, Upper_Layer_Absorb
        '*
        '*  Revised: Colleen Wade 13 July 1995
        '*  =====================================================================
        'If stepcount = 100 Then Stop
        Dim InsideConvCoeff As Double
        Dim VentAreaUpper, VentAreaLower As Double
        Dim prd(4, 1) As Double
        Dim A3, A1, A2, A4 As Double
        Dim UpperAbsorbRad, LowerAbsorbRad As Double
        Dim lowerwallconvect, ceilingconvect, upperwallconvect, floorconvect As Double

        If UT < InteriorTemp - 1 And UT < ExteriorTemp - 1 Then
            UT = uppertemp(room, stepcount)
        End If

        'get radiant heat fluxes to room surfaces
        Call FourWallRad(room, Q, TC, UTW, LTW, FT, UT, LT, Z, prd, A1, A2, A3, A4, matc, YCO2, YH2O, YCO2Lower, YH2Olower, volume)

        'get vent areas
        Call Get_Vent_Area(room, X, VentAreaUpper, VentAreaLower, Z)

        'get radiant heat absorbed by the upper layer
        UpperAbsorbRad = Upper_Layer_Absorb(room, VentAreaUpper, VentAreaLower, Z, prd, A1, A2, A3, A4, UT, LT, TC, UTW, LTW, FT) 'kW

        'get convective component for upper layer
        'find convection components
        InsideConvCoeff = Get_Convection_Coefficient2(A1, UT, TC, CStr(HORIZONTAL))
        ceilingconvect = InsideConvCoeff / 1000 * (TC - UT) * A1

        InsideConvCoeff = Get_Convection_Coefficient2(A2, UT, UTW, CStr(VERTICAL))
        upperwallconvect = InsideConvCoeff / 1000 * (UTW - UT) * A2

        'get total heat added to upper layer
        upperabsorb = UpperAbsorbRad + ceilingconvect + upperwallconvect

        'get radiant heat absorbed by the lower layer
        LowerAbsorbRad = Lower_Layer_Absorb(room, VentAreaUpper, VentAreaLower, Z, prd, A1, A2, A3, A4, UT, LT, TC, UTW, LTW, FT) 'kW

        'get convective component for lower layer
        InsideConvCoeff = Get_Convection_Coefficient2(A3, LT, LTW, CStr(VERTICAL))
        lowerwallconvect = InsideConvCoeff / 1000 * (LTW - LT) * A3

        InsideConvCoeff = Get_Convection_Coefficient2(A4, LT, FT, CStr(HORIZONTAL))
        floorconvect = InsideConvCoeff / 1000 * (FT - LT) * A4

        'get total heat added to lower layer
        lowerabsorb = LowerAbsorbRad + floorconvect + lowerwallconvect

        'If lowerabsorb < gcd_Machine_Error Then lowerabsorb = 0
        'If upperabsorb < gcd_Machine_Error Then upperabsorb = 0

    End Sub

    Private Sub Smoke_Detector_Flags(ByVal room As Short, ByVal i As Integer, ByRef tim As Double, ByRef depth As Double)
        '*  =====================================================
        '*      This procedure evaluates the endpoint conditions
        '*      and sets flag if the condition is reached.
        '*
        '*  14/3/2008
        '*  =====================================================

        'Dim diameter, h, SDConcentration As Single
        'Dim OpticalDensity, AvgExtinction As Single
        Dim velocity, Tau As Double
        Dim v2, L, transit As Double
        Dim s(1) As Double

        If HaveSD(room) = True Then
            If room = fireroom Then
                'get the gas velocity at the location of the smoke detector
                'velocity = CeilingJet_Velocity2(HeatRelease(room, i, 2), room, layerheight(room, i), SDRadialDist(room), SDdepth(room))
                velocity = CeilingJet_Velocity3(HeatRelease(room, i, 2), room, layerheight(room, i), SDRadialDist(room), SDdepth(room))

                If velocity < 0.01 Then velocity = 0.01 'to prevent a divide by zero error

                If SDinside(room) = True Then
                    'response criteria based on OD inside the detector chamber
                    L = SDdelay(room) 'characteristic length of smoke detector m
                    Tau = L / velocity

                    s(1) = OD_inside(room, i) 'optical density inside chamber

                    If Tau > 0 Then
                        Call Smoke_Solver(Tau, s, OD_outside(room, i))
                        OD_inside(room, i + 1) = s(1) 'optical density inside chamber
                    Else
                        OD_inside(room, i) = OD_outside(room, i)
                        's(1) = OD_outside(room, i)
                        'OD_inside(room, i + 1) = s(1) 'optical density inside chamber
                    End If

                    'Debug.Print tim, OD_outside(room, i), OD_inside(room, i)

                    'If OD_inside(room, i + 1) >= SmokeOD(room) And SDFlag(room) = 0 And depth > minSDdepth Then
                    If OD_inside(room, i) >= SmokeOD(room) And SDFlag(room) = 0 And depth > minSDdepth Then
                        'estimate the gas transit time to reach the detector
                        'transit = 0.813 * (1 + SDRadialDist(room) / (roomheight(room) - FireHeight(1))) * (G / SpecificHeat_air / ReferenceTemp / ReferenceDensity) ^ (-1 / 5) * ((1 - RadiantLossFraction) * HeatRelease(room, i, 1) / tim) ^ (-1 / 5) * (roomheight(room) - FireHeight(1)) ^ (4 / 5)

                        'gas velocity at r=0
                        v2 = CeilingJet_Velocity3(HeatRelease(room, i, 2), room, layerheight(room, i), 0, SDdepth(room))
                        transit = 2 * (RoomHeight(room) - FireHeight(1)) / v2 + SDRadialDist(room) * 2 / (velocity + v2)

                        If tim > transit Then 'why this?
                            SDTime(room) = tim + transit 'time of actuation
                            SDFlag(room) = 1
                            Dim Message As String = Format(SDTime(room), "0") & " sec. Smoke detector operates in room " & room.ToString
                            If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                            If i > 1 And velocity < 0.16 Then 'critical value to achieve smoke entry into the sensing chamber
                                OtherMessages = OtherMessages & Chr(13) & "At " & CStr(tim) & " seconds - The gas velocity in the ceiling jet at the location of the smoke detector in room " & Str(fireroom) & " is calculated to be less than 0.16 m/s. Smoke detector response prediction may be unreliable."
                            End If
                        End If
                    End If
                Else
                    'response criteria based on OD outside the detector chamber
                    If OD_outside(room, i) >= SmokeOD(room) And SDFlag(room) = 0 And depth > minSDdepth Then

                        v2 = CeilingJet_Velocity3(HeatRelease(room, i, 2), room, layerheight(room, i), 0, SDdepth(room))
                        'use the velocity at the ceiling to calculate transit time to ceiling, and the average velocity at the ceiling and at the detector to calculate the transit time from plume centerline to detector
                        transit = 2 * (RoomHeight(room) - FireHeight(1)) / v2 + SDRadialDist(room) * 2 / (velocity + v2)
                        If tim > transit Then
                            SDTime(room) = tim + transit 'time of actuation
                            SDFlag(room) = 1
                            Dim Message As String = Format(SDTime(room), "0") & " sec. Smoke detector operates in room " & room.ToString
                            If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                        End If
                    End If
                End If
            Else
                'smoke detector not in room of fire origin
                'response criteria based on OD outside the detector chamber
                If OD_outside(room, i) >= SmokeOD(room) And SDFlag(room) = 0 And depth > minSDdepth Then
                    SDFlag(room) = 1
                    SDTime(room) = tim 'time of actuation
                    Dim Message As String = CStr(SDTime(room)) & " sec. Smoke detector operates in room " & room.ToString
                    If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                End If

            End If
            End If
    End Sub
    Private Sub SD_Flags(ByRef oSmokeDet As Object, ByVal OD_outside As Single, ByVal room As Short, ByVal i As Integer, ByRef tim As Double, ByRef depth As Double)

        Dim velocity, Tau, maxr As Double
        Dim v2, L, transit As Double
        Dim s(1) As Double
        Dim id As Integer = oSmokeDet.sdid
        Dim OD As Single = oSmokeDet.OD
        Dim responsetime As Single = oSmokeDet.responsetime
        Dim charlength As Single = oSmokeDet.charlength


        maxr = Sqrt(RoomLength(fireroom) ^ 2 + RoomWidth(fireroom) ^ 2)

        If oSmokeDet.sdbeam = True Then
            'OD for alarm
            OD = -Log(oSmokeDet.sdbeamalarmtrans) / (2.303 * oSmokeDet.sdbeampathlength)
            oSmokeDet.OD = OD

            'is the beam in upper or lower layer
            If oSmokeDet.sdz < depth Then
                'in upper layer
                If OD_upper(room, i) >= OD And responsetime = 0 Then
                    'alarm threshold reached
                    oSmokeDet.responsetime = tim 'time of actuation

                    SDFlagSD(id) = 1

                    If SDFlag(room) = 0 Then
                        SDFlag(room) = 1
                        SDTime(room) = oSmokeDet.responsetime
                    End If

                    Dim Message As String = Format(oSmokeDet.responsetime, "0") & " sec. Beam detector " & oSmokeDet.sdid.ToString & " operates in room " & oSmokeDet.room.ToString
                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                End If
            Else
                'in lower layer
                If OD_lower(room, i) >= OD And responsetime = 0 Then
                    'alarm threshold reached
                    oSmokeDet.responsetime = tim 'time of actuation

                    SDFlagSD(id) = 1

                    If SDFlag(room) = 0 Then
                        SDFlag(room) = 1
                        SDTime(room) = oSmokeDet.responsetime
                    End If

                    Dim Message As String = Format(oSmokeDet.responsetime, "0") & " sec. Beam detector " & oSmokeDet.sdid.ToString & " operates in room " & oSmokeDet.room.ToString
                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                End If
            End If
            'no consideration of transit time
            Exit Sub
        End If

        If room = fireroom Then

            If maxr > oSmokeDet.sdr Then
                maxr = oSmokeDet.sdr 'if the specified radial distance doesn't fit in the room, then reduce it
            End If

            'get the gas velocity at the location of the smoke detector
            velocity = CeilingJet_Velocity3(HeatRelease(room, i, 2), room, layerheight(room, i), maxr, oSmokeDet.sdz)

            If velocity < 0.01 Then velocity = 0.01 'to prevent a divide by zero error

            If oSmokeDet.sdinside = True Then
                'response criteria based on OD inside the detector chamber

                'estimate the gas transit time to reach the detector
                v2 = CeilingJet_Velocity3(HeatRelease(room, i, 2), room, layerheight(room, i), 0, oSmokeDet.sdz)
                transit = 2 * (RoomHeight(room) - FireHeight(1)) / v2 + oSmokeDet.sdr * 2 / (velocity + v2)
                oSmokeDet.transit = transit

                'characteristic length of smoke detector m
                L = charlength

                Tau = L / velocity

                's(1) = OD_inside(room, i) 'optical density inside chamber
                s(1) = OD_insideSD(id, i) 'optical density inside chamber

                If Tau > 0 Then
                    Call Smoke_Solver(Tau, s, OD_outside)
                Else
                    s(1) = OD_outside
                End If

                'OD_inside(room, i + 1) = s(1) 'optical density inside chamber
                OD_insideSD(id, i + 1) = s(1) 'optical density inside chamber

                If tim > transit Then

                    If OD_insideSD(id, i) >= OD And responsetime = 0 Then

                        oSmokeDet.responsetime = tim + transit 'time of actuation


                        SDFlagSD(id) = 1

                        If SDFlag(room) = 0 Then
                            SDFlag(room) = 1
                            SDTime(room) = oSmokeDet.responsetime
                        End If

                        Dim Message As String = Format(oSmokeDet.responsetime, "0") & " sec. Smoke detector " & oSmokeDet.sdid.ToString & " operates in room " & oSmokeDet.room.ToString
                        frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                        'If i > 1 And velocity < 0.16 Then 'critical value to achieve smoke entry into the sensing chamber
                        '    Message = Format(oSmokeDet.responsetime, "0") & " sec. The gas velocity in the ceiling jet at the location of the smoke detector in room " & Str(fireroom) & " is calculated to be less than 0.16 m/s. Smoke detector response prediction may be unreliable."
                        '    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                        'End If

                    End If
                End If
            Else
                'response criteria based on OD outside the detector chamber
                v2 = CeilingJet_Velocity3(HeatRelease(room, i, 2), room, layerheight(room, i), 0, oSmokeDet.sdz)
                'use the velocity at the ceiling to calculate transit time to ceiling, and the average velocity at the ceiling and at the detector to calculate the transit time from plume centerline to detector
                transit = 2 * (RoomHeight(room) - FireHeight(1)) / v2 + maxr * 2 / (velocity + v2)
                oSmokeDet.transit = transit

                If OD_outside >= OD And responsetime = 0 Then

                    If tim > transit Then

                        oSmokeDet.responsetime = tim + transit 'time of actuation
                        'oSmokeDet.responsetime = tim  'time of actuation

                        SDFlagSD(id) = 1

                        If SDFlag(room) = 0 Then
                            SDFlag(room) = 1
                            'SDTime(room) = tim + transit
                            SDTime(room) = oSmokeDet.responsetime
                        End If
                        'If tim + transit < SDTime(room) Then SDTime(room) = tim + transit

                        Dim Message As String = Format(oSmokeDet.responsetime, "0") & " sec. Smoke detector " & oSmokeDet.sdid.ToString & " operates in room " & oSmokeDet.room.ToString
                        frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                    End If

                End If
            End If
        Else
            'smoke detector not in room of fire origin

            'check that response time is not less than the max transport time for the fireroom
            v2 = CeilingJet_Velocity3(HeatRelease(fireroom, i, 2), fireroom, layerheight(fireroom, i), 0, 0.025) 'velocity above the plume

            'get the gas velocity at the location of maxr
            velocity = CeilingJet_Velocity3(HeatRelease(fireroom, i, 2), fireroom, layerheight(fireroom, i), maxr, 0.025)

            If velocity < 0.01 Then velocity = 0.01 'to prevent a divide by zero error

            'use the velocity at the ceiling to calculate transit time to ceiling, and the average velocity at the ceiling and at the detector to calculate the transit time from plume centerline to detector
            transit = 2 * (RoomHeight(fireroom) - FireHeight(1)) / v2 + maxr * 2 / (velocity + v2)
            oSmokeDet.transit = transit  'time of actuation

            If tim > transit Then
              
                'response criteria based on OD outside the detector chamber
                If OD_outside >= OD And responsetime = 0 Then

                    oSmokeDet.responsetime = tim + transit  'time of actuation

                    SDFlagSD(id) = 1

                    If SDFlag(room) = 0 Then
                        SDFlag(room) = 1
                        SDTime(room) = oSmokeDet.responsetime
                    End If

                    Dim Message As String = Format(oSmokeDet.responsetime, "0") & " sec. Smoke detector " & oSmokeDet.sdid.ToString & " operates in room " & oSmokeDet.room.ToString
                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                End If

            End If 'transit
        End If 'room

    End Sub
    Sub main_program2(ByVal itcounter As Integer, ByVal oSprinklers As Object, ByVal osprdistributions As Object, ByVal oSmokeDets As Object, ByVal osddistributions As Object)
        '*  ====================================================================
        '*  Solve the ODE's
        '*
        '*  Calls:
        '   for version 2001.0
        '*  ====================================================================

        'vb2008

        'secondary item fire spread model
        Dim ItemFTP_sum_pilot() As Double
        Dim ItemFTP_sum_auto() As Double
        Dim ItemFTP_sum_wall() As Double
        Dim ItemFTP_sum_ceiling() As Double

        ReDim F(4, 4, NumberRooms)

        Dim SPstart() As Double
        Dim Xstart(,) As Double
        Dim fstart() As Double

        Dim X(6) As Double
        Dim k As Integer
        Dim OpticalDensity As Single
        Dim velocity(,) As Double
        Dim fvelocity() As Double
        Dim s(1) As Double
        Dim V(6) As Double
        Dim MaxLength As Double
        Dim SootMass As Double
        Dim Ystart(,) As Double
        Dim j As Short 'variables for zone model ODE's
        Dim i As Integer
        Dim qstar, layerz, deltaZ As Double
        Dim WallFTP_sum() As Double
        ReDim ItemFTP_sum_pilot(NumberObjects)
        ReDim ItemFTP_sum_auto(NumberObjects)
        ReDim ItemFTP_sum_wall(NumberObjects)
        ReDim ItemFTP_sum_ceiling(NumberObjects)
        ReDim ObjectRad(0 To 1, 0 To NumberObjects + 1, 0 To NumberTimeSteps + 1) 'vert surface =0 horizont surface=1
        Dim WallFTP_sum_ar() As Double
        Dim CeilingFTP_sum() As Double
        Dim FloorFTP_sum(NumberRooms, NumberTimeSteps + 1) As Double
        Dim IncidentFlux As Double
        Dim h As Double
        Dim VF As Double
        Dim gastemp As Double
        Dim wall, MaxTime, ceiling As Integer
        Dim Visi_upper, Visi_lower As Single
        Dim MMaxCeilingNodes, MMaxWallNodes, MMaxFloorNodes As Integer
        Dim MaxFloorNodes() As Double
        Dim GShift, Gmultiplier As Single
        Dim room As Integer
        Dim Tflame, eFLame As Double
        Dim Graphdata() As Double
        Dim vwidth As Single
        Dim MaxWallNodes() As Integer
        Dim MaxCeilingNodes() As Integer
        Dim Area As Double
        Dim projection As Double
        Dim avent, maxarea As Single
        Dim maxvent, vent As Integer
        Dim fheight As Double
        Dim rcnone, quintiere As Boolean
        Dim rgBL, rgTL, rgTR, rgBR As String
        Dim runtimegraph(3) As System.Windows.Forms.Control
        Dim graphon, ThickLines As Boolean
        Dim LAxis, LX As Double
        Dim graphdataTL() As Double
        Dim graphdataTR() As Double
        Dim graphdataBR() As Double
        Dim graphdataBL() As Double
        Dim mw_upper, QLstar, QL, awallflux_r, mw_lower As Double


        Static flux_wall_AR, diameter As Double

        ReDim enzbreakflag(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
        ReDim breakflag(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
        ReDim breakflag2(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberCVents)
        ReDim HOFlag(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
        ReDim HRRFlag(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
        ReDim HRR_threshold_time(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
        ReDim MaxWallNodes(NumberRooms)
        ReDim WallFTP_sum(NumberTimeSteps + 1)
        ReDim WallFTP_sum_ar(NumberTimeSteps + 1)
        ReDim CeilingFTP_sum(NumberTimeSteps + 1)
        'reDim FloorFTP_sum(NumberRooms, NumberTimeSteps + 1)
        ReDim MaxCeilingNodes(NumberRooms)
        ReDim MaxFloorNodes(NumberRooms)
        ReDim Ystart(NumberRooms + 1, 19) 'variables for zone model ODE's
        ReDim Xstart(NumberRooms, 6) 'flamespread variables for Quintiere's Room Corner Model

        ReDim velocity(NumberRooms, 6)
        ReDim AutoFlag(NumberRooms + 1)
        ReDim SDFlag(NumberRooms + 1)
        ReDim SDFlagSD(NumSmokeDetectors + 1)
        ReDim SDtransit(0 To NumSmokeDetectors + 1, 0 To NumberTimeSteps)
        ReDim Preserve NewEnergyYield(NumberObjects)

        'find number of sprinkler devices
        'Dim oSprinklers As New List(Of oSprinkler)
        'oSprinklers = SprinklerDB.GetSprinklers
        'NumSprinklers = oSprinklers.Count
        ReDim SPstart(NumSprinklers)
        ReDim SprinkTemp(NumSprinklers, NumberTimeSteps + 1)

        On Error Resume Next

        gb_first_time_vent = True
        ERRNO = 0
        Err.Clear()
        FHflag = False
        VentLogFlag = 0
        FlameVelocityFlag = 0
        FloorVelocityFlag = 0
        NeutralPlaneFlag = 0
        flagstop = 0
        SprinklerFlag = 0
        HDFlag = 0
        timeH = 1000000 'time at which Y front reaches the ceiling
        rcnone = frmOptions1.optRCNone.Checked
        quintiere = frmOptions1.optQuintiere.Checked
        flashover_time = SimTime + 1
        fuelmasswithCLT = 0

        MDIFrmMain.ToolStripStatusLabel3.Text = ""
        CVentlogfile = "empty"
        WVentlogfile = "empty"
        WVentlogfile2 = "empty"

        'some variables to be used in the dimensioning of temporary storage arrays
        For room = 1 To NumberRooms
            If HaveCeilingSubstrate(room) = False And HaveWallSubstrate(room) = False Then
                MaxWallNodes(room) = Wallnodes
                MaxCeilingNodes(room) = Ceilingnodes
                If useCLTmodel = True Then
                    MaxWallNodes(room) = (Wallnodes - 1) * CInt(WallThickness(room) / 1000 / Lamella) + 1
                    MaxCeilingNodes(room) = (Ceilingnodes - 1) * CInt(CeilingThickness(room) / 1000 / Lamella) + 1
                End If

            ElseIf HaveWallSubstrate(room) = True And HaveCeilingSubstrate(room) = True Then
                    MaxWallNodes(room) = 2 * Wallnodes - 1
                MaxCeilingNodes(room) = 2 * Ceilingnodes - 1
            ElseIf HaveWallSubstrate(room) = False And HaveCeilingSubstrate(room) = True Then
                MaxWallNodes(room) = Wallnodes
                MaxCeilingNodes(room) = 2 * Ceilingnodes - 1
            ElseIf HaveWallSubstrate(room) = True And HaveCeilingSubstrate(room) = False Then
                MaxWallNodes(room) = 2 * Wallnodes - 1
                MaxCeilingNodes(room) = Ceilingnodes
            End If
            If HaveFloorSubstrate(room) = True Then
                MaxFloorNodes(room) = 2 * Floornodes - 1
            Else
                MaxFloorNodes(room) = Floornodes
            End If
        Next room

        MaxTime = NumberTimeSteps + 1

        'define the arrays
       
        Dim IgWallNode(,,) As Double 'for quintiere's room corner model
        Dim IgCeilingNode(,,) As Double 'for quintiere's room corner model

        MMaxWallNodes = 0
        MMaxCeilingNodes = 0
        MMaxFloorNodes = 0
        For i = 1 To NumberRooms
            If MaxWallNodes(i) > MMaxWallNodes Then MMaxWallNodes = MaxWallNodes(i)
            If MaxCeilingNodes(i) > MMaxCeilingNodes Then MMaxCeilingNodes = MaxCeilingNodes(i)
            If MaxFloorNodes(i) > MMaxFloorNodes Then MMaxFloorNodes = MaxFloorNodes(i)
        Next i
        'dimension the arrays
        ReDim IgWallNode(NumberRooms, MMaxWallNodes, MaxTime) 'for quintiere's room corner model
        ReDim IgCeilingNode(NumberRooms, MMaxCeilingNodes, MaxTime)
        ReDim UWallNode(NumberRooms, MMaxWallNodes, MaxTime)
        ReDim CeilingNode(NumberRooms, MMaxCeilingNodes, MaxTime)
        ReDim LWallNode(NumberRooms, MMaxWallNodes, MaxTime)
        ReDim FloorNode(NumberRooms, MMaxFloorNodes, MaxTime)

        ReDim CeilingNodeMaxTemp(MMaxCeilingNodes)
        ReDim WallNodeMaxTemp(MMaxWallNodes)
        ReDim CeilingNodeStatus(MMaxCeilingNodes)
        ReDim WallNodeStatus(MMaxWallNodes)
        ReDim flamespread(1)
        ReDim product(1)
        ReDim wallparam(3, NumberTimeSteps + 1)
        ReDim cpux(NumberRooms)
        ReDim cplx(NumberRooms)

        If useCLTmodel = True Then

            ReDim wall_char(NumberTimeSteps + 1, 0 To 2) '0 MJ/m2 cumulative; 1 MLR (kg/s); 2 chardepth (m)
            ReDim ceil_char(NumberTimeSteps + 1, 0 To 2)

            If KineticModel = True Then
                'kinetic model
                Dim Zstart(,) As Double
                ReDim CeilingElementMF(MMaxCeilingNodes - 1, 4, MaxTime) 'store the residual mass fractions of each component at each timestep
                ReDim UWallElementMF(MMaxWallNodes - 1, 4, MaxTime)

                ReDim CeilingCharResidue(MMaxCeilingNodes - 1, MaxTime)
                ReDim UWallCharResidue(MMaxWallNodes - 1, MaxTime)
                ReDim CeilingResidualMass(MMaxCeilingNodes - 1, MaxTime)
                ReDim WallResidualMass(MMaxWallNodes - 1, MaxTime)
                ReDim CeilingApparentDensity(MMaxCeilingNodes - 1, MaxTime)
                ReDim WallApparentDensity(MMaxWallNodes - 1, MaxTime)
                ReDim CeilingWoodMLR(MMaxCeilingNodes - 1, MaxTime)
                ReDim WallWoodMLR(MMaxWallNodes - 1, MaxTime)
                ReDim CeilingWoodMLR_tot(MaxTime + 1)
                ReDim WallWoodMLR_tot(MaxTime + 1)


                'initialise
                For m = 1 To MMaxCeilingNodes - 1
                    For j = 0 To 3
                        CeilingElementMF(m, j, 1) = 1

                    Next
                    CeilingCharResidue(m, 1) = 0
                    CeilingResidualMass(m, 1) = 0
                    CeilingWoodMLR(m, 1) = 0
                Next

                For m = 1 To MMaxWallNodes - 1
                    For j = 0 To 3
                        UWallElementMF(m, j, 1) = 1


                    Next
                    UWallCharResidue(m, 1) = 0

                Next

                ReDim Zstart(0 To 3, MMaxCeilingNodes)  'variables for wood pyrolysis model

                'three component kinetic scheme
                Dim dt As Integer = i 'current timestep
                For count = 1 To MMaxCeilingNodes
                    For m = 0 To 3
                        Zstart(m, count) = 1
                    Next
                Next

            End If

        End If

        'initialize variables
        For room = 1 To NumberRooms
            AutoFlag(room) = False
            SDFlag(room) = 0
            SDTime(room) = 0
            SPR(room, 1) = 0
            OD_upper(room, 1) = 0
            'SDOpticalDensity(room, 1) = 0
            OD_lower(room, 1) = 0
            Visibility(room, 1) = 20
            If TwoZones(room) = True Then
                UpperVolume(room, 1) = 0.0001 * RoomVolume(room)
                'UpperVolume(room, 1) = smallvalue
                Call Layer_Height(UpperVolume(room, 1), layerz, room) 'for flat ceiling
            Else
                'model this room as a single zone
                layerz = 0.1 'm 'qqq
                'layerz = 0.005
                'UpperVolume(room, 1) = RoomVolume(room)
                UpperVolume(room, 1) = RoomVolume(room) - layerz * RoomWidth(room) * RoomLength(room)
            End If

            layerheight(room, 1) = layerz
            uppertemp(room, 1) = InteriorTemp
            lowertemp(room, 1) = InteriorTemp
            Upperwalltemp(room, 1) = InteriorTemp

            UnexposedUpperwalltemp(room, 1) = InteriorTemp
            UnexposedLowerwalltemp(room, 1) = InteriorTemp
            UnexposedCeilingtemp(room, 1) = InteriorTemp
            UnexposedFloortemp(room, 1) = InteriorTemp

            LowerWallTemp(room, 1) = InteriorTemp
            CeilingTemp(room, 1) = InteriorTemp
            FloorTemp(room, 1) = InteriorTemp
            IgCeilingTemp(room, 1) = InteriorTemp
            LinkTemp(room, 1) = InteriorTemp
            IgWallTemp(room, 1) = InteriorTemp
            IgFloorTemp(room, 1) = InteriorTemp
            pyrolarea(room, 1, 1) = 0 'wall
            pyrolarea(room, 2, 1) = 0 'ceiling
            pyrolarea(room, 3, 1) = 0 'floor
            delta_area(room, 1, 1) = 0 'wall
            delta_area(room, 2, 1) = 0 'ceiling
            delta_area(room, 3, 1) = 0 'floor
        Next room

        For j = 1 To NumSprinklers
            SprinkTemp(j, 1) = InteriorTemp
        Next

        For j = 1 To NumberObjects
            NewEnergyYield(j) = EnergyYield(j)
            NewRadiantLossFraction(j) = ObjectRLF(j)

        Next j

        NewHoC_fuel = HoC_fuel
        tim(1, 1) = 0

        'initialize species mass fractions in layers at time=0
        For room = 1 To NumberRooms + 1
            TUHC(room, 1, 1) = Smallvalue
            TUHC(room, 1, 2) = Smallvalue
            COMassFraction(room, 1, 1) = Smallvalue
            COMassFraction(room, 1, 2) = Smallvalue
            H2OMassFraction(room, 1, 1) = Initial_H2O()
            H2OMassFraction(room, 1, 2) = H2OMassFraction(room, 1, 1)
            O2MassFraction(room, 1, 1) = O2Ambient
            O2MassFraction(room, 1, 2) = O2Ambient
            CO2MassFraction(room, 1, 2) = CO2Ambient
            CO2MassFraction(room, 1, 1) = CO2Ambient
            SootMassFraction(room, 1, 2) = Smallvalue
            SootMassFraction(room, 1, 1) = Smallvalue
            HCNMassFraction(room, 1, 1) = Smallvalue
            HCNMassFraction(room, 1, 2) = Smallvalue

            If room = fireroom Then GlobalER(1) = 0 'initial value
        Next room

        '1st and only setting of value for molecular weight of air
        MW_air = MolecularWeightCO2 * CO2Ambient + MolecularWeightH2O * H2OMassFraction(1, 1, 2) + MolecularWeightO2 * O2Ambient + MolecularWeightN2 * (1 - H2OMassFraction(1, 1, 2) - CO2Ambient - O2Ambient)

     
        'initial values
        For room = 1 To NumberRooms
            RoomPressure(room, 1) = -G * FloorElevation(room) * (Atm_Pressure / (Gas_Constant / MW_air * InteriorTemp)) 'relative to atm_pressure

            Ystart(room, 1) = UpperVolume(room, 1) 'upper layer volume
            Ystart(room, 2) = uppertemp(room, 1) 'upper layer temp
            Ystart(room, 3) = lowertemp(room, 1) 'lower layer temp
            Ystart(room, 4) = RoomPressure(room, 1) 'pressure 'Pa
            Ystart(room, 5) = O2MassFraction(room, 1, 1) 'o2 mass fraction  - upper
            Ystart(room, 6) = TUHC(room, 1, 1) 'TUHC mass - upper
            Ystart(room, 7) = TUHC(room, 1, 2) 'TUHC mass - lower
            Ystart(room, 8) = COMassFraction(room, 1, 1) 'co mass fraction
            Ystart(room, 9) = COMassFraction(room, 1, 2) 'co mass fraction - lower
            Ystart(room, 10) = CO2MassFraction(room, 1, 1) 'co2 mass fraction
            Ystart(room, 11) = CO2MassFraction(room, 1, 2) 'co2 mass fraction - lower
            Ystart(room, 12) = SootMassFraction(room, 1, 1) 'soot mass fraction
            Ystart(room, 13) = SootMassFraction(room, 1, 2) 'soot mass fraction - lower
            Ystart(room, 14) = HCNMassFraction(room, 1, 2) 'Hcn mass fraction  - lower
            Ystart(room, 15) = O2MassFraction(room, 1, 2) 'o2 mass fraction  - lower
            Ystart(room, 16) = H2OMassFraction(room, 1, 1) 'H2O mass fraction  - upper
            Ystart(room, 17) = H2OMassFraction(room, 1, 2) 'H2O mass fraction  - lower
            Ystart(room, 18) = HCNMassFraction(room, 1, 1) 'HCN mass fraction  - upper
            Ystart(room, 19) = LinkTemp(room, 1) 'detector link temp
            X_pyrolysis(room, 1) = 0
            Z_pyrolysis(room, 1) = 0
        Next room

        '1st and only setting of value for specific heat of air
        'SpecificHeat_air = specific_heat("CO2", InteriorTemp) * CO2Ambient + specific_heat("H2O", InteriorTemp) * H2OMassFraction(1, 1, 2) + specific_heat("O2", InteriorTemp) * O2Ambient + specific_heat("N2", InteriorTemp) * (1 - H2OMassFraction(1, 1, 2) - CO2Ambient - O2Ambient)
        SpecificHeat_air = specific_heat_upper(fireroom, Ystart, InteriorTemp)
        cplx(fireroom) = SpecificHeat_air

        For k = 1 To NumSprinklers
            SPstart(k) = SprinkTemp(k, 1)
        Next

        'Quintiere's room corner model - initialise variables
        'these apply to the fire room
        WallIgniteTime(fireroom) = 100000000 'a high number - for Quintiere's room corner model
        YPFlag = 0

        'initialize room surface temps
        For room = 1 To NumberRooms
            For i = 1 To MaxWallNodes(room)
                UWallNode(room, i, 1) = InteriorTemp
                LWallNode(room, i, 1) = InteriorTemp
                IgWallNode(room, i, 1) = InteriorTemp 'quintiere's room corner model
            Next i

            For i = 1 To MaxCeilingNodes(room)
                IgCeilingNode(room, i, 1) = InteriorTemp
                CeilingNode(room, i, 1) = InteriorTemp
            Next i

            For i = 1 To MaxFloorNodes(room)
                FloorNode(room, i, 1) = InteriorTemp
            Next i

            FloorIgniteStep(room) = 0
            CeilingIgniteStep(room) = 0
            CeilingIgniteFlag(room) = 0
            FloorIgniteFlag(room) = 0
            CeilingIgniteTime(room) = 100000000 'a high number - for Quintiere's room corner model
            FloorIgniteTime(room) = 100000000 'a high number - for Quintiere's room corner model
            WallIgniteStep(room) = 0
            WallIgniteFlag(room) = 0
            WallIgniteTime(room) = 100000000 'a high number - for Quintiere's room corner model

            'get view factor between ceiling and floor
            F(1, 4, room) = ViewFactor_Parallel(room, RoomHeight(room))

            If room = fireroom Then
                Y_pyrolysis(room, 1) = FireHeight(burner_id) 'Y_pyrolysis height is measured from the floor
                Y_burnout(room, 1) = FireHeight(burner_id)
            Else
                Y_pyrolysis(room, 1) = 0 'Y_pyrolysis height is from the centre of the vent
            End If

            FlameVelocity(room, 1, 1) = 0 'wall and ceiling spread
            FlameVelocity(room, 2, 1) = 0
            FlameVelocity(room, 3, 1) = 0
            FloorVelocity(room, 1, 1) = 0 'floor spread
            FloorVelocity(room, 2, 1) = 0
            FloorVelocity(room, 3, 1) = 0

            Xstart(room, 1) = Y_pyrolysis(room, 1) 'Y pyrolysis front
            Xstart(room, 2) = X_pyrolysis(room, 1) 'X pyrolysis front
            Xstart(room, 3) = Y_burnout(room, 1) 'Y burnout front
            Xstart(room, 4) = Yf_pyrolysis(room, 1) 'Y pyrolysis front
            Xstart(room, 6) = Xf_pyrolysis(room, 1) 'X pyrolysis front
            Xstart(room, 5) = Yf_burnout(room, 1) 'Y burnout front

        Next room

        If usepowerlawdesignfire = True Then firstitem = 1
        If firstitemtemp = 0 Then firstitemtemp = 1

        frmInputs.rtb_log.Text = "0 sec. Item " & firstitemtemp.ToString & " " & ObjectDescription(1) & " ignited. " & Chr(13) & frmInputs.rtb_log.Text
        originalpeakhrr = PeakHRR

        'solve the differential equations for layer height, gas temperature,
        'and species for each timestep.

        'start of main loop
        '**************************************
        Dim Lf As Double
        Dim fwidth As Single
        Dim fheight2 As Double
        Dim zcount As Integer = 0
        Dim ventstatus(0 To number_vents) As Boolean

        For i = 1 To NumberTimeSteps + 1

            stepcount = i

            If FuelResponseEffects = True Then
                Call Do_Stuff_fuelresponse(i) 'Utiskul, Quintiere
            Else
                Call Do_Stuff(i)
            End If

            If ISD_windspeed > 0 Then
                Call Target_distance_ISD(i)
            Else
                'did this earlier
            End If

            'stop the calculation
            For room = 1 To NumberRooms
                If uppertemp(room, i) > 1500 + 273 Then
                    'MsgBox("Upper layer temperature has exceeded 1500C, this temperature is rather unlikely and the simulation is terminated.", MB_ICONSTOP, "BRANZFIRE")
                    uppertemp(room, i) = 1500 + 273

                    Dim Message As String = CStr(tim(i, 1)) & " sec. Upper layer temperature has exceeded 1500C, this temperature is rather unlikely and the simulation is terminated. "
                    MDIFrmMain.ToolStripStatusLabel2.Text = Message.ToString
                    If ProjectDirectory = RiskDataDirectory Then
                        frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                    End If

99:
                    If i - 2 > 1 Then
                        NumberTimeSteps = i - 2
                    Else
                        NumberTimeSteps = 1
                    End If
                    MDIFrmMain.ToolStripProgressBar1.Value = 100

                    MDIFrmMain.ToolStripStatusLabel3.Text = " "
                    TwoZones(fireroom) = True
                    Exit Sub
                ElseIf ERRNO = 6 Then
                    MsgBox(ErrorToString(ERRNO))
                    NumberTimeSteps = i - 2
                    MDIFrmMain.ToolStripProgressBar1.Value = 100

                    MDIFrmMain.ToolStripStatusLabel3.Text = " "
                    TwoZones(fireroom) = True
                    Exit Sub
                End If
            Next room

            If i <> NumberTimeSteps + 1 Then 'Tag A
                For room = 1 To NumberRooms
                    'update wall, floor and ceiling temperatures
                    'if not the last time step
                    If HaveWallSubstrate(room) = False And HaveCeilingSubstrate(room) = False Then
                        'implicit FD scheme
                        'one layer boundary
                        If useCLTmodel = True Then

                            If KineticModel = True And i > 1 Then
                                Call Implicit_Temps_Ceil_kinetic(room, (i), CeilingNode) 'any substrate is ignored for HT
                                Call Implicit_Temps_UWall_kinetic(room, (i), UWallNode) 'any substrate is ignored for HT
                                Call Implicit_Surface_Temps_LWall(room, (i), LWallNode) 'no lower wall involvement in pyrolysis, all wall burns using the uwall properties
                            Else
                                'use for simple CLT and Integral models
                                Call Implicit_Surface_Temps_CLTC_Char(room, (i), CeilingNode) 'any substrate is ignored for HT
                                Call Implicit_Surface_Temps_CLTW_Char(room, (i), UWallNode, LWallNode) 'any substrate is ignored for HT
                            End If

                            Call Implicit_Surface_Temps_floor(room, (i), FloorNode) 'floor with or without substrate, only heat transfer effect inlcuded, no CLT contrib

                            Call Debond_test()

                        Else
                                Call Implicit_Surface_Temps(room, (i), UWallNode, CeilingNode, LWallNode, FloorNode)
                        End If

                    ElseIf HaveWallSubstrate(room) = True And HaveCeilingSubstrate(room) = True Then
                        'two layer boundary surfaces
                        Call Implicit_Surface_Temps2(room, wall, ceiling, (i), UWallNode, CeilingNode, LWallNode, FloorNode)
                    ElseIf HaveWallSubstrate(room) = True And HaveCeilingSubstrate(room) = False Then
                        Call Implicit_Surface_Temps3(room, wall, (i), UWallNode, CeilingNode, LWallNode, FloorNode)
                    ElseIf HaveWallSubstrate(room) = False And HaveCeilingSubstrate(room) = True Then
                        Call Implicit_Surface_Temps4(room, ceiling, (i), UWallNode, CeilingNode, LWallNode, FloorNode)
                    End If

                    '=================================
                    ' calculate the mass loss rate using a kinetic pyrolysis model
                    If useCLTmodel = True And KineticModel = True And room = fireroom Then

                        'The ODE solver needs to go in here.
                        'Zstart_ceiling(0 to 3) 'each component

                        Call MLR_kinetic((i), MMaxCeilingNodes, MMaxWallNodes)

                    End If
                    '=================================


                    'UWallNode arguments are: room, node, timestep
                    'only use where timber is not protected
                    If useCLTmodel = True And room = fireroom And IntegralModel = False And KineticModel = False Then
                        Dim keepnode As Integer
                        Dim cdr, wFLED, cFLED, FLEDwithCLT As Double

                        cdr = chardepth(room, wFLED) 'returns char depth
                        wall_char(i, 0) = wFLED 'save the cumulative MJ/m2 contributed from by wall

                        cdr = chardepth_ceil(room, cFLED)
                        ceil_char(i, 0) = cFLED 'save the cumulative MJ/m2 contributed from by ceiling

                        FLEDwithCLT = FLED + wFLED + cFLED 'MJ/m2
                        fuelmasswithCLT = FLEDwithCLT * RoomFloorArea(room) / HoC_fuel 'kg
                        'fuelmasswithCLT = RoomFloorArea(room) * (FLED / HoC_fuel + wFLED / WallEffectiveHeatofCombustion(room) + cFLED / CeilingEffectiveHeatofCombustion(room)) 'kg

                        If g_post = False And i > 0 Then 'not using the wood crib postflashover model
                            Dim m As Double

                            'calc the mass loss rate over the previous timestep
                            m = (wFLED - wall_char(i - 1, 0)) / WallEffectiveHeatofCombustion(room) * RoomFloorArea(room) 'kg contributed over the timestep
                            If m > 0 Then wall_char(i, 1) = m / Timestep 'kg/s contributed
                            m = (cFLED - ceil_char(i - 1, 0)) / CeilingEffectiveHeatofCombustion(room) * RoomFloorArea(room)  'kg contributed over the timestep
                            If m > 0 Then ceil_char(i, 1) = m / Timestep 'kg/s contributed

                        End If

                    ElseIf useCLTmodel = True And room = fireroom And IntegralModel = True Then
                        Dim keepnode As Integer
                        Dim cdr, wFLED, cFLED, FLEDwithCLT As Double

                        cdr = chardepth(room, wFLED) 'returns char depth
                        wall_char(i, 2) = cdr / 1000 'char depth wall m

                        cdr = chardepth_ceil(room, cFLED)
                        ceil_char(i, 2) = cdr / 1000 'char depth ceiling m

                    ElseIf useCLTmodel = True And room = fireroom And KineticModel = True Then
                        Dim cdr, wFLED, cFLED As Double

                        cdr = Chardepth(room, wFLED) 'returns char depth
                        wall_char(i, 2) = cdr / 1000 'char depth wall m

                        cdr = Chardepth_ceil(room, cFLED)
                        ceil_char(i, 2) = cdr / 1000 'char depth ceiling m

                        'may need to do something for integral model + not wood crib?
                    ElseIf useCLTmodel = True And room = fireroom And CLT_instant = True Then
                        'not using this
                        Dim cdr, wFLED, cFLED As Double

                        cdr = chardepth(room, wFLED) 'returns char depth
                        wall_char(i, 0) = wFLED 'save the cumulative MJ/m2 contributed from by wall
                        wall_char(i, 2) = cdr / 1000 'char depth wall m
                        cdr = chardepth_ceil(room, cFLED)
                        ceil_char(i, 2) = cdr / 1000 'char depth ceiling m
                        ceil_char(i, 0) = cFLED 'save the cumulative MJ/m2 contributed from by ceiling

                    End If

                    'vent - glass breakage solution here
                    If i > 0 Then Call window_break(i, room, breakflag, Ystart)

                    'auto vent opening options
                    If i > 0 Then Call door_open(i, room, breakflag, HOFlag) '2009.20

                    'calculates specific heats to use at this timestep
                    'cpux(room) = specific_heat("CO", Ystart(room, 2)) * Ystart(room, 8) + specific_heat("CO2", Ystart(room, 2)) * Ystart(room, 10) + specific_heat("H2O", Ystart(room, 2)) * Ystart(room, 16) + specific_heat("O2", Ystart(room, 2)) * Ystart(room, 5) + specific_heat("N2", Ystart(room, 2)) * (1 - Ystart(room, 5) - Ystart(room, 16) - Ystart(room, 8) - Ystart(room, 10))
                    'cplx(room) = specific_heat("CO", Ystart(room, 3)) * Ystart(room, 9) + specific_heat("CO2", Ystart(room, 3)) * Ystart(room, 11) + specific_heat("H2O", Ystart(room, 3)) * Ystart(room, 17) + specific_heat("O2", Ystart(room, 3)) * Ystart(room, 15) + specific_heat("N2", Ystart(room, 3)) * (1 - Ystart(room, 15) - Ystart(room, 17) - Ystart(room, 9) - Ystart(room, 11))

                    'calculates specific heats to use at this timestep APR
                    cpux(room) = specific_heat_upper(room, Ystart, Ystart(room, 2)) 'specific_heat("CO", Ystart(room, 2)) * Ystart(room, 8) + specific_heat("CO2", Ystart(room, 2)) * Ystart(room, 10) + specific_heat("H2O", Ystart(room, 2)) * Ystart(room, 16) + specific_heat("O2", Ystart(room, 2)) * Ystart(room, 5) + specific_heat("N2", Ystart(room, 2)) * (1 - Ystart(room, 5) - Ystart(room, 16) - Ystart(room, 8) - Ystart(room, 10))
                    cplx(room) = specific_heat_lower(room, Ystart, Ystart(room, 3)) 'specific_heat("CO", Ystart(room, 3)) * Ystart(room, 9) + specific_heat("CO2", Ystart(room, 3)) * Ystart(room, 11) + specific_heat("H2O", Ystart(room, 3)) * Ystart(room, 17) + specific_heat("O2", Ystart(room, 3)) * Ystart(room, 15) + specific_heat("N2", Ystart(room, 3)) * (1 - Ystart(room, 15) - Ystart(room, 17) - Ystart(room, 9) - Ystart(room, 11))

                Next room

                If i > 0 Then Call cvent_open(i) '2014.16 FR ceiling vents

                If rcnone = False Then 'tag B
                    'room lining fire growth model applies

                    If burner_id = 0 Then burner_id = 1 '2011.26 moved

                    '**********************
                    'FLOOR IGNITION
                    '**********************
                    For room = 1 To NumberRooms
                        If FloorIgniteFlag(room) = 0 Then 'floor has not ignited
                            '----------------------------------------------------------------------
                            If FloorFTP(room) > 0 And IgnCorrelation = vbFTP Then
                                'use flux time product method to determine ignition time
                                If room = fireroom Then
                                    'If frmQuintiere.chkFloorFlameSpread.Value = vbChecked Then
                                    If frmQuintiere.optNoneFloor.Checked = False Then
                                        IncidentFlux = QFB - QFloor(room, i) 'kW/m2  QFB same as for wall adjacent burner
                                    Else
                                        'combustible floor can ignite near flashover due to heating by hot layer, but flame spread effects ignored
                                        IncidentFlux = Target(room, i) 'kW/m2
                                    End If
                                Else
                                    'If frmQuintiere.chkFloorFlameSpread.Value = vbChecked Then

                                    'IncidentFlux = QPyrol - QFloor(room, i)  'kW/m2  QFB same as for wall adjacent burner
                                    'Else
                                    IncidentFlux = Target(room, i) 'kW/m2
                                    'End If
                                End If
                                If i = 1 Then
                                    If IncidentFlux >= FloorQCrit(room) Then
                                        FloorFTP_sum(room, i) = (IncidentFlux - FloorQCrit(room)) ^ Floorn(room) * Timestep
                                    End If
                                Else
                                    If IncidentFlux >= FloorQCrit(room) Then
                                        FloorFTP_sum(room, i) = FloorFTP_sum(room, i - 1) + (IncidentFlux - FloorQCrit(room)) ^ Floorn(room) * Timestep
                                    Else
                                        FloorFTP_sum(room, i) = FloorFTP_sum(room, i - 1)
                                    End If
                                End If

                                'test to see if floor ignited in next room
                                If FireLocation(burner_id) = 0 Then
                                    'fire in centre of room
                                    If Yf_pyrolysis(fireroom, i) >= 0.5 * RoomLength(fireroom) Then
                                        FloorIgniteFlag(room) = 1
                                    End If
                                ElseIf FireLocation(burner_id) = 1 Then
                                    'fire against wall
                                    If Yf_pyrolysis(fireroom, i) >= RoomLength(fireroom) Then
                                        FloorIgniteFlag(room) = 1
                                    End If
                                Else
                                End If

                                If (FloorIgniteFlag(room) = 1 Or FloorFTP_sum(room, i) > FloorFTP(room)) And FloorEffectiveHeatofCombustion(room) > 0 Then 'floor ignited
                                    FloorIgniteFlag(room) = 1
                                    MDIFrmMain.ToolStripStatusLabel2.Text = "Floor in Room " & CStr(room) & " has ignited at " & CStr(tim(i, 1)) & " seconds."
                                    FloorIgniteTime(room) = tim(i, 1)
                                    FloorIgniteStep(room) = i

                                    If room = fireroom Then
                                        If frmQuintiere.optWindAided.Checked = True Then
                                            If Yf_pyrolysis(room, i) < BurnerFlameLength Then
                                                Yf_pyrolysis(room, i) = BurnerFlameLength
                                            End If
                                            Xf_pyrolysis(room, i) = BurnerWidth / 2 'lateral pyrolysis measured from the centreline of the burner
                                            Yf_burnout(room, i) = 0
                                        ElseIf frmQuintiere.optOpposedFlow.Checked = True Then
                                            Yf_burnout(room, i) = 0
                                            Yf_pyrolysis(room, i) = 0
                                            Xf_pyrolysis(room, i) = BurnerDiameter / 2
                                        Else
                                            Yf_burnout(room, i) = 0
                                            Yf_pyrolysis(room, i) = 0
                                            Xf_pyrolysis(room, i) = 0
                                        End If
                                    Else
                                        If NumberVents(fireroom, room) > 0 Then
                                            'Yf_pyrolysis(room, i) = 0.1 'm ignite a strip 100 mm wide
                                            Yf_pyrolysis(room, i) = 0.5 * BurnerWidth 'm ignite a strip 100 mm wide
                                            Xf_pyrolysis(room, i) = Xf_pyrolysis(fireroom, i)
                                            If 2 * Xf_pyrolysis(room, i) > VentWidth(fireroom, room, 1) Then Xf_pyrolysis(room, i) = (VentWidth(fireroom, room, 1)) / 2
                                            Yf_burnout(room, i) = -(Yf_pyrolysis(fireroom, i) - Yf_burnout(fireroom, i))
                                            'Yf_burnout(room, i) = 0
                                        End If
                                    End If

                                    Xstart(room, 4) = Yf_pyrolysis(room, i)
                                    Xstart(room, 6) = Xf_pyrolysis(room, i)
                                    Xstart(room, 5) = Yf_burnout(room, i)
                                End If

                            End If
                            '----------------------------------------------------------------------

                        End If
                    Next room


                    '**********************
                    'CEILING IGNITION 
                    '**********************
                    If CeilingIgniteFlag(fireroom) = 0 And CeilingEffectiveHeatofCombustion(fireroom) > 0 Then 'tag C
                        'ceiling has not ignited

                        'incident flux on the ceiling, SFPE h/b 3rd ed. page 2-272
                        '9June2003
                        'for a corner fire
                        If FireLocation(1) = 2 Then
                            Lf = BurnerWidth * 5.9 * Sqrt(Qburner / 1110 / BurnerWidth ^ (5 / 2))
                            If (RoomHeight(fireroom) - FireHeight(1)) / Lf <= 0.52 Then
                                IncidentFlux = 120
                            Else
                                IncidentFlux = 13 * ((RoomHeight(fireroom) - FireHeight(1)) / Lf) ^ (-3.5)
                            End If
                            IncidentFlux = IncidentFlux - QCeiling(fireroom, i)

                        ElseIf FireLocation(1) = 1 Then
                            Lf = 0.23 * Qburner ^ (2 / 5) - 1.02 * BurnerWidth

                            IncidentFlux = 200 * (1 - Exp(-0.09 * Qburner ^ (1 / 3))) 'if the flame reaches the ceiling use this

                            'this was previously changes to be as below, not sure why?
                            'If Lf > 0 Then
                            '    If (RoomHeight(fireroom) - FireHeight(1)) / Lf <= 0.4 Then
                            '        IncidentFlux = IncidentFlux - QCeiling(fireroom, i)
                            '    ElseIf (RoomHeight(fireroom) - FireHeight(1)) / Lf <= 1 Then 'flame reaches the ceiling
                            '        IncidentFlux = IncidentFlux - 5 / 3 * ((RoomHeight(fireroom) - FireHeight(1)) / Lf - 2 / 5) * (IncidentFlux - 20) - QCeiling(fireroom, i)
                            '    Else
                            '        IncidentFlux = 20 * ((RoomHeight(fireroom) - FireHeight(1)) / Lf) ^ (-5 / 3) - QCeiling(fireroom, i)
                            '    End If
                            'End If

                            If Lf < RoomHeight(fireroom) - FireHeight(1) Then
                                'flame not touching the ceiling
                                If Lf > 0 Then IncidentFlux = 20 * ((RoomHeight(fireroom) - FireHeight(1)) / Lf) ^ (-5 / 3) - QCeiling(fireroom, i)
                            Else
                                'flame is touching the ceiling
                                IncidentFlux = IncidentFlux - QCeiling(fireroom, i)   'kW/m2
                            End If
                        Else
                            'burner in centre of room
                            IncidentFlux = 0.28 * Qburner ^ (5 / 6) * (RoomHeight(fireroom) - FireHeight(1)) ^ (-7 / 3)
                            'eqn 15b Fire Technology Vol 21 No 4 p267 - mostly convective
                            IncidentFlux = IncidentFlux - QCeiling(fireroom, i)
                        End If

                        If CeilingFTP(fireroom) > 0 And IgnCorrelation = vbFTP Then
                            If i = 1 Then
                                If IncidentFlux >= CeilingQCrit(fireroom) Then
                                    CeilingFTP_sum(i) = (IncidentFlux - CeilingQCrit(fireroom)) ^ Ceilingn(fireroom) * Timestep
                                End If
                            Else
                                If IncidentFlux >= CeilingQCrit(fireroom) Then
                                    CeilingFTP_sum(i) = CeilingFTP_sum(i - 1) + (IncidentFlux - CeilingQCrit(fireroom)) ^ Ceilingn(fireroom) * Timestep
                                Else
                                    CeilingFTP_sum(i) = CeilingFTP_sum(i - 1)
                                End If
                            End If

                            'test to see if ceiling ignited
                            If CeilingFTP_sum(i) > CeilingFTP(fireroom) And CeilingEffectiveHeatofCombustion(fireroom) > 0 Then 'ceiling ignited
                                'If (IgCeilingNode(fireroom, 1, i) >= IgTempC(fireroom) Or Y_pyrolysis(fireroom, i) > RoomHeight(fireroom)) And CeilingEffectiveHeatofCombustion(fireroom) > 0 Then
                                CeilingIgniteFlag(fireroom) = 1

                                Dim Message As String = CStr(tim(i, 1)) & " sec. Ceiling in Room " & CStr(fireroom) & " ignited."
                                MDIFrmMain.ToolStripStatusLabel2.Text = Message.ToString
                                If ProjectDirectory = RiskDataDirectory Then
                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                                End If

                                CeilingIgniteTime(fireroom) = tim(i, 1)
                                CeilingIgniteStep(fireroom) = i
                                If RoomHeight(fireroom) + 2 * BurnerWidth >= Y_pyrolysis(fireroom, i) Then
                                    'why 2 x burner width here?
                                    Y_pyrolysis(fireroom, i) = RoomHeight(fireroom) + 2 * BurnerWidth
                                    '25/6/15
                                    'Y_pyrolysis(fireroom, i) = RoomHeight(fireroom) + BurnerWidth
                                End If
                                If FireHeight(1) > Y_burnout(fireroom, i) Then
                                    Y_burnout(fireroom, i) = FireHeight(1)
                                End If
                                If FireLocation(1) = 0 Then
                                    'fire in centre of room, wall not involved
                                    Y_burnout(fireroom, i) = RoomHeight(fireroom)
                                End If
                                If WallIgniteFlag(fireroom) = 0 Then
                                    'ceiling is ignited but wall not involved
                                    Y_burnout(fireroom, i) = RoomHeight(fireroom)
                                End If
                                Xstart(fireroom, 1) = Y_pyrolysis(fireroom, i)
                                Xstart(fireroom, 3) = Y_burnout(fireroom, i)
                            End If
                            'Else
                        ElseIf CeilingFTP(fireroom) > 0 Then  '27/4/05  not using FTP method
                            If CeilingEffectiveHeatofCombustion(fireroom) > 0 Then 'only do this if we have a combustible ceiling 12/3/2003

                                'FTP data not available
                                'ceiling has not ignited
                                If HaveCeilingSubstrate(fireroom) = False Then
                                    'one layer ceiling
                                    Call Ceiling_Ignite((i), IgCeilingNode)
                                Else
                                    'two layer ceiling
                                    Call Ceiling_Ignite2((i), IgCeilingNode)
                                End If
                                'test to see if ceiling ignited
                                If (IgCeilingNode(fireroom, 1, i) >= IgTempC(fireroom) Or Y_pyrolysis(fireroom, i) > RoomHeight(fireroom)) And CeilingEffectiveHeatofCombustion(fireroom) > 0 Then

                                    CeilingIgniteFlag(fireroom) = 1
                                    CeilingIgniteTime(fireroom) = tim(i, 1)
                                    'MDIFrmMain.ToolStripStatusLabel2.Text = "Ceiling in Room " & CStr(fireroom) & " has ignited at " & CStr(tim(i, 1)) & " seconds."

                                    Dim Message As String = CStr(tim(i, 1)) & " sec. Ceiling in Room " & CStr(fireroom) & " ignited."
                                    MDIFrmMain.ToolStripStatusLabel2.Text = Message.ToString
                                    If ProjectDirectory = RiskDataDirectory Then
                                        frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                                    End If

                                    CeilingIgniteStep(fireroom) = i
                                    If RoomHeight(fireroom) + BurnerWidth >= Y_pyrolysis(fireroom, i) Then
                                        Y_pyrolysis(fireroom, i) = RoomHeight(fireroom) + BurnerWidth
                                    End If
                                    Xstart(fireroom, 1) = Y_pyrolysis(fireroom, i)
                                End If
                            End If
                        End If
                    End If 'tag C

                    For room = 1 To NumberRooms
                        If room <> fireroom Then 'tag D
                            'ignition of wall and ceiling linings in adjacent room
                            If IgniteNextRoom = True And (CeilingEffectiveHeatofCombustion(room) > 0 Or WallEffectiveHeatofCombustion(room) > 0 Or FloorEffectiveHeatofCombustion(room) > 0) Then
                                'find largest vent
                                'If room <> fireroom Then
                                'where more than 1 vent connects the two rooms, the vent with the largest area is used
                                avent = 0
                                maxarea = 0
                                maxvent = 0
                                If fireroom < room Then
                                    For vent = 1 To NumberVents(fireroom, room)
                                        avent = VentHeight(fireroom, room, vent) * VentWidth(fireroom, room, vent)
                                        If (tim(i, 1) - VentOpenTime(fireroom, room, vent) < 2) And VentOpenTime(fireroom, room, vent) > 0 Then
                                            avent = avent * (tim(i, 1) - VentOpenTime(fireroom, room, vent)) / 2
                                        End If
                                        If tim(i, 1) > VentCloseTime(fireroom, room, vent) Then
                                            If tim(i, 1) - VentCloseTime(fireroom, room, vent) < 2 And VentCloseTime(fireroom, room, vent) > VentOpenTime(fireroom, room, vent) + 2 Then
                                                avent = avent * (tim(i, 1) - VentCloseTime(fireroom, room, vent)) / 2
                                            End If
                                        End If
                                        If avent > maxarea Then
                                            maxarea = avent
                                            maxvent = vent
                                        End If
                                    Next vent
                                ElseIf fireroom > room Then
                                    For vent = 1 To NumberVents(room, fireroom)
                                        avent = VentHeight(room, fireroom, vent) * VentWidth(room, fireroom, vent)
                                        If (tim(i, 1) - VentOpenTime(room, fireroom, vent) < 2) And VentOpenTime(room, fireroom, vent) > 0 Then
                                            avent = avent * (tim(i, 1) - VentOpenTime(room, fireroom, vent)) / 2
                                        End If
                                        If tim(i, 1) > VentCloseTime(room, fireroom, vent) Then
                                            If tim(i, 1) - VentCloseTime(room, fireroom, vent) < 2 And VentCloseTime(room, fireroom, vent) > VentOpenTime(room, fireroom, vent) + 2 Then
                                                avent = avent * (tim(i, 1) - VentCloseTime(room, fireroom, vent)) / 2
                                            End If
                                        End If
                                        If avent > maxarea Then
                                            maxarea = avent
                                            maxvent = vent
                                        End If
                                    Next vent
                                End If
                                vwidth = 0
                                If fireroom < room Then
                                    If maxvent > 0 Then vwidth = VentWidth(fireroom, room, maxvent)
                                Else
                                    If maxvent > 0 Then vwidth = VentWidth(room, fireroom, maxvent)
                                End If

                                'here
                                '=========================
                                'put floor spread in here
                                'burner in centre of room

                                fwidth = vwidth

                                '==========================
                                If vwidth > 0 Then
                                    If CeilingEffectiveHeatofCombustion(room) > 0 Or WallEffectiveHeatofCombustion(room) > 0 Then
                                        If fireroom < room Then
                                            'take width of flame as distance of layerheight below vent soffit
                                            If VentHeight(fireroom, room, maxvent) + VentSillHeight(fireroom, room, maxvent) > layerheight(fireroom, i) Then
                                                projection = VentHeight(fireroom, room, maxvent) + VentSillHeight(fireroom, room, maxvent) - layerheight(fireroom, i)
                                            End If
                                            'wall flame height using Delichatsios correlation
                                            If i > 1 Then QL = (ventfire(room, i - 1) - SpecificHeat_air * uppertemp(fireroom, i - 1) * FlowToUpper(fireroom, i - 1)) / vwidth 'kW/m
                                            'If i > 1 Then QL = (ventfire(room, i - 1)) / vwidth 'kW/m
                                            QLstar = QL / (ReferenceTemp * ReferenceDensity * SpecificHeat_air * Sqrt(G) * vwidth ^ (3 / 2))
                                            fheight = 2.8 * QLstar ^ (2 / 3) * vwidth
                                            If projection <> 0 And fheight <> 0 Then
                                                eFLame = 1 - Exp(-EmissionCoefficient * projection)
                                                'eFLame = 1 - Exp(-0.3 * projection)
                                                'awallflux_r = RadiantLossFraction * QL * vwidth / 2 / ((vwidth * fheight)) 'kW/m2
                                                awallflux_r = ObjectRLF(1) * QL * vwidth / 2 / ((vwidth * fheight)) 'kW/m2
                                                If ventfire(room, i - 1) < 10 Then
                                                    Tflame = uppertemp(fireroom, i - 1)
                                                Else
                                                    Tflame = Flametemp
                                                End If
                                            End If
                                            If fheight <= 0 Then
                                                fheight = 0
                                                Tflame = lowertemp(room, 1)
                                                eFLame = 0
                                            End If
                                        End If

                                        If fireroom > room Then
                                            'take width of flame as distance of layerheight below vent soffit
                                            If VentHeight(room, fireroom, maxvent) + VentSillHeight(room, fireroom, maxvent) > layerheight(fireroom, i) Then
                                                projection = VentHeight(room, fireroom, maxvent) + VentSillHeight(room, fireroom, maxvent) - layerheight(fireroom, i)
                                            End If
                                            'wall flame height using Delichatsios correlation
                                            QL = (ventfire(room, i - 1)) / vwidth 'kW/m
                                            If i > 1 Then QL = (ventfire(room, i - 1) - SpecificHeat_air * uppertemp(fireroom, i - 1) * FlowToUpper(fireroom, i - 1)) / vwidth 'kW/m
                                            QLstar = QL / (ReferenceTemp * ReferenceDensity * SpecificHeat_air * Sqrt(G) * vwidth ^ (3 / 2))
                                            fheight = 2.8 * QLstar ^ (2 / 3) * vwidth
                                            If projection <> 0 And fheight <> 0 Then
                                                eFLame = 1 - Exp(-EmissionCoefficient * projection)
                                                'awallflux_r = RadiantLossFraction * QL * vwidth / 2 / ((vwidth * fheight)) 'kW/m2
                                                awallflux_r = ObjectRLF(1) * QL * vwidth / 2 / ((vwidth * fheight)) 'kW/m2
                                                If ventfire(room, i - 1) < 10 Then
                                                    Tflame = uppertemp(fireroom, i - 1)
                                                Else
                                                    Tflame = Flametemp
                                                End If
                                            End If
                                            If fheight <= 0 Then
                                                fheight = 0
                                                Tflame = lowertemp(room, 1)
                                                eFLame = 0
                                            End If
                                        End If

                                        If HaveWallSubstrate(room) = False Then
                                            'one layer wall
                                            Call Wall_Ignite_Adj1(layerz, vwidth, fheight, Tflame, eFLame, room, (i), IgWallNode, flux_wall_AR, awallflux_r)
                                        Else
                                            'two layer wall
                                            Call Wall_Ignite_Adj2(layerz, vwidth, fheight, Tflame, eFLame, room, (i), IgWallNode, flux_wall_AR, awallflux_r)
                                        End If

                                        wallparam(1, i) = flux_wall_AR
                                        wallparam(2, i) = fheight
                                        wallparam(3, i) = Tflame

                                        If fireroom < room Then
                                            fheight2 = VentSillHeight(fireroom, room, maxvent) + VentHeight(fireroom, room, maxvent) + fheight
                                        Else
                                            fheight2 = VentSillHeight(room, fireroom, maxvent) + VentHeight(fireroom, fireroom, maxvent) + fheight
                                        End If

                                        IncidentFlux = flux_wall_AR

                                        If i = 1 Then
                                            If IncidentFlux >= WallQCrit(room) Then
                                                WallFTP_sum_ar(i) = (IncidentFlux - WallQCrit(room)) ^ Walln(room) * Timestep
                                            End If
                                        Else
                                            If IncidentFlux >= WallQCrit(room) Then
                                                WallFTP_sum_ar(i) = WallFTP_sum_ar(i - 1) + (IncidentFlux - WallQCrit(room)) ^ Walln(room) * Timestep
                                            End If
                                        End If

                                        If WallIgniteFlag(room) = 0 Then
                                            'test to see if lining ignited
                                            If WallFTP_sum_ar(i) > WallFTP(room) And WallEffectiveHeatofCombustion(room) > 0 Then 'wall ignited
                                                WallIgniteFlag(room) = 1

                                                Dim Message As String = CStr(tim(i, 1)) & " sec. Wall in Room " & CStr(room) & " ignited."
                                                MDIFrmMain.ToolStripStatusLabel2.Text = Message.ToString
                                                If ProjectDirectory = RiskDataDirectory Then
                                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                                                End If

                                                WallIgniteTime(room) = tim(i, 1)
                                                WallIgniteStep(room) = i
                                                If CeilingIgniteFlag(room) = 0 Then
                                                    Y_pyrolysis(room, i) = fheight2
                                                    X_pyrolysis(room, i) = vwidth / 2
                                                    For k = 1 To i - 1
                                                        Y_pyrolysis(room, k) = VentSillHeight(fireroom, room, maxvent) + VentHeight(fireroom, room, maxvent)
                                                    Next k
                                                    Xstart(room, 1) = Y_pyrolysis(room, i)
                                                    Xstart(room, 2) = X_pyrolysis(room, i)
                                                    Xstart(room, 3) = Y_burnout(room, i)
                                                End If
                                            End If
                                        End If

                                        If CeilingIgniteFlag(room) = 0 Then
                                            'call heat conduction routine
                                            If HaveCeilingSubstrate(room) = False Then
                                                'one layer ceiling
                                                If i > 1 Then
                                                    Call Ceiling_Ignite_Adj1(projection, room, (i), IgCeilingNode, layerheight(fireroom, i), ventfire(room, i - 1), RoomHeight(room), vwidth)
                                                Else
                                                    Call Ceiling_Ignite_Adj1(projection, room, (i), IgCeilingNode, layerheight(fireroom, i), 0, RoomHeight(room), vwidth)
                                                End If
                                            Else
                                                'two layer ceiling
                                                If i > 1 Then
                                                    Call Ceiling_Ignite_Adj2(projection, room, (i), IgCeilingNode, layerheight(fireroom, i), ventfire(room, i - 1), RoomHeight(room), vwidth)
                                                Else
                                                    Call Ceiling_Ignite_Adj2(projection, room, (i), IgCeilingNode, layerheight(fireroom, i), 0, RoomHeight(room), vwidth)
                                                End If
                                            End If

                                            'test to see if ceiling ignited
                                            If IgCeilingNode(room, 1, i) >= IgTempC(room) And CeilingEffectiveHeatofCombustion(room) > 0 Then
                                                CeilingIgniteFlag(room) = 1
                                                CeilingIgniteTime(room) = tim(i, 1)
                                                'MDIFrmMain.ToolStripStatusLabel2.Text = "Ceiling in Room " & CStr(room) & " has ignited at " & CStr(tim(i, 1)) & " seconds."

                                                Dim Message As String = CStr(tim(i, 1)) & " sec. Ceiling in Room " & CStr(room) & " has ignited."
                                                MDIFrmMain.ToolStripStatusLabel2.Text = Message.ToString
                                                If ProjectDirectory = RiskDataDirectory Then
                                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                                                End If
                                                CeilingIgniteStep(room) = i
                                                Y_pyrolysis(room, i) = RoomHeight(room) + vwidth / 2 'assume a semi-circle centred on the vetn opening with radius = 1/2 x vent width
                                                For k = 1 To i - 1
                                                    Y_pyrolysis(room, k) = RoomHeight(room)
                                                Next k
                                                Xstart(room, 1) = Y_pyrolysis(room, i)
                                            End If
                                        End If
                                    End If

                                    If WallIgniteFlag(room) = 1 Or CeilingIgniteFlag(room) = 1 Or FloorIgniteFlag(room) = 1 Then
                                        'adjacent room has ignited
                                        'get the max hrr
                                        HeatRelease(room, i, 1) = HRR_nextroom(vwidth, room, tim(i, 1), flux_wall_AR)
                                    End If
                                End If
                                'End If
                            End If
                        End If 'tag D
                    Next room

                    'call floor_ignition


                    '**********************
                    'WALL IGNITION
                    '**********************
                    'if using quintiere's room corner model then also update wall temperatures adjacent burner
                    'only call this procedure if the wall has not ignited
                    'If quintiere = True And WallIgniteFlag(fireroom) = 0 And FireLocation(1) <> 0 Then 'tag E
                    If quintiere = True And WallIgniteFlag(fireroom) = 0 Then 'tag E

                        If WallFTP(fireroom) = 0 Or IgnCorrelation = vbJanssens Then
                            'we have to look at surface temperatures if FTP data not available
                            If HaveWallSubstrate(fireroom) = False Then
                                'one layer wall
                                Call Wall_Ignite((i), IgWallNode)
                            Else
                                'two layer wall
                                Call Wall_Ignite2((i), IgWallNode, IncidentFlux)
                            End If
                        Else ' FTP method

                            IncidentFlux = QFB - QUpperWall(fireroom, i) 'kW/m2   'added 9June2003
                            'QFB includes convection and radiation
                            If i = 1 Then
                                If IncidentFlux >= WallQCrit(fireroom) Then
                                    WallFTP_sum(i) = (IncidentFlux - WallQCrit(fireroom)) ^ Walln(fireroom) * Timestep
                                End If
                            Else
                                If IncidentFlux >= WallQCrit(fireroom) Then
                                    WallFTP_sum(i) = WallFTP_sum(i - 1) + (IncidentFlux - WallQCrit(fireroom)) ^ Walln(fireroom) * Timestep
                                End If
                            End If
                        End If

                        'test to see if wall ignited
                        If WallFTP(fireroom) > 0 And IgnCorrelation = vbFTP Then 'tag F
                            'this only applicable where FTP method used and ignition data provided
                            If WallFTP_sum(i) > WallFTP(fireroom) And WallEffectiveHeatofCombustion(fireroom) > 0 Then 'wall ignited
                                WallIgniteFlag(fireroom) = 1
                                'If burner_id = 0 Then burner_id = 1


                                Dim Message As String = CStr(tim(i, 1)) & " sec. Wall in Room " & CStr(fireroom) & " has ignited."
                                MDIFrmMain.ToolStripStatusLabel2.Text = Message.ToString
                                If ProjectDirectory = RiskDataDirectory Then
                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                                    BurnerWidth = (ObjLength(1) + ObjWidth(1)) / 2
                                End If
                                WallIgniteTime(fireroom) = tim(i, 1)
                                WallIgniteStep(fireroom) = i

                                'CW                         'BurnerFlameLength = Get_BurnerFlameLength(Qburner)
                                'If frmQuintiere.chkFelix.Value = vbChecked Then BurnerFlameLength = CSng(frmQuintiere.txtFlameHeight.Text)

                                If Y_pyrolysis(fireroom, i) < 0.4 * BurnerFlameLength + FireHeight(burner_id) Then
                                    Y_pyrolysis(fireroom, i) = 0.4 * BurnerFlameLength + FireHeight(burner_id)
                                End If
                                X_pyrolysis(fireroom, i) = BurnerWidth
                                If Y_burnout(fireroom, i) < FireHeight(burner_id) Then
                                    Y_burnout(fireroom, i) = FireHeight(burner_id)
                                End If
                                Xstart(fireroom, 1) = Y_pyrolysis(fireroom, i)
                                Xstart(fireroom, 2) = X_pyrolysis(fireroom, i)
                                Xstart(fireroom, 3) = Y_burnout(fireroom, i)
                            End If
                        Else
                            'FTP data not available
                            If IgWallNode(fireroom, 1, i) >= IgTempW(fireroom) And WallEffectiveHeatofCombustion(fireroom) > 0 Then
                                WallIgniteFlag(fireroom) = 1
                                'MDIFrmMain.ToolStripStatusLabel2.Text = "Wall in Room " & CStr(fireroom) & " has ignited at " & CStr(tim(i, 1)) & " seconds."

                                Dim Message As String = CStr(tim(i, 1)) & " sec. Wall in Room " & CStr(fireroom) & " has ignited."
                                MDIFrmMain.ToolStripStatusLabel2.Text = Message.ToString
                                If ProjectDirectory = RiskDataDirectory Then
                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                                End If
                                WallIgniteTime(fireroom) = tim(i, 1)
                                WallIgniteStep(fireroom) = i

                                BurnerFlameLength = Get_BurnerFlameLength(1, Qburner)

                                'If frmQuintiere.chkFelix.Value = vbChecked Then BurnerFlameLength = CSng(frmQuintiere.txtFlameHeight.Text)
                                '%%% chow
                                Y_pyrolysis(fireroom, i) = 0.4 * BurnerFlameLength + FireHeight(1)
                                X_pyrolysis(fireroom, i) = BurnerWidth
                                Y_burnout(fireroom, i) = FireHeight(burner_id)
                                Xstart(fireroom, 1) = Y_pyrolysis(fireroom, i)
                                Xstart(fireroom, 2) = X_pyrolysis(fireroom, i)
                                Xstart(fireroom, 3) = Y_burnout(fireroom, i)
                            End If
                        End If 'tag F
                    End If 'tag E
                End If 'tag B
            End If 'tag A

            'code to deal with ignition of secondary items
            ' If NumberObjects > 1 And frmInputs.chkIgniteTargets.Checked = True And frmInputs.chkFireMode.Checked = True Then

            If ignitetargets = True Then
                'Call secondary_targets(rcnone, i, itcounter, ItemFTP_sum_pilot, ItemFTP_sum_auto, ItemFTP_sum_wall, ItemFTP_sum_ceiling)
                Call secondary_targets_ISD(rcnone, i, itcounter, ItemFTP_sum_pilot, ItemFTP_sum_auto, ItemFTP_sum_wall, ItemFTP_sum_ceiling)
                If rcnone = False And WallIgniteFlag(fireroom) = 1 Then
                    Xstart(fireroom, 1) = Y_pyrolysis(fireroom, i)
                    Xstart(fireroom, 2) = X_pyrolysis(fireroom, i)
                    Xstart(fireroom, 3) = Y_burnout(fireroom, i)
                End If
            End If

            'call the ODE solver routine for the zone model
            Call ODE_Solver3(Ystart)
            'Call ODE_Solver_NMath(Ystart)

            If NumSprinklers > 0 Then
                'call the ODE solver for sprinkler devices only
                Call ODE_Sprinkler_Solver(oSprinklers, SPstart)
            End If

            'terminate the simulate early
            If flagstop = 1 Then
                'copied from 99: above, instead of goto 99
                If i - 2 > 1 Then
                    NumberTimeSteps = i - 2
                Else
                    NumberTimeSteps = 1
                End If
                MDIFrmMain.ToolStripProgressBar1.Value = 100
                MDIFrmMain.ToolStripStatusLabel3.Text = " "
                TwoZones(fireroom) = True

                Exit For
            End If
            'returned values
            If i <> NumberTimeSteps + 1 Then
                tim(i + 1, 1) = Round(tim(i, 1) + Timestep, 3)
                'tim(i + 1, 1) = tim(i, 1) + Timestep

                j = 1
                For Each oSprinkler In oSprinklers
                    SprinkTemp(j, i + 1) = SPstart(j)
                    If SprinkTemp(j, i) - 273 > oSprinkler.acttemp And oSprinkler.responsetime = 0 Then

                        oSprinkler.responsetime = tim(i, 1)
                        Dim message As String = ""
                        'don't count heat detectors
                        If oSprinkler.cfactor > 0 Or oSprinkler.sprdensity > 0 Then
                            zcount = zcount + 1
                            message = Format(tim(i, 1), "0") & " Sec. Sprinkler " & oSprinkler.sprid & " responded."
                        Else
                            message = Format(tim(i, 1), "0") & " Sec. Heat Detector " & oSprinkler.sprid & " responded."
                            If HDFlag = 0 Then
                                HDFlag = 1
                                HDTime = tim(i, 1) 'time of actuation of hd 
                            End If
                        End If

                        frmEndPoints.lblSprinkler.Text = Message.ToString
                        frmInputs.rtb_log.Text = message.ToString & Chr(13) & frmInputs.rtb_log.Text


                        If SprinklerFlag = 0 And zcount >= NumOperatingSpr Then

                            SprinklerFlag = 1
                            SprinklerTime = tim(i, 1) 'time of actuation of last needed operating sprinkler 
                            SprinklerHRR = HeatRelease(fireroom, i, 2)
                            WaterSprayDensity(fireroom) = oSprinkler.sprdensity

                            Select Case sprink_mode
                                Case 0
                                    message = "Fire HRR not affected by sprinkler"
                                Case 1
                                    message = "Fire HRR is controlled by sprinkler"
                                Case 2
                                    message = "Fire HRR is suppressed by sprinkler"
                            End Select
                            frmEndPoints.lblSprinkler.Text = message.ToString
                            frmInputs.rtb_log.Text = message.ToString & Chr(13) & frmInputs.rtb_log.Text

                        End If

                    End If
                    j = j + 1
                Next

                For j = 1 To NumberRooms
                    UpperVolume(j, i + 1) = Ystart(j, 1)

                    For b = 1 To 19
                        If Single.IsNaN(Ystart(j, b)) = True Then
                            'Stop
                            Dim Message As String = tim(i, 1).ToString & " sec. Room " & j.ToString & " Convergence error. Run Terminated."
                            frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                            Ystart(j, 1) = UpperVolume(j, i)
                            Ystart(j, 2) = uppertemp(j, i)
                            Ystart(j, 3) = lowertemp(j, i)
                            Ystart(j, 4) = RoomPressure(j, i)
                            Ystart(j, 5) = O2MassFraction(j, i, 1)
                            Ystart(j, 6) = TUHC(j, i, 1)
                            Ystart(j, 7) = TUHC(j, i, 2)
                            Ystart(j, 8) = COMassFraction(j, i, 1)
                            Ystart(j, 9) = COMassFraction(j, i, 2)
                            Ystart(j, 10) = CO2MassFraction(j, i, 1)
                            Ystart(j, 11) = CO2MassFraction(j, i, 2)
                            Ystart(j, 12) = SootMassFraction(j, i, 1)
                            Ystart(j, 13) = SootMassFraction(j, i, 2)
                            Ystart(j, 14) = HCNMassFraction(j, i, 2)
                            Ystart(j, 15) = O2MassFraction(j, i, 2)
                            Ystart(j, 16) = H2OMassFraction(j, i, 1)
                            Ystart(j, 17) = H2OMassFraction(j, i, 2)
                            Ystart(j, 18) = HCNMassFraction(j, i, 1)
                            Ystart(j, 19) = LinkTemp(j, i)

                            flagstop = 1
                            Exit For

                            'Oops - not a number
                        Else
                            'Got a number
                        End If
                    Next

                    If TwoZones(j) = True Then
                        Call Layer_Height(UpperVolume(j, i + 1), layerz, j)
                    Else
                        layerz = 0.1
                        'layerz = 0.005
                        ' Ystart(j, 3) = Ystart(j, 2) 'make the LL = UL temp for single zone
                    End If

                    'Call Layer_Height(UpperVolume(j, i + 1), layerz, j)
                    layerheight(j, i + 1) = layerz

                    If Ystart(j, 2) < Min(InteriorTemp, ExteriorTemp) Then
                        Ystart(j, 2) = Min(InteriorTemp, ExteriorTemp)
                    End If
                    uppertemp(j, i + 1) = Ystart(j, 2)
                   
                    If Ystart(j, 3) < Min(InteriorTemp, ExteriorTemp) Then
                        Ystart(j, 3) = Min(InteriorTemp, ExteriorTemp)
                    End If

                    If Ystart(j, 3) > Ystart(j, 2) Then
                        Ystart(j, 3) = Ystart(j, 2) 'do not permit LL temp > UL temp
                    End If

                    lowertemp(j, i + 1) = Ystart(j, 3)

                    LinkTemp(j, i + 1) = Ystart(j, 19) '%%% change this to save temp for each device, see sprinktemp

                    COMassFraction(j, i + 1, 1) = Ystart(j, 8)

                    If Ystart(j, 10) > 1 Then
                        Dim Message As String = CStr(tim(i, 1)) & " sec. Warning - there is a numerical problem in room " & CStr(j) & ". Try using a smaller timestep."
                        frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                    End If

                    CO2MassFraction(j, i + 1, 1) = Ystart(j, 10)

                    SootMassFraction(j, i + 1, 1) = Ystart(j, 12)
                    'If SootMassFraction(j, i + 1, 1) < 0 Then Stop
                    O2MassFraction(j, i + 1, 1) = Ystart(j, 5)
                    If O2MassFraction(j, i + 1, 1) < 0 Then O2MassFraction(j, i + 1, 1) = 0
                    H2OMassFraction(j, i + 1, 1) = Ystart(j, 16)
                    If H2OMassFraction(j, i + 1, 1) < 0 Then H2OMassFraction(j, i + 1, 1) = 0
                    HCNMassFraction(j, i + 1, 1) = Ystart(j, 18)
                    RoomPressure(j, i + 1) = Ystart(j, 4)
                    COMassFraction(j, i + 1, 2) = Ystart(j, 9)
                    CO2MassFraction(j, i + 1, 2) = Ystart(j, 11)
                    SootMassFraction(j, i + 1, 2) = Ystart(j, 13)

                    If Ystart(j, 15) > O2Ambient Then
                        Ystart(j, 15) = O2Ambient
                    End If
                    O2MassFraction(j, i + 1, 2) = Ystart(j, 15)

                    H2OMassFraction(j, i + 1, 2) = Ystart(j, 17)
                    HCNMassFraction(j, i + 1, 2) = Ystart(j, 14)
                    If Ystart(j, 6) < 0 Then Ystart(j, 6) = Smallvalue
                    If Ystart(j, 7) < 0 Then Ystart(j, 7) = Smallvalue

                    TUHC(j, i + 1, 1) = Ystart(j, 6)
                    TUHC(j, i + 1, 2) = Ystart(j, 7)
                Next j
            End If

            'only do if quintiere's room corner model is being used
            If quintiere = True Then
                'solve the flame spread ODE's for the fireroom
                If tim(i, 1) >= WallIgniteTime(fireroom) Or tim(i, 1) >= CeilingIgniteTime(fireroom) Or tim(i, 1) >= FloorIgniteTime(fireroom) Then
                    If tim(i, 1) >= WallIgniteTime(fireroom) Or tim(i, 1) >= CeilingIgniteTime(fireroom) Then
                        X(1) = Xstart(fireroom, 1) 'upward wall + ceiling
                        X(3) = Xstart(fireroom, 3) 'burnout wall + ceiling
                        If tim(i, 1) >= WallIgniteTime(fireroom) Then X(2) = Xstart(fireroom, 2) 'lateral wall
                        V(1) = velocity(fireroom, 1)
                        If tim(i, 1) >= WallIgniteTime(fireroom) Then V(2) = velocity(fireroom, 2)
                        V(3) = velocity(fireroom, 3)
                    End If
                    If tim(i, 1) >= FloorIgniteTime(fireroom) Then
                        X(4) = Xstart(fireroom, 4) 'upward floor
                        X(5) = Xstart(fireroom, 5) 'burnout floor
                        If tim(i, 1) >= FloorIgniteTime(fireroom) Then X(6) = Xstart(fireroom, 6) 'lateral floor
                        V(4) = velocity(fireroom, 4)
                        If tim(i, 1) >= FloorIgniteTime(fireroom) Then V(6) = velocity(fireroom, 6)
                        V(5) = velocity(fireroom, 5)
                    End If
                    '*****************************************************************
                    Call ODE_Spread_Solver(flux_wall_AR, fireroom, fheight, V, X) 'returns arrays velocity and xstart
                    '*****************************************************************

                    Xstart(fireroom, 1) = X(1)
                    Xstart(fireroom, 3) = X(3)
                    If tim(i, 1) >= WallIgniteTime(fireroom) Then Xstart(fireroom, 2) = X(2)
                    velocity(fireroom, 1) = V(1)
                    If tim(i, 1) >= WallIgniteTime(fireroom) Then velocity(fireroom, 2) = V(2)
                    velocity(fireroom, 3) = V(3)

                    Xstart(fireroom, 4) = X(4) 'wind-aided floor spread
                    Xstart(fireroom, 5) = X(5) 'burnout
                    If tim(i, 1) >= FloorIgniteTime(fireroom) Then Xstart(fireroom, 6) = X(6) 'lateral
                    velocity(fireroom, 4) = V(4)
                    If tim(i, 1) >= FloorIgniteTime(fireroom) Then velocity(fireroom, 6) = V(6)
                    velocity(fireroom, 5) = V(5)

                    'returned values
                    If i <> NumberTimeSteps + 1 Then
                        'limit the maximum length of the Y fronts
                        Y_pyrolysis(fireroom, i + 1) = Xstart(fireroom, 1) 'Y pyrolysis front
                        If Y_pyrolysis(fireroom, i + 1) < 0.4 * BurnerFlameLength + Y_pyrolysis(fireroom, 1) Then
                            Y_pyrolysis(fireroom, i + 1) = 0.4 * BurnerFlameLength + Y_pyrolysis(fireroom, 1)
                        End If

                        'If PessimiseCombWall = False Then
                        '    If wallVFSlimit < RoomHeight(room) Then
                        '        If Y_pyrolysis(fireroom, i + 1) > wallVFSlimit Then
                        '            Y_pyrolysis(fireroom, i + 1) = wallVFSlimit 'constrain
                        '        End If
                        '    End If
                        'End If

                        Y_burnout(fireroom, i + 1) = Xstart(fireroom, 3) 'Y burnout front
                        'Xstart(fireroom, 3) = Y_burnout(fireroom, i + 1)
                        FlameVelocity(fireroom, 1, i + 1) = velocity(fireroom, 1) 'upward
                        FlameVelocity(fireroom, 2, i + 1) = velocity(fireroom, 2) 'lateral
                        FlameVelocity(fireroom, 3, i + 1) = velocity(fireroom, 3) 'burnout

                        X_pyrolysis(fireroom, i + 1) = Xstart(fireroom, 2) 'X pyrolysis front
                        If X_pyrolysis(fireroom, i + 1) > RoomLength(fireroom) + RoomWidth(fireroom) Then
                            X_pyrolysis(fireroom, i + 1) = RoomLength(fireroom) + RoomWidth(fireroom)
                        End If

                        If PessimiseCombWall = False Then
                            If X_pyrolysis(fireroom, i + 1) > wallHFSlimit Then
                                X_pyrolysis(fireroom, i + 1) = wallHFSlimit 'constrain
                            End If
                        End If

                        If Y_pyrolysis(fireroom, i + 1) < Y_pyrolysis(fireroom, i) Then Y_pyrolysis(fireroom, i + 1) = Y_pyrolysis(fireroom, i)
                        If Y_burnout(fireroom, i + 1) < Y_burnout(fireroom, i) Then Y_burnout(fireroom, i + 1) = Y_burnout(fireroom, i)
                        If X_pyrolysis(fireroom, i + 1) < X_pyrolysis(fireroom, i) Then X_pyrolysis(fireroom, i + 1) = X_pyrolysis(fireroom, i)

                        Xstart(fireroom, 1) = Y_pyrolysis(fireroom, i + 1)
                        Xstart(fireroom, 3) = Y_burnout(fireroom, i + 1)
                        Xstart(fireroom, 2) = X_pyrolysis(fireroom, i + 1)

                        'likewise for the floor
                        Yf_pyrolysis(fireroom, i + 1) = Xstart(fireroom, 4) 'Y pyrolsis front
                        Yf_burnout(fireroom, i + 1) = Xstart(fireroom, 5) 'Y burnout front
                        Xf_pyrolysis(fireroom, i + 1) = Xstart(fireroom, 6) 'X pyrolysis front

                        If frmQuintiere.optWindAided.Checked = True Then
                            If Yf_pyrolysis(fireroom, i + 1) < BurnerFlameLength Then
                                Yf_pyrolysis(fireroom, i + 1) = BurnerFlameLength
                            End If
                        End If

                        FloorVelocity(fireroom, 1, i + 1) = velocity(fireroom, 4) 'upward
                        FloorVelocity(fireroom, 2, i + 1) = velocity(fireroom, 6) 'lateral
                        FloorVelocity(fireroom, 3, i + 1) = velocity(fireroom, 5) 'burnout

                        If Yf_pyrolysis(fireroom, i + 1) < Yf_pyrolysis(fireroom, i) Then Yf_pyrolysis(fireroom, i + 1) = Yf_pyrolysis(fireroom, i)
                        If Yf_burnout(fireroom, i + 1) < Yf_burnout(fireroom, i) Then Yf_burnout(fireroom, i + 1) = Yf_burnout(fireroom, i)
                        If Xf_pyrolysis(fireroom, i + 1) < Xf_pyrolysis(fireroom, i) Then Xf_pyrolysis(fireroom, i + 1) = Xf_pyrolysis(fireroom, i)

                        If FireLocation(burner_id) = 0 And frmQuintiere.optWindAided.Checked = True Then
                            'centre of room
                            If RoomLength(fireroom) > RoomWidth(fireroom) Then
                                If Yf_pyrolysis(fireroom, i + 1) >= RoomLength(fireroom) / 2 Then Yf_pyrolysis(fireroom, i + 1) = RoomLength(fireroom) / 2
                            Else
                                If Yf_pyrolysis(fireroom, i + 1) >= RoomWidth(fireroom) / 2 Then Yf_pyrolysis(fireroom, i + 1) = RoomWidth(fireroom) / 2
                            End If
                        ElseIf FireLocation(burner_id) = 1 Then
                            'against a wall
                            If RoomLength(fireroom) > RoomWidth(fireroom) Then
                                If Yf_pyrolysis(fireroom, i + 1) >= RoomLength(fireroom) Then Yf_pyrolysis(fireroom, i + 1) = RoomLength(fireroom)
                            Else
                                If Yf_pyrolysis(fireroom, i + 1) >= RoomWidth(fireroom) Then Yf_pyrolysis(fireroom, i + 1) = RoomWidth(fireroom)
                            End If
                        Else
                        End If

                        If FireLocation(burner_id) = 0 And frmQuintiere.optWindAided.Checked = True Then
                            If Yf_burnout(fireroom, i + 1) >= RoomLength(fireroom) / 2 Then Yf_burnout(fireroom, i + 1) = RoomLength(fireroom) / 2
                        Else
                            If Yf_burnout(fireroom, i + 1) >= RoomLength(fireroom) Then Yf_burnout(fireroom, i + 1) = RoomLength(fireroom)
                        End If

                        If FireLocation(burner_id) = 0 Then 'centre of the room
                            If RoomLength(fireroom) > RoomWidth(fireroom) Then
                                If Xf_pyrolysis(fireroom, i + 1) > 0.5 * RoomLength(fireroom) Then
                                    Xf_pyrolysis(fireroom, i + 1) = 0.5 * RoomLength(fireroom)
                                End If
                            Else
                                If Xf_pyrolysis(fireroom, i + 1) > 0.5 * RoomWidth(fireroom) Then
                                    Xf_pyrolysis(fireroom, i + 1) = 0.5 * RoomWidth(fireroom)
                                End If
                            End If
                        ElseIf FireLocation(burner_id) = 1 Then  'against wall
                            If RoomLength(fireroom) > RoomWidth(fireroom) Then
                                If Xf_pyrolysis(fireroom, i + 1) > RoomLength(fireroom) Then
                                    Xf_pyrolysis(fireroom, i + 1) = RoomLength(fireroom)
                                End If
                            Else
                                If Xf_pyrolysis(fireroom, i + 1) > RoomWidth(fireroom) Then
                                    Xf_pyrolysis(fireroom, i + 1) = RoomWidth(fireroom)
                                End If
                            End If
                        Else 'corner
                            If Xf_pyrolysis(fireroom, i + 1) > Sqrt(RoomWidth(fireroom) ^ 2 + RoomLength(fireroom) ^ 2) Then
                                Xf_pyrolysis(fireroom, i + 1) = Sqrt(RoomWidth(fireroom) ^ 2 + RoomLength(fireroom) ^ 2)
                            End If
                        End If

                        Xstart(fireroom, 4) = Yf_pyrolysis(fireroom, i + 1)
                        Xstart(fireroom, 5) = Yf_burnout(fireroom, i + 1)
                        Xstart(fireroom, 6) = Xf_pyrolysis(fireroom, i + 1)
                    End If

                    'record the time when the Y_pyrolysis front reaches the ceiling
                    If Y_pyrolysis(fireroom, i + 1) >= RoomHeight(fireroom) And YPFlag = 0 Then
                        timeH = tim(i, 1)
                        YPFlag = 1
                    End If
                Else
                    'if the wall has not yet ignited
                    Y_pyrolysis(fireroom, i + 1) = Y_pyrolysis(fireroom, i)
                    X_pyrolysis(fireroom, i + 1) = X_pyrolysis(fireroom, i)
                    Y_burnout(fireroom, i + 1) = Y_burnout(fireroom, i)
                    FlameVelocity(fireroom, 1, i + 1) = 0
                    FlameVelocity(fireroom, 2, i + 1) = 0
                    FlameVelocity(fireroom, 3, i + 1) = 0
                End If


                'ignition of floor coverings in an adjacent room
                'If Flashover = True Or FloorIgniteFlag(fireroom) = 1 Then
                '    Call Flooring_nextroom(Fstart(), fvelocity())
                'End If

                'solve the flame spread ODE's for the other rooms
                For room = 1 To NumberRooms
                    If room <> fireroom Then
                        If tim(i, 1) > CeilingIgniteTime(room) Or tim(i, 1) > WallIgniteTime(room) Or tim(i, 1) > FloorIgniteTime(room) Then
                            X(1) = Xstart(room, 1)
                            V(1) = velocity(room, 1)
                            X(2) = Xstart(room, 2)
                            V(2) = velocity(room, 2)
                            X(3) = Xstart(room, 3)
                            V(3) = velocity(room, 3)

                            X(4) = Xstart(room, 4) 'upward floor
                            X(5) = Xstart(room, 5) 'burnout floor
                            If tim(i, 1) >= FloorIgniteTime(room) Then X(6) = Xstart(room, 6) 'lateral floor
                            V(4) = velocity(room, 4)
                            If tim(i, 1) >= FloorIgniteTime(room) Then V(6) = velocity(room, 6)
                            V(5) = velocity(room, 5)

                            '===============================================================================
                            Call ODE_Spread_Solver(flux_wall_AR, room, fheight, V, X)
                            '===============================================================================

                            velocity(room, 1) = V(1) 'upward velocity
                            Xstart(room, 1) = X(1) 'upward pyrolysis front
                            velocity(room, 2) = V(2) 'lateral velocity
                            Xstart(room, 2) = X(2) 'lateral pyrolysis front
                            Xstart(room, 3) = X(3) 'burnout front
                            velocity(room, 3) = V(3) 'burnout velocity

                            Xstart(room, 4) = X(4) 'wind-aided floor spread
                            Xstart(room, 5) = X(5) 'burnout
                            If tim(i, 1) >= FloorIgniteTime(room) Then Xstart(room, 6) = X(6) 'lateral
                            velocity(room, 4) = V(4)
                            If tim(i, 1) >= FloorIgniteTime(room) Then velocity(room, 6) = V(6)
                            velocity(room, 5) = V(5)

                            'returned values
                            If i <> NumberTimeSteps + 1 Then
                                'limit the maximum length of the Y fronts
                                If RoomWidth(room) = WallLength2(fireroom, room, 1) Then
                                    MaxLength = RoomLength(room)
                                Else
                                    MaxLength = RoomWidth(room)
                                End If
                                If Xstart(room, 1) < MaxLength Then
                                    Y_pyrolysis(room, i + 1) = Xstart(room, 1) 'Y pyrolsis front
                                Else
                                    Y_pyrolysis(room, i + 1) = MaxLength + RoomHeight(room) 'Y pyrolsis front
                                End If
                                If Xstart(room, 2) < RoomLength(room) + RoomWidth(room) Then
                                    X_pyrolysis(room, i + 1) = Xstart(room, 2) 'x pyrolsis front
                                Else
                                    X_pyrolysis(room, i + 1) = RoomLength(room) + RoomWidth(room) 'Y pyrolsis front
                                End If
                                FlameVelocity(room, 1, i + 1) = velocity(room, 1) 'upward
                                FlameVelocity(room, 2, i + 1) = velocity(room, 2) 'lateral
                                FlameVelocity(room, 3, i + 1) = velocity(room, 3) 'burnout
                                Y_burnout(room, i + 1) = Xstart(room, 3)

                                If Y_pyrolysis(room, i + 1) < Y_pyrolysis(room, i) Then Y_pyrolysis(room, i + 1) = Y_pyrolysis(room, i)
                                If Y_burnout(room, i + 1) < Y_burnout(room, i) Then Y_burnout(room, i + 1) = Y_burnout(room, i)
                                If X_pyrolysis(room, i + 1) < X_pyrolysis(room, i) Then X_pyrolysis(room, i + 1) = X_pyrolysis(room, i)

                                Xstart(room, 1) = Y_pyrolysis(room, i + 1)
                                Xstart(room, 3) = Y_burnout(room, i + 1)
                                Xstart(room, 2) = X_pyrolysis(room, i + 1)


                                'likewise for the floor
                                If Xstart(room, 4) < MaxLength Then
                                    Yf_pyrolysis(room, i + 1) = Xstart(room, 4) 'Y pyrolsis front
                                Else
                                    Yf_pyrolysis(room, i + 1) = MaxLength 'Y pyrolsis front
                                End If
                                'If Yf_pyrolysis(room, i + 1) < BurnerFlameLength Then
                                '    Yf_pyrolysis(room, i + 1) = BurnerFlameLength
                                'End If
                                'Xstart(room, 4) = Yf_pyrolysis(room, i + 1)
                                Yf_burnout(room, i + 1) = Xstart(room, 5) 'Y burnout front
                                FloorVelocity(room, 1, i + 1) = velocity(room, 4) 'upward
                                FloorVelocity(room, 2, i + 1) = velocity(room, 6) 'lateral
                                FloorVelocity(room, 3, i + 1) = velocity(room, 5) 'burnout
                                Xf_pyrolysis(room, i + 1) = Xstart(room, 6) 'X pyrolysis front
                                If Xf_pyrolysis(room, i + 1) > 0.5 * RoomWidth(room) Then
                                    Xf_pyrolysis(room, i + 1) = 0.5 * RoomWidth(room) 'assume spread out from centre of room
                                End If
                                'Xstart(room, 6) = Xf_pyrolysis(room, i + 1)

                                If Yf_pyrolysis(room, i + 1) < Yf_pyrolysis(room, i) Then Yf_pyrolysis(room, i + 1) = Yf_pyrolysis(room, i)
                                If Yf_burnout(room, i + 1) < Yf_burnout(room, i) Then Yf_burnout(room, i + 1) = Yf_burnout(room, i)
                                If Xf_pyrolysis(room, i + 1) < Xf_pyrolysis(room, i) Then Xf_pyrolysis(room, i + 1) = Xf_pyrolysis(room, i)

                                Xstart(room, 4) = Yf_pyrolysis(room, i + 1)
                                Xstart(room, 5) = Yf_burnout(room, i + 1)
                                Xstart(room, 6) = Xf_pyrolysis(room, i + 1)

                            End If

                        Else
                            'if the ceiling and wall not yet ignited
                            Y_pyrolysis(room, i + 1) = Y_pyrolysis(room, i)
                            Yf_pyrolysis(room, i + 1) = Yf_pyrolysis(room, i)
                            FlameVelocity(room, 1, i + 1) = 0 'upward flame spread
                            FloorVelocity(room, 1, i + 1) = 0 'upward flame spread
                            X_pyrolysis(room, i + 1) = X_pyrolysis(room, i)
                            Xf_pyrolysis(room, i + 1) = Xf_pyrolysis(room, i)
                            FlameVelocity(room, 2, i + 1) = 0 'lateral flame spread
                            FloorVelocity(room, 2, i + 1) = 0 'lateral flame spread
                            FlameVelocity(room, 3, i + 1) = 0 'lateral flame spread
                            FloorVelocity(room, 3, i + 1) = 0 'lateral flame spread
                            Y_burnout(room, i + 1) = Y_burnout(room, i)
                            Yf_burnout(room, i + 1) = Yf_burnout(room, i)
                        End If
                    End If
                Next room
            End If

            For j = 1 To NumberRooms
                'total soot generation rate g-soot/sec
                If FuelResponseEffects = True Then
                    SootMass = SootMass_Rate_withfuelresponse(tim(i, 1), i)
                Else
                    SootMass = SootMass_Rate(tim(i, 1), i)
                End If


                'effective molecular weight of the layers
                mw_upper = MolecularWeightCO * COMassFraction(j, stepcount, 1) + MolecularWeightCO2 * CO2MassFraction(j, stepcount, 1) + MolecularWeightH2O * H2OMassFraction(j, stepcount, 1) + MolecularWeightO2 * O2MassFraction(j, stepcount, 1) + MolecularWeightN2 * (1 - O2MassFraction(j, stepcount, 1) - COMassFraction(j, stepcount, 1) - CO2MassFraction(j, stepcount, 1) - H2OMassFraction(j, stepcount, 1))
                mw_lower = MolecularWeightCO * COMassFraction(j, stepcount, 2) + MolecularWeightCO2 * CO2MassFraction(j, stepcount, 2) + MolecularWeightH2O * H2OMassFraction(j, stepcount, 2) + MolecularWeightO2 * O2MassFraction(j, stepcount, 2) + MolecularWeightN2 * (1 - O2MassFraction(j, stepcount, 2) - COMassFraction(j, stepcount, 2) - CO2MassFraction(j, stepcount, 2) - H2OMassFraction(j, stepcount, 2))

                'visibility in the upper layer, also returns the optical density value
                Visi_upper = Get_Visibility(mw_upper, RoomPressure(j, i), uppertemp(j, i), SootMassFraction(j, i, 1), FuelMassLossRate(i, fireroom), SootMass, OpticalDensity)

                'save optical density in the upper layer
                OD_upper(j, i) = OpticalDensity

                If j = fireroom Then SPR(j, i) = SootMassGen(i) * 8790

                Visi_lower = Get_Visibility(mw_lower, RoomPressure(j, i), lowertemp(j, i), SootMassFraction(j, i, 2), FuelMassLossRate(i, fireroom), SootMass, OpticalDensity)
                OD_lower(j, i) = OpticalDensity

                'calulate the visibility at the monitoring height
                If layerheight(j, i) <= MonitorHeight Then
                    Visibility(j, i) = Visi_upper
                Else
                    Visibility(j, i) = Visi_lower
                End If

                'if we have a smoke detector in room of origin,
                'get the smoke concentration at the appropriate radial distance
                If HaveSD(j) = True Then

                    If j = fireroom Then
                        If ObjectMLUA(2, 1) > 0 Then 'object 1 only
                            'hrrua
                            diameter = 2 * Sqrt((HeatRelease(fireroom, i, 2) / ObjectMLUA(2, 1)) / PI)
                        Else

                            'get mass loss rate for object
                            Dim mlua As Single = ObjectMLUA(0, 1) * (Target(fireroom, stepcount - 1) - Target(fireroom, 1)) + ObjectMLUA(1, 1)

                            'find fire diameter for each object
                            diameter = Fire_Diameter(FuelMassLossRate(i, fireroom), mlua)
                        End If

                        If sootmode = True Then 'manual entry of pre post flashover soot yields
                            If Flashover = True Then
                                SmokeConcentration(j, i) = SmokeJET(SDRadialDist(j), mw_upper, mw_lower, RoomPressure(j, i), postSoot, fireroom, layerheight(j, i), HeatRelease(fireroom, i, 2), SootMassFraction(j, i, 1), diameter, uppertemp(j, i), lowertemp(j, i))
                            Else
                                SmokeConcentration(j, i) = SmokeJET(SDRadialDist(j), mw_upper, mw_lower, RoomPressure(j, i), preSoot, fireroom, layerheight(j, i), HeatRelease(fireroom, i, 2), SootMassFraction(j, i, 1), diameter, uppertemp(j, i), lowertemp(j, i))
                            End If
                        Else
                            'based on soot yield from object number 1 only
                            SmokeConcentration(j, i) = SmokeJET(SDRadialDist(j), mw_upper, mw_lower, RoomPressure(j, i), SootYield(1), fireroom, layerheight(j, i), HeatRelease(fireroom, i, 2), SootMassFraction(j, i, 1), diameter, uppertemp(j, i), lowertemp(j, i))
                        End If

                        OD_outside(j, i) = SmokeConcentration(j, i) * ParticleExtinction / 2.3 '1/m

                    Else
                        'no ceiling jet
                        'is the detector in the upper layer or lower layer?
                        If SDdepth(j) < RoomHeight(j) - layerheight(j, i) Then
                            OD_outside(j, i) = OD_upper(j, i)
                            SmokeConcentration(j, i) = OD_upper(j, i) * 2.3 / ParticleExtinction
                        Else
                            OD_outside(j, i) = OD_lower(j, i)
                            SmokeConcentration(j, i) = OD_lower(j, i) * 2.3 / ParticleExtinction
                        End If
                    End If
                End If
            Next j

            For Each oSmokeDet In oSmokeDets
                If oSmokeDet.responsetime = 0 Then 'only do this if SD not yet responded

                    Dim sdid As Integer = oSmokeDet.sdid
                    j = oSmokeDet.room
                    Dim r As Single = oSmokeDet.sdr

                    If j = fireroom Then
                        If ObjectMLUA(2, 1) > 0 Then 'object 1 only
                            'hrrua
                            diameter = 2 * Sqrt((HeatRelease(fireroom, i, 2) / ObjectMLUA(2, 1)) / PI)
                        Else
                            'get mass loss rate for object
                            Dim mlua As Single = ObjectMLUA(0, 1) * (Target(fireroom, stepcount - 1) - Target(fireroom, 1)) + ObjectMLUA(1, 1)

                            'find fire diameter for each object
                            diameter = Fire_Diameter(FuelMassLossRate(i, fireroom), mlua)
                        End If
                        If sootmode = True Then 'manual entry of pre post flashover soot yields
                            'If GlobalER(i) > 1 Then
                            If Flashover = True Then
                                'need to change this to store data for each det, instead of for each room
                                SmokeConcentrationSD(sdid, i) = SmokeJET(r, mw_upper, mw_lower, RoomPressure(j, i), postSoot, fireroom, layerheight(j, i), HeatRelease(fireroom, i, 2), SootMassFraction(j, i, 1), diameter, uppertemp(j, i), lowertemp(j, i))
                            Else
                                SmokeConcentrationSD(sdid, i) = SmokeJET(r, mw_upper, mw_lower, RoomPressure(j, i), preSoot, fireroom, layerheight(j, i), HeatRelease(fireroom, i, 2), SootMassFraction(j, i, 1), diameter, uppertemp(j, i), lowertemp(j, i))
                            End If
                        Else
                            'based on soot yield from object number 1 only
                            'do we need pre/post flashover option here for sootyield(1)
                            SmokeConcentrationSD(sdid, i) = SmokeJET(r, mw_upper, mw_lower, RoomPressure(j, i), SootYield(1), fireroom, layerheight(j, i), HeatRelease(fireroom, i, 2), SootMassFraction(j, i, 1), diameter, uppertemp(j, i), lowertemp(j, i))
                        End If

                        'need to change this to store data for each det, instead of for each room
                        OD_outsideSD(sdid, i) = SmokeConcentrationSD(sdid, i) * ParticleExtinction / 2.3 '1/m

                        'If OD_outsideSD(sdid, i) > oSmokeDet.od Then
                        'check for smoke detector operation including transit time
                        If SDFlagSD(sdid) = 0 And i < NumberTimeSteps + 1 Then
                            If sd_mode = True Then 'smoke detection system active
                                Call SD_Flags(oSmokeDet, OD_outsideSD(sdid, i), j, i, tim(i, 1), RoomHeight(j) - layerheight(j, i))
                                SDtransit(sdid, i) = oSmokeDet.transit
                            End If
                            'If oSmokeDet.responsetime > 0 Then Stop
                        End If
                        'End If

                    Else
                        'not room  of fire origin
                        'is the detector in the upper or lower layer?

                        OD_outsideSD(sdid, i) = OD_upper(j, i) '1/m
                        'OD_upper(1, 1)
                        'check for smoke detector operation
                        If SDFlagSD(sdid) = 0 And i < NumberTimeSteps + 1 Then
                            If sd_mode = True Then 'smoke detection system active
                                Call SD_Flags(oSmokeDet, OD_outsideSD(sdid, i), j, i, tim(i, 1), RoomHeight(j) - layerheight(j, i))
                            End If
                        End If
                    End If
                End If
            Next

            If Flashover = False Then

                If FOFluxCriteria = True Then
                    'flashover criteria = radiant flux on floor
                    If Target(fireroom, i) >= flashover_crit_flux Then
                        Flashover = True
                        flashover_time = tim(i, 1)
                        HRRatFO = HeatRelease(fireroom, i - 1, 2)
                        Dim Message As String = CStr(tim(i, 1)) & " sec. Flashover in Room " & CStr(fireroom) & "."
                        If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                    End If
                Else
                    'flashover criteria = upper layer temperature
                    If uppertemp(fireroom, i) >= flashover_crit_temp Then
                        Flashover = True
                        flashover_time = tim(i, 1)
                        HRRatFO = HeatRelease(fireroom, i - 1, 2)
                        Dim Message As String = CStr(tim(i, 1)) & " sec. Flashover in Room " & CStr(fireroom) & "."
                        If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                    End If
                End If

                If Flashover = True And usepowerlawdesignfire = False Then

                    alphaTfitted(0, itcounter - 1) = HeatRelease(fireroom, i, 2) / (tim(i, 1) ^ 2)
                    alphaTfitted(1, itcounter - 1) = HeatRelease(fireroom, i, 2)
                    alphaTfitted(2, itcounter - 1) = tim(i, 1)
                    Dim Message As String = "Fitted t2 alpha coefficient (kW/s2) = " & CStr(Format(alphaTfitted(0, itcounter - 1), "0.000"))
                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                End If

            End If

            'keep track of total fuel pyrolysed
            If i > 1 Then

                If IntegralModel = True And useCLTmodel = True Then
                    'only want the contents mass consumed here, so deduct the contribution from wood surfaces
                    TotalFuel(i) = TotalFuel(i - 1) + (FuelMassLossRate(i, fireroom) - WoodBurningRate(i)) * Timestep
                ElseIf KineticModel = True And useCLTmodel = True Then
                    TotalFuel(i) = TotalFuel(i - 1) + (FuelMassLossRate(i, fireroom) - WoodBurningRate(i)) * Timestep
                Else
                    TotalFuel(i) = TotalFuel(i - 1) + FuelMassLossRate(i, fireroom) * Timestep 'use this for simple dynamic CLT model
                End If
            Else
                TotalFuel(i) = 0
            End If

            'call procedure to evaluate endpoint flags
            EndPoint_Flags(tim(i, 1), FEDRadSum(fireroom, i), uppertemp(fireroom, i), Visibility(fireroom, i), LinkTemp(fireroom, i), layerheight(fireroom, i), lowertemp(fireroom, i), HeatRelease(fireroom, i, 2))

            For j = 1 To NumberRooms
                If fanon(j) = True Then

                    If FanAutoStart(j) = True And AutoFlag(j) = False Then
                        'fan not yet started
                        ExtractStartTime(j) = SimTime
                        If (j = fireroom And SprinklerFlag = 1) Then
                            'sprinkler in fireroom starts fan in fireroom
                            ExtractStartTime(j) = SprinklerTime
                            AutoFlag(j) = 1
                            Dim Message As String = CStr(SprinklerTime) & " sec. Fan started in room " & j.ToString
                            If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                        End If
                        If (j = fireroom And HDFlag = 1) Then
                            'hd in fireroom starts fan in fireroom
                            ExtractStartTime(j) = HDTime
                            AutoFlag(j) = 1
                            Dim Message As String = CStr(HDTime) & " sec. Fan started in room " & j.ToString
                            If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                        End If
                        If SDFlag(j) = 1 Then
                            'smoke detector in room starts fan
                            ExtractStartTime(j) = SDTime(j)
                            AutoFlag(j) = True
                            Dim Message As String = Format(SDTime(j), "0") & " sec. Fan started in room " & j.ToString
                            If ProjectDirectory = RiskDataDirectory Then frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                        End If
                    End If
                End If
            Next j

            If Flashover = True And terminate_fo = True Then
                flagstop = 1
            End If
            If terminate_fuelgone = True Then
                If TotalFuel(stepcount - 1) > InitialFuelMass Then
                    flagstop = 1
                End If
            End If

            'Increases the progress indicator to reflect the % progress made
            MDIFrmMain.ToolStripProgressBar1.Value = i / (NumberTimeSteps + 1) * 100

            If i = 1 Then
                'show run-time graphs
                If MDIFrmMain.mnuGraphsOn_HRR.Checked = True Then
                    MDIFrmMain.ChartRuntime1.Visible = True
                    MDIFrmMain.ChartRuntime2.Visible = True
                End If
            End If

            'update run-time graphs
            If MDIFrmMain.mnuGraphsOn_HRR.Checked = True Then
                '	'call procedure to plot data
                Dim Title As String, DataShift As Single
                Dim DataMultiplier As Single

                'define variables
                Title = "Heat Release Rate (kW)"
                DataShift = 0
                DataMultiplier = 1
                'call procedure to plot data
                Graph_Data_Runtime(2, Title, HeatRelease, DataShift, DataMultiplier)
                Title = "Upper Layer Temperature (C)"
                DataShift = -273
                DataMultiplier = 1
                Graph_Data_Runtime2(MDIFrmMain.ChartRuntime2, Title, uppertemp, DataShift, DataMultiplier)
            End If

            If TalkToEVACNZ = True Then
                brisksimtime = tim(i, 1)

                'save output for this timestep 
                Call frmInputs.save_evacnz_xml2(ventstatus)

                'read - returns ventstatus array
                Call frmInputs.read_evacnz_xml2(ventstatus)
            End If

        Next i
        'end of main loop ************************************

        If TalkToEVACNZ = True Then
            brisksimtime = SimTime
            Call frmInputs.save_evacnz_xml2(ventstatus) 'last step update
        End If
       

        MDIFrmMain.ToolStripStatusLabel4.Text = FormatNumber(SimTime, 1) & " sec - finished."

        SprinklerDB.SaveSprinklers2(oSprinklers, osprdistributions)
        SmokeDetDB.SaveSmokeDets(oSmokeDets, osddistributions)

        MDIFrmMain.ToolStripStatusLabel3.Text = " "

        MDIFrmMain.ToolStripStatusLabel2.Text = " "
        TwoZones(fireroom) = True

    End Sub


    Public Sub window_break(ByVal i As Integer, ByVal room As Short, ByRef breakflag(,,) As Boolean, ByRef Ystart(,) As Double)
        '===================================================================================
        ' code for calling the glass breakage routines
        ' c wade 11/4/2002
        '===================================================================================
        Dim j, vent As Integer

        For j = 1 To NumberRooms + 1
            If j > room Then
                For vent = 1 To NumberVents(room, j)
                    If breakflag(room, j, vent) = False And AutoBreakGlass(room, j, vent) = True Then
                        VentOpenTime(room, j, vent) = SimTime
                        If j < NumberRooms + 1 Then
                            'if the vent connects to another internal space
                            If soffitheight(room, j, vent) > layerheight(room, i) Then
                                'If window in hot layer, otherwise it's in the lower layer
                                Call Break_Glass(room, j, vent, tim(i, 1), breakflag, VentHeight(room, j, vent), VentWidth(room, j, vent), Ystart(room, 2), 1 - TransmissionFactor(2, 2, room)) ', -QUpperWall(room, i), 1 - TransmissionFactor(3, 3, room), 1 - TransmissionFactor(3, 3, j))
                            Else
                                Call Break_Glass(room, j, vent, tim(i, 1), breakflag, VentHeight(room, j, vent), VentWidth(room, j, vent), Ystart(room, 3), 1 - TransmissionFactor(3, 3, room))
                            End If
                        Else
                            'if the vent connects to the outside
                            If soffitheight(room, j, vent) > layerheight(room, i) Then
                                'If window in hot layer
                                Call Break_Glass(room, j, vent, tim(i, 1), breakflag, VentHeight(room, j, vent), VentWidth(room, j, vent), Ystart(room, 2), 1 - TransmissionFactor(2, 2, room))
                            Else : Call Break_Glass(room, j, vent, tim(i, 1), breakflag, VentHeight(room, j, vent), VentWidth(room, j, vent), Ystart(room, 3), 1 - TransmissionFactor(3, 3, room))
                            End If
                        End If
                        If breakflag(room, j, vent) = True Then
                            'glass has broken so open the vent
                            VentOpenTime(room, j, vent) = tim(i, 1)
                        End If
                    End If
                Next vent
            End If
        Next j

    End Sub
    Public Sub cvent_open(ByVal i As Long)

        'fire resistance criteria for opening a ceiling vent

        Dim j As Long, room As Integer, vent As Integer
        Try

            For room = 1 To NumberRooms + 1 'upper room
                For j = 1 To NumberRooms + 1 'lower room
                    'If j > room Then
                    For vent = 1 To NumberCVents(room, j)

                        'Fire Barrier Model
                        If breakflag2(room, j, vent) = False And trigger_device2(6, room, j, vent) = True Then

                            CVentOpenTime(room, j, vent) = SimTime 'vent is closed
                            CVentCloseTime(room, j, vent) = SimTime

                            If FRcriteria2(room, j, vent) = 0 Then
                                'upper layer gas temperature
                                If uppertemp(j, i) > FRgastemp2(room, j, vent) + 273 Then
                                    'integrity failure
                                    breakflag2(room, j, vent) = True
                                    CVentBreakTime(room, j, vent) = tim(i, 1)
                                    CVentOpenTime(room, j, vent) = tim(i, 1)
                                    Dim Message As String = CStr(tim(i, 1)) & " sec. Ceiling Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened - upper layer temperature limit exceeded."
                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                                    VentilationLimitFlag = False 'reset
                                End If
                            ElseIf FRcriteria2(room, j, vent) = 2 Then
                                'NHL Harmathy Furnace
                                'calculate time of equiv. exposure in NRCC furnace from Harmathy
                                Dim H As Double = 0

                                H = NHL(1, j, i) 'NHL in compartment area-weighted
                                NHL(2, j, i) = 60 * (0.11 + 0.000016 * H + 0.00000000013 * H ^ 2) 'minutes
                                If NHL(2, j, i) > FRintegrity2(room, j, vent) Then
                                    'integrity failure
                                    breakflag2(room, j, vent) = True
                                    CVentBreakTime(room, j, vent) = tim(i, 1)
                                    CVentOpenTime(room, j, vent) = tim(i, 1)
                                    Dim Message As String = CStr(tim(i, 1)) & " sec. Ceiling Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened - integrity failure, NHL = " & Format(H, 0)
                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                                    VentilationLimitFlag = False 'reset
                                End If
                            ElseIf FRcriteria2(room, j, vent) = 1 Then
                                'Energy Nyman's Method
                                'check to see if energy dose received has exceeded integrity value
                                If energy(j, stepcount) > FRfaildata2(room, j, vent) Then
                                    'integrity failure
                                    breakflag2(room, j, vent) = True
                                    CVentBreakTime(room, j, vent) = tim(i, 1)
                                    CVentOpenTime(room, j, vent) = tim(i, 1)
                                    Dim Message As String = CStr(tim(i, 1)) & " sec. Ceiling Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened - integrity failure."
                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                                    VentilationLimitFlag = False 'reset
                                End If
                            ElseIf FRcriteria2(room, j, vent) = 3 Then
                                'NHL ceiling
                                If NHL(3, j, i) > FRfaildata2(room, j, vent) Then
                                    'integrity failure
                                    breakflag2(room, j, vent) = True
                                    CVentBreakTime(room, j, vent) = tim(i, 1)
                                    CVentOpenTime(room, j, vent) = tim(i, 1)
                                    Dim Message As String = CStr(tim(i, 1)) & " sec. Ceiling Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened - integrity failure, NHL = " & Format(NHL(3, j, i), 0)
                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                                    VentilationLimitFlag = False 'reset
                                End If
                            ElseIf FRcriteria2(room, j, vent) = 4 Then
                                'NHL upper wall
                                If NHL(4, j, i) > FRfaildata2(room, j, vent) Then
                                    'integrity failure
                                    breakflag2(room, j, vent) = True
                                    CVentBreakTime(room, j, vent) = tim(i, 1)
                                    CVentOpenTime(room, j, vent) = tim(i, 1)
                                    Dim Message As String = CStr(tim(i, 1)) & " sec. Ceiling Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened - integrity failure, NHL = " & Format(NHL(4, j, i), 0)
                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                                    VentilationLimitFlag = False 'reset
                                End If
                            ElseIf FRcriteria2(room, j, vent) = 5 Then
                                'NHL lower wall
                                If NHL(5, j, i) > FRfaildata2(room, j, vent) Then
                                    'integrity failure
                                    breakflag2(room, j, vent) = True
                                    CVentBreakTime(room, j, vent) = tim(i, 1)
                                    CVentOpenTime(room, j, vent) = tim(i, 1)
                                    Dim Message As String = CStr(tim(i, 1)) & " sec. Ceiling Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened - integrity failure, NHL = " & Format(NHL(5, j, i), 0)
                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                                    VentilationLimitFlag = False 'reset
                                End If
                            ElseIf FRcriteria2(room, j, vent) = 6 Then
                                'NHL floor
                                If NHL(6, j, i) > FRfaildata2(room, j, vent) Then
                                    'integrity failure
                                    breakflag2(room, j, vent) = True
                                    CVentBreakTime(room, j, vent) = tim(i, 1)
                                    CVentOpenTime(room, j, vent) = tim(i, 1)
                                    Dim Message As String = CStr(tim(i, 1)) & " sec. Ceiling Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened - integrity failure, NHL = " & Format(NHL(6, j, i), 0)
                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                                    VentilationLimitFlag = False 'reset
                                End If
                            End If

                        End If

                        If breakflag2(room, j, vent) = False And CVentAuto(room, j, vent) = True Then

                            'if the auto vent open box is checked, then the initial state of the vent is closed, 
                            CVentOpenTime(room, j, vent) = SimTime + Timestep 'vent is closed
                            'If trigger_device(5, room, j, vent) = True And HOFlag(room, j, vent) = False Then
                            '    VentOpenTime(room, j, vent) = 0 'keep vent initially open if we have hold-open device
                            'End If

                            CVentCloseTime(room, j, vent) = SimTime + Timestep

                            'code to test vent trigger conditions
                            'if SD has operated
                            If trigger_device2(0, room, j, vent) = True And SDFlag(SDtriggerroom2(room, j, vent)) = 1 And breakflag2(room, j, vent) = False Then
                                If tim(i, 1) > SDTime(SDtriggerroom2(room, j, vent)) + trigger_ventopendelay2(room, j, vent) Then
                                    breakflag2(room, j, vent) = True
                                    VentilationLimitFlag = False 'reset
                                    PeakHRR = originalpeakhrr
                                    CVentOpenTime(room, j, vent) = tim(i, 1) - Timestep
                                    CVentBreakTime(room, j, vent) = tim(i, 1) - Timestep
                                    CVentCloseTime(room, j, vent) = tim(i, 1) - Timestep + trigger_ventopenduration2(room, j, vent)
                                    CVentBreakTimeClose(room, j, vent) = tim(i, 1) - Timestep + trigger_ventopenduration2(room, j, vent)

                                    Dim Message As String = CStr(CVentOpenTime(room, j, vent)) & " sec. Ceiling Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened " & trigger_ventopendelay2(room, j, vent).ToString & " sec following smoke detector operation."
                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                                    Message = CStr(CVentCloseTime(room, j, vent)) & " sec. Ceiling Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " closed."
                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                                End If
                            End If

                            'if HD/spr has operated
                            If trigger_device2(1, room, j, vent) = True And breakflag2(room, j, vent) = False Then
                                If SprinklerFlag = 1 Then

                                    If tim(i, 1) > SprinklerTime + trigger_ventopendelay2(room, j, vent) Then
                                        breakflag2(room, j, vent) = True
                                        VentilationLimitFlag = False 'reset
                                        PeakHRR = originalpeakhrr
                                        'if ventopentime is reset exactly at the sprinklertime+delay, convergence problems? so reset 1 timestep later
                                        CVentOpenTime(room, j, vent) = tim(i, 1) - Timestep
                                        CVentBreakTime(room, j, vent) = tim(i, 1) - Timestep
                                        CVentCloseTime(room, j, vent) = tim(i, 1) + trigger_ventopenduration2(room, j, vent) - Timestep
                                        CVentBreakTimeClose(room, j, vent) = tim(i, 1) + trigger_ventopenduration2(room, j, vent) - Timestep

                                        Dim Message As String = CStr(CVentOpenTime(room, j, vent)) & " sec. Ceiling Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened " & trigger_ventopendelay2(room, j, vent).ToString & " sec following sprinkler operation."
                                        frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                                        Message = CStr(CVentCloseTime(room, j, vent)) & " sec. Ceiling Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " closed."
                                        frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                                    End If
                                End If
                                If HDFlag = 1 Then
                                    If tim(i, 1) > HDTime + trigger_ventopendelay2(room, j, vent) Then
                                        breakflag2(room, j, vent) = True
                                        VentilationLimitFlag = False 'reset
                                        PeakHRR = originalpeakhrr
                                        CVentOpenTime(room, j, vent) = tim(i, 1) - Timestep
                                        CVentBreakTime(room, j, vent) = tim(i, 1) - Timestep
                                        CVentCloseTime(room, j, vent) = tim(i, 1) - Timestep + trigger_ventopenduration2(room, j, vent)
                                        CVentBreakTimeClose(room, j, vent) = tim(i, 1) - Timestep + trigger_ventopenduration2(room, j, vent)

                                        Dim Message As String = CStr(CVentOpenTime(room, j, vent)) & " sec. Ceiling Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened " & trigger_ventopendelay2(room, j, vent).ToString & " sec following heat detector operation."
                                        'MDIFrmMain.ToolStripStatusLabel2.Text = Message.ToString
                                        frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                                        Message = CStr(CVentCloseTime(room, j, vent)) & " sec. Ceiling Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " closed."
                                        frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                                    End If
                                End If
                            End If
                            'if fire size criteria met
                            If trigger_device2(2, room, j, vent) = True And (room = fireroom Or j = fireroom) Then
                                If HeatRelease(fireroom, i, 2) > HRR_threshold2(room, j, vent) And HRRFlag2(room, j, vent) = 0 Then
                                    HRR_threshold_time2(room, j, vent) = tim(i, 1)
                                    HRRFlag2(room, j, vent) = 1
                                End If
                                If tim(i, 1) > HRR_threshold_time2(room, j, vent) + HRR_ventopendelay2(room, j, vent) And HRRFlag2(room, j, vent) = 1 Then
                                    breakflag2(room, j, vent) = True
                                    VentilationLimitFlag = False 'reset
                                    PeakHRR = originalpeakhrr
                                    CVentOpenTime(room, j, vent) = tim(i, 1) - Timestep
                                    CVentBreakTime(room, j, vent) = tim(i, 1) - Timestep
                                    CVentCloseTime(room, j, vent) = tim(i, 1) - Timestep + HRR_ventopenduration2(room, j, vent)
                                    CVentBreakTimeClose(room, j, vent) = tim(i, 1) - Timestep + HRR_ventopenduration2(room, j, vent)

                                    Dim Message As String = CStr(CVentOpenTime(room, j, vent)) & " sec. Ceiling Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened when HRR reached " & HRR_threshold2(room, j, vent) & " kW"
                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                                    Message = CStr(CVentCloseTime(room, j, vent)) & " sec. Ceiling Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " closed."
                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                                End If
                            End If

                            'if flashover reached
                            If trigger_device2(3, room, j, vent) = True And breakflag2(room, j, vent) = False And (room = fireroom Or j = fireroom) Then
                                If Flashover = True Then
                                    breakflag2(room, j, vent) = True
                                    VentilationLimitFlag = False 'reset
                                    PeakHRR = originalpeakhrr
                                    CVentOpenTime(room, j, vent) = tim(i, 1)
                                    CVentBreakTime(room, j, vent) = tim(i, 1)

                                    Dim Message As String = CStr(tim(i, 1)) & " sec. Ceiling Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened at flashover."
                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                                End If
                            End If

                            'if ventilation limit reached
                            If trigger_device2(4, room, j, vent) = True And breakflag2(room, j, vent) = False And (room = fireroom Or j = fireroom) Then
                                If VentilationLimitFlag = True Then
                                    If tim(i, 1) > alphaTfitted(4, itcounter - 1) Then 'allow enough time for the vent to fully open

                                        breakflag2(room, j, vent) = True
                                        VentilationLimitFlag = False 'reset (when we open vent, fire might not be vent limited any more, so will need to retest
                                        PeakHRR = originalpeakhrr
                                        CVentOpenTime(room, j, vent) = tim(i, 1)
                                        CVentBreakTime(room, j, vent) = tim(i, 1)

                                        Dim Message As String = CStr(tim(i, 1)) & " sec. Ceiling Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened when ventilation limit reached."
                                        frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                                        For r = j To NumberRooms + 1
                                            For s = 1 To NumberCVents(room, r)
                                                If trigger_device2(4, room, r, s) = True And breakflag2(room, r, s) = False And AutoOpenCVent(room, r, s) = True Then
                                                    breakflag2(room, r, s) = True
                                                    CVentOpenTime(room, r, s) = tim(i, 1)
                                                    CVentBreakTime(room, r, s) = tim(i, 1)

                                                    Dim Message1 As String = CStr(tim(i, 1)) & " sec. Ceiling Vent " & CStr(room) & "-" & CStr(r) & "-" & CStr(s) & " opened when ventilation limit reached."
                                                    frmInputs.rtb_log.Text = Message1.ToString & Chr(13) & frmInputs.rtb_log.Text

                                                End If
                                            Next
                                        Next
                                    End If
                                End If
                            End If
                        End If

                    Next vent
                Next
            Next

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in DIFFEQNS.vb cvent_open")

        End Try
    End Sub
   

    Public Sub door_open(ByVal i As Long, ByVal room As Integer, ByVal breakflag(,,) As Boolean, ByVal HOflag(,,) As Boolean)
        'auto open doors 2009.1

        'if opendelaytime is 0, then the door remains open continuously
        'if the simulation is terminated early - the ventopentime and ventclosetimes are reset to original value before creating the view_output file

        Dim j As Long, vent As Long
        For j = 1 To NumberRooms + 1
            If j > room Then
                For vent = 1 To NumberVents(room, j)

                    'Talk to EvacuatioNZ
                    If TalkToEVACNZ = True Then
                        If enzbreakflag(room, j, vent) = False Then
                            'If breakflag(room, j, vent) = False Then
                            'close the vent
                            VentOpenTime(room, j, vent) = SimTime + Timestep 'vent is closed
                            VentCloseTime(room, j, vent) = SimTime + Timestep
                        Else
                            'open the vent
                            VentOpenTime(room, j, vent) = 0
                            VentCloseTime(room, j, vent) = SimTime + Timestep

                            'VentilationLimitFlag = False 'reset
                            'PeakHRR = originalpeakhrr
                            'VentBreakTime(room, j, vent) = tim(i, 1)
                            'VentBreakTimeClose(room, j, vent) = SimTime

                        End If
                    End If

                    'Fire Barrier Model
                    If breakflag(room, j, vent) = False And trigger_device(6, room, j, vent) = True Then

                        VentOpenTime(room, j, vent) = SimTime 'vent is closed
                        VentCloseTime(room, j, vent) = SimTime

                        If FRcriteria(room, j, vent) = 0 Then
                            'upper layer gas temperature
                            If uppertemp(room, i) > FRgastemp(room, j, vent) + 273 Then
                                'integrity failure
                                breakflag(room, j, vent) = True
                                VentBreakTime(room, j, vent) = tim(i, 1)
                                VentOpenTime(room, j, vent) = tim(i, 1)
                                Dim Message As String = CStr(tim(i, 1)) & " sec. Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened - upper layer temperature limit exceeded."
                                frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                                VentilationLimitFlag = False 'reset
                            End If
                        ElseIf FRcriteria(room, j, vent) = 2 Then
                            'NHL Harmathy Furnace
                            'calculate time of equiv. exposure in NRCC furnace from Harmathy
                            Dim H As Double = 0

                            H = NHL(1, room, stepcount) 'NHL in compartment area-weighted
                            NHL(2, room, stepcount) = 60 * (0.11 + 0.000016 * H + 0.00000000013 * H ^ 2) 'minutes
                            If NHL(2, room, stepcount) > FRintegrity(room, j, vent) Then
                                'integrity failure
                                breakflag(room, j, vent) = True
                                VentBreakTime(room, j, vent) = tim(i, 1)
                                VentOpenTime(room, j, vent) = tim(i, 1)
                                Dim Message As String = CStr(tim(i, 1)) & " sec. Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened - integrity failure."
                                frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                                VentilationLimitFlag = False 'reset
                            End If
                        ElseIf FRcriteria(room, j, vent) = 1 Then
                            'Energy Nyman's Method
                            'check to see if energy dose received has exceeded integrity value
                            If energy(room, stepcount) > FRfaildata(room, j, vent) Then
                                'integrity failure
                                breakflag(room, j, vent) = True
                                VentBreakTime(room, j, vent) = tim(i, 1)
                                VentOpenTime(room, j, vent) = tim(i, 1)
                                Dim Message As String = CStr(tim(i, 1)) & " sec. Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened - integrity failure."
                                frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                                VentilationLimitFlag = False 'reset
                            End If
                        ElseIf FRcriteria(room, j, vent) = 3 Then
                            'NHL ceiling
                            If NHL(3, room, i) > FRfaildata(room, j, vent) Then
                                'integrity failure
                                breakflag(room, j, vent) = True
                                VentBreakTime(room, j, vent) = tim(i, 1)
                                VentOpenTime(room, j, vent) = tim(i, 1)
                                Dim Message As String = CStr(tim(i, 1)) & " sec. Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened - integrity failure."
                                frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                                VentilationLimitFlag = False 'reset
                            End If
                        ElseIf FRcriteria(room, j, vent) = 4 Then
                            'NHL upper wall
                            If NHL(4, room, i) > FRfaildata(room, j, vent) Then
                                'integrity failure
                                breakflag(room, j, vent) = True
                                VentBreakTime(room, j, vent) = tim(i, 1)
                                VentOpenTime(room, j, vent) = tim(i, 1)
                                Dim Message As String = CStr(tim(i, 1)) & " sec. Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened - integrity failure."
                                frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                                VentilationLimitFlag = False 'reset
                            End If
                        ElseIf FRcriteria(room, j, vent) = 5 Then 'not currently selectable
                            'NHL lower wall
                            If NHL(5, room, i) > FRfaildata(room, j, vent) Then
                                'integrity failure
                                breakflag(room, j, vent) = True
                                VentBreakTime(room, j, vent) = tim(i, 1)
                                VentOpenTime(room, j, vent) = tim(i, 1)
                                Dim Message As String = CStr(tim(i, 1)) & " sec. Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened - integrity failure."
                                frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                                VentilationLimitFlag = False 'reset
                            End If
                        ElseIf FRcriteria(room, j, vent) = 6 Then 'not currently selectable
                            'NHL floor
                            If NHL(6, room, i) > FRfaildata(room, j, vent) Then
                                'integrity failure
                                breakflag(room, j, vent) = True
                                VentBreakTime(room, j, vent) = tim(i, 1)
                                VentOpenTime(room, j, vent) = tim(i, 1)
                                Dim Message As String = CStr(tim(i, 1)) & " sec. Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened - integrity failure."
                                frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                                VentilationLimitFlag = False 'reset
                            End If
                        End If

                    End If

                    If breakflag(room, j, vent) = False And AutoOpenVent(room, j, vent) = True Then

                        'if the auto vent open box is checked, then the initial state of the vent is closed, 
                        'except if a hold open device is fitted

                        VentOpenTime(room, j, vent) = SimTime + Timestep 'vent is closed
                        If trigger_device(5, room, j, vent) = True And HOflag(room, j, vent) = False Then
                            VentOpenTime(room, j, vent) = 0 'keep vent initially open if we have hold-open device
                        End If

                        VentCloseTime(room, j, vent) = SimTime + Timestep

                        'code to test vent trigger conditions
                        'if SD has operated
                        If trigger_device(0, room, j, vent) = True And SDFlag(SDtriggerroom(room, j, vent)) = 1 And breakflag(room, j, vent) = False Then
                            If tim(i, 1) > SDTime(SDtriggerroom(room, j, vent)) + trigger_ventopendelay(room, j, vent) Then
                                breakflag(room, j, vent) = True
                                VentilationLimitFlag = False 'reset
                                PeakHRR = originalpeakhrr
                                VentOpenTime(room, j, vent) = tim(i, 1) - Timestep
                                VentBreakTime(room, j, vent) = tim(i, 1) - Timestep
                                VentCloseTime(room, j, vent) = tim(i, 1) - Timestep + trigger_ventopenduration(room, j, vent)
                                VentBreakTimeClose(room, j, vent) = tim(i, 1) - Timestep + trigger_ventopenduration(room, j, vent)

                                Dim Message As String = CStr(VentOpenTime(room, j, vent)) & " sec. Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened " & trigger_ventopendelay(room, j, vent).ToString & " sec following smoke detector operation."
                                frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                                Message = CStr(VentCloseTime(room, j, vent)) & " sec. Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " closed."
                                frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                            End If
                        End If

                        'if SD operated activate hold open device
                        If trigger_device(5, room, j, vent) = True And SDFlag(SDtriggerroom(room, j, vent)) = 1 And HOflag(room, j, vent) = False And HOactive(room, j, vent) = True Then
                            If tim(i, 1) > SDTime(SDtriggerroom(room, j, vent)) + trigger_ventopendelay(room, j, vent) Then

                                HOflag(room, j, vent) = True
                                VentCloseTime(room, j, vent) = tim(i, 1) - Timestep
                                VentBreakTimeClose(room, j, vent) = tim(i, 1) - Timestep

                                Dim Message As String = CStr(VentCloseTime(room, j, vent)) & " sec. Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " closed by hold-open device " & trigger_ventopendelay(room, j, vent) & " sec after smoke detector activated"
                                frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                            End If
                        End If

                        'if HD/spr has operated
                        If trigger_device(1, room, j, vent) = True And breakflag(room, j, vent) = False Then
                            If SprinklerFlag = 1 Then

                                If tim(i, 1) > SprinklerTime + trigger_ventopendelay(room, j, vent) Then
                                    breakflag(room, j, vent) = True
                                    VentilationLimitFlag = False 'reset
                                    PeakHRR = originalpeakhrr
                                    'if ventopentime is reset exactly at the sprinklertime+delay, convergence problems? so reset 1 timestep later
                                    VentOpenTime(room, j, vent) = tim(i, 1) - Timestep
                                    VentBreakTime(room, j, vent) = tim(i, 1) - Timestep
                                    VentCloseTime(room, j, vent) = tim(i, 1) + trigger_ventopenduration(room, j, vent) - Timestep
                                    VentBreakTimeClose(room, j, vent) = tim(i, 1) + trigger_ventopenduration(room, j, vent) - Timestep

                                    Dim Message As String = CStr(VentOpenTime(room, j, vent)) & " sec. Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened " & trigger_ventopendelay(room, j, vent).ToString & " sec following sprinkler activation"
                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                                    Message = CStr(VentCloseTime(room, j, vent)) & " sec. Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " closed"
                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                                End If
                            End If
                            If HDFlag = 1 Then
                                If tim(i, 1) > HDTime + trigger_ventopendelay(room, j, vent) Then
                                    breakflag(room, j, vent) = True
                                    VentilationLimitFlag = False 'reset
                                    PeakHRR = originalpeakhrr
                                    VentOpenTime(room, j, vent) = tim(i, 1) - Timestep
                                    VentBreakTime(room, j, vent) = tim(i, 1) - Timestep
                                    VentCloseTime(room, j, vent) = tim(i, 1) - Timestep + trigger_ventopenduration(room, j, vent)
                                    VentBreakTimeClose(room, j, vent) = tim(i, 1) - Timestep + trigger_ventopenduration(room, j, vent)

                                    Dim Message As String = CStr(VentOpenTime(room, j, vent)) & " sec. Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened " & trigger_ventopendelay(room, j, vent).ToString & " sec following heat detector operation."
                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                                    Message = CStr(VentCloseTime(room, j, vent)) & " sec. Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " closed"
                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text
                                End If
                            End If
                        End If

                        'if fire size criteria met
                        If trigger_device(2, room, j, vent) = True And (room = fireroom Or j = fireroom) Then
                            If HeatRelease(fireroom, i, 2) > HRR_threshold(room, j, vent) And HRRFlag(room, j, vent) = 0 Then '2009.22
                                HRR_threshold_time(room, j, vent) = tim(i, 1)
                                HRRFlag(room, j, vent) = 1
                            End If
                            If tim(i, 1) > HRR_threshold_time(room, j, vent) + HRR_ventopendelay(room, j, vent) And HRRFlag(room, j, vent) = 1 Then
                                breakflag(room, j, vent) = True
                                VentilationLimitFlag = False 'reset
                                PeakHRR = originalpeakhrr
                                VentOpenTime(room, j, vent) = tim(i, 1) - Timestep
                                VentBreakTime(room, j, vent) = tim(i, 1) - Timestep
                                VentCloseTime(room, j, vent) = tim(i, 1) - Timestep + HRR_ventopenduration(room, j, vent)
                                VentBreakTimeClose(room, j, vent) = tim(i, 1) - Timestep + HRR_ventopenduration(room, j, vent)

                                Dim Message As String = CStr(VentOpenTime(room, j, vent)) & " sec. Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened when HRR reached " & HRR_threshold(room, j, vent) & " kW"
                                'MDIFrmMain.ToolStripStatusLabel2.Text = Message.ToString
                                frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                            End If
                        End If

                        'if flashover reached
                        If trigger_device(3, room, j, vent) = True And breakflag(room, j, vent) = False And (room = fireroom Or j = fireroom) Then
                            If Flashover = True Then
                                breakflag(room, j, vent) = True
                                VentilationLimitFlag = False 'reset
                                PeakHRR = originalpeakhrr
                                VentOpenTime(room, j, vent) = tim(i, 1)
                                VentBreakTime(room, j, vent) = tim(i, 1)

                                Dim Message As String = CStr(tim(i, 1)) & " sec. Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened at flashover."
                                'MDIFrmMain.ToolStripStatusLabel2.Text = Message.ToString
                                frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                            End If
                        End If

                        'if ventilation limit reached
                        If trigger_device(4, room, j, vent) = True And breakflag(room, j, vent) = False And (room = fireroom Or j = fireroom) Then
                            If VentilationLimitFlag = True Then
                                If tim(i, 1) > alphaTfitted(4, itcounter - 1) Then 'allow enough time for the vent to fully open

                                    breakflag(room, j, vent) = True
                                    VentilationLimitFlag = False 'reset (when we open vent, fire might not be vent limited any more, so will need to retest
                                    PeakHRR = originalpeakhrr
                                    VentOpenTime(room, j, vent) = tim(i, 1)
                                    VentBreakTime(room, j, vent) = tim(i, 1)

                                    Dim Message As String = CStr(tim(i, 1)) & " sec. Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened when ventilation limit reached."
                                    'MDIFrmMain.ToolStripStatusLabel2.Text = Message.ToString
                                    frmInputs.rtb_log.Text = Message.ToString & Chr(13) & frmInputs.rtb_log.Text

                                    For r = j To NumberRooms + 1
                                        'For s = vent + 1 To NumberVents(room, r)
                                        For s = 1 To NumberVents(room, r)
                                            If trigger_device(4, room, r, s) = True And breakflag(room, r, s) = False And AutoOpenVent(room, r, s) = True Then
                                                breakflag(room, r, s) = True
                                                VentOpenTime(room, r, s) = tim(i, 1)
                                                VentBreakTime(room, r, s) = tim(i, 1)

                                                Dim Message1 As String = CStr(tim(i, 1)) & " sec. Vent " & CStr(room) & "-" & CStr(r) & "-" & CStr(s) & " opened when ventilation limit reached."
                                                frmInputs.rtb_log.Text = Message1.ToString & Chr(13) & frmInputs.rtb_log.Text

                                            End If
                                        Next
                                    Next
                                End If
                            End If
                        End If

                    Else
                        'manual open/close?
                        If breakflag(room, j, vent) = False And (VentCloseTime(room, j, vent) > VentOpenTime(room, j, vent)) Then
                            If tim(i, 1) >= VentOpenTime(room, j, vent) Then
                                breakflag(room, j, vent) = True

                                Dim Message1 As String = CStr(tim(i, 1)) & " sec. Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " opened by user."
                                frmInputs.rtb_log.Text = Message1.ToString & Chr(13) & frmInputs.rtb_log.Text
                                Message1 = CStr(tim(i, 1) + VentCloseTime(room, j, vent) - VentOpenTime(room, j, vent)) & " sec. Vent " & CStr(room) & "-" & CStr(j) & "-" & CStr(vent) & " closed by user."
                                frmInputs.rtb_log.Text = Message1.ToString & Chr(13) & frmInputs.rtb_log.Text
                            End If
                        End If

                    End If
                Next vent
            End If
        Next

    End Sub
End Module