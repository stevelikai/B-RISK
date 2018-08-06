Option Strict Off
Option Explicit On
Imports System.Math

Module spillplumemod

    Public Sub spillplume_2013(ByVal minwidth As Single, ByVal dteps As Single, ByRef Mass_Upper() As Double, ByRef vententrain(,,) As Double, ByVal vent As Integer, ByVal sroom As Integer, ByVal NProd As Short, ByRef prod1L() As Double, ByRef prod2L() As Double, ByRef prod2U() As Double, ByRef prod1U() As Double, ByRef conu(,) As Double, ByRef conl(,) As Double, ByRef TU() As Double, ByRef TL() As Double, ByRef cpu() As Double, ByRef cpl() As Double, ByVal room As Short, ByRef e1l As Double, ByRef e1u As Double, ByRef e2l As Double, ByRef e2u As Double, ByRef r2u As Double, ByRef r2l As Double, ByRef r1u As Double, ByRef r1l As Double, ByRef ylay() As Double, ByRef UFLW2(,,) As Double, ByRef ventfire As Double, ByRef mw_upper() As Double, ByRef mw_lower() As Double, ByRef weighted_hc As Double, ByVal nelev As Integer, ByRef dp1m2() As Double, ByRef yelev() As Double)
        '=============================================================================
        ' Subroutine to calculate entrainment
        ' including ventfires
        '
        '   vent = identifies which vent (=m)
        '   sroom = source room (=i)
        '   room = mass_destination room (=j)
        '
        '   this routine is only called when ylay(1) < ylay(2)
        '   this routine is only called where the destination room is an inside room 

        '   ylay(1) - variable holding layer height in the source room (m) relative to ref datum
        '   TU(1) - variable holding the upper layer temperature in the source room (K)
        '   cpu(1) - variable holding the specific heat of the upper layer gases in the source room (kJ/kgK)
        '=============================================================================

        Dim mflow, ventfire_new, QPlume As Double
        Dim W As Single
        Dim vent_entrain, QEQ As Double
        Dim h As Integer
        Dim Zb, ds, ztrans, Ms, Mb As Double
        Static warning As Boolean
        If stepcount = 1 Then warning = False

        'entrainment in door vent

        'flow passing through vent,
        If sroom = fireroom Then Ms = UFLW2(1, 1, 2) ' to/from upper layer of room 1
        If room = fireroom Then Ms = UFLW2(2, 1, 2) 'to/from upper layer of room 2

        'this subroutine will only calculate spill plume entrainment if one of the rooms is the fire room

        'flow is to an inside room
        If Ms < 0 And sroom = fireroom Then 'flow to/from upper layer of room 1
            'flow leaving upper layer of room 1 (-ve) and entering room 2
            'vent flow will entrain air from the lower layer of the adjacent room into the upper layer of the adjacent room
tryagain1:

            W = VentWidth(sroom, room, vent)

            'entrainment height in the destination room
            'measured from the spill edge to the underside of layer
            'ylay is relative to the ref datum not floor level
            If spillbalconyprojection(sroom, room, vent) > 0 Then
                'with projection 
                'rise from projection/ceiling level to underside of smoke layer
                Zb = ylay(2) - (RoomHeight(sroom) + FloorElevation(sroom))
            Else
                'no projection
                'rise from top of opening to underside of smoke layer
                Zb = ylay(2) - (VentHeight(sroom, room, vent) + VentSillHeight(sroom, room, vent) + FloorElevation(sroom))
            End If
            '********change this
            'Zb = ylay(2) - (VentHeight(sroom, room, vent) + VentSillHeight(sroom, room, vent) + FloorElevation(sroom))

            ds = RoomHeight(sroom) + FloorElevation(sroom) - ylay(1) 'smoke layer depth in source room

            ztrans = 3.4 * (W ^ (2 / 3) + 1.56 * ds ^ (2 / 3)) ^ (3 / 2)

            'enthalpy or convective flux associated with the mass flow leaving through the compartment opening
            QEQ = (cpu(1) * TU(1) - cpl(2) * TL(2)) * Abs(Ms) + ventfire 'kw

            If QEQ < 0 Then QEQ = 0

            'height of "plume"
            If QEQ > 0 Then
                Dim angle As Single = 30
                Dim radians As Single = angle * (PI / 180)
                'minwidth is the lesser of the length or width of the room
                'assuming the plume is close to one wall, ymax gives a upper limit estimate 
                'of the height at which one side of the plume intercepts the wall 
                'plume assumed to diverge at an angle of 30 deg from vertical
                Dim ymax As Double = minwidth / Tan(radians)

                If Zb > ymax Then
                    'limit the entrainment height when plume (30 deg to vertical intercepts the wall of the shaft before reaching the ceiling)
                    Zb = ymax
                Else
                    'Stop
                End If
                'doesn't affect the calculated layer height in R2 ?
            End If

            'for determining any additional entrainment between the downstand and the spill edge of a balcony or soffit
            If spillbalconyprojection(sroom, room, vent) <= 0 Then
                Mb = Abs(Ms) 'no additional entrainment
            Else
                Mb = 0.89 * (VentHeight(sroom, room, vent) / W) ^ (-0.92) * RoomHeight(sroom) * Abs(Ms) / W
                'only applies where a horizontally flowing layer forms beneath the higher balcony
                'and the flow is channelled by screens or walls
                'elevation of balcony and fire compartment ceiling to be the same
                'check limits of applicability?
            End If

            If spillplumemodel(sroom, room, vent) = 1 Then
                '2D adhered plume
                mflow = 0.08 * (QEQ) ^ (1 / 3) * (W ^ (2 / 3)) * Zb + 1.34 * Abs(Mb)
            ElseIf spillplumemodel(sroom, room, vent) = 2 Then
                '2D balcony 
                mflow = 0.16 * (QEQ) ^ (1 / 3) * (W ^ (2 / 3)) * Zb + 1.34 * Abs(Mb)
            ElseIf spillplumemodel(sroom, room, vent) = 3 Then
                '3D adhered
                If Zb > ztrans Then
                    'use axisymmetric plume equation
                    mflow = 0.071 * (QEQ) ^ (1 / 3) * Zb ^ (5 / 3)
                ElseIf W / ds <= 13 Then
                    '3D adhered plume
                    mflow = 0.3 * (QEQ) ^ (1 / 3) * (W ^ (1 / 6)) * ds ^ (1 / 2) * Zb + 1.34 * Abs(Mb)
                Else
                    'use 2D adhered plume
                    mflow = 0.08 * (QEQ) ^ (1 / 3) * (W ^ (2 / 3)) * Zb + 1.34 * Abs(Mb)
                End If

            ElseIf spillplumemodel(sroom, room, vent) = 4 Then
                '3D channelled balcony

                If Zb <= ztrans Then
                    mflow = 0.16 * (QEQ) ^ (1 / 3) * (W ^ (2 / 3) + 1.56 * ds ^ (2 / 3)) * Zb + 1.34 * Mb
                    'all ok
                Else
                    'use axisymmetric plume equation
                    mflow = 0.071 * (QEQ) ^ (1 / 3) * Zb ^ (5 / 3)
                End If

            ElseIf spillplumemodel(sroom, room, vent) = 5 Then
                '3D unchannelled balcony
                Mb = Abs(Ms)

                If W >= 2 * spillbalconyprojection(sroom, room, vent) Then
                    mflow = 0.16 * (QEQ) ^ (1 / 3) * ((W + spillbalconyprojection(sroom, room, vent)) ^ (2 / 3) + 1.56 * ds ^ (2 / 3)) * Zb + 1.34 * Mb
                Else
                    'out of range
                    mflow = 0
                    'validate input on entering data so this shouldn't happen
                End If
            Else
            End If

            vent_entrain = mflow - Abs(Ms) 'the entrained portion 
            '(mass flow taken from the lower layer of dest. room and added to the upper layer of dest. room)

            If vent_entrain < 0 Then vent_entrain = 0

            If ventfire > 0 Then
                ventfire_new = O2_limit_cfast(room, Mass_Upper(room), ventfire, vent_entrain, TU(2), conu(1, 2), conl(1, 2), mw_upper(room), mw_lower(room), QPlume)
                If Abs(ventfire_new - ventfire) / ventfire > 1 Then
                    ventfire = ventfire_new
                    GoTo tryagain1
                End If
            End If

            'store entrainment into new variable
            vententrain(sroom, room, vent) = vent_entrain

            'track species etc into next room
            For h = 3 To NProd + 2
                prod2U(h) = prod2U(h) + conl(h - 2, 2) * vent_entrain
                prod2L(h) = prod2L(h) - conl(h - 2, 2) * vent_entrain
            Next h

            r2u = r2u + vent_entrain
            r2l = r2l - vent_entrain
            e2u = e2u + 1000 * cpl(2) * vent_entrain * TL(2) ' enthalpy
            e2l = e2l - 1000 * cpl(2) * vent_entrain * TL(2)

        ElseIf Ms < 0 And room = fireroom Then 'flow to/from upper layer of room 2
            'flow leaving upper layer of room 2 (-ve) and entering room 1
            'vent flow will entrain air from the lower layer of the adjacent room 1 into the upper layer of the adjacent room 1

tryagain2:

            W = VentWidth(sroom, room, vent)

            'entrainment height in the destination room
            'measured from the spill edge to the underside of layer
            If spillbalconyprojection(sroom, room, vent) > 0 Then
                'with projection 
                Zb = ylay(1) - (RoomHeight(room) + FloorElevation(sroom))
            Else
                'no projection
                Zb = ylay(1) - (VentHeight(sroom, room, vent) + VentSillHeight(sroom, room, vent) + FloorElevation(sroom))
            End If

            '********change this
            'Zb = ylay(2) - (VentHeight(sroom, room, vent) + VentSillHeight(sroom, room, vent) + FloorElevation(sroom))

            ds = RoomHeight(room) + FloorElevation(room) - ylay(2) 'smoke layer depth in fire room

            ztrans = 3.4 * (W ^ (2 / 3) + 1.56 * ds ^ (2 / 3)) ^ (3 / 2)

            'enthalpy or convective flux associated with the mass flow leaving through the compartment opening
            QEQ = (cpu(2) * TU(2) - cpl(1) * TL(1)) * Abs(Ms) + ventfire 'kw

            If QEQ < 0 Then QEQ = 0

            'height of "plume"
            If QEQ > 0 Then
                Dim angle As Single = 30
                Dim radians As Single = angle * (PI / 180)
                'minwidth is the lesser of the length or width of the room
                'assuming the plume is close to one wall, ymax gives a upper limit estimate 
                'of the height at which one side of the plume intercepts the wall 
                'plume assumed to diverge at an angle of 30 deg from vertical
                Dim ymax As Double = minwidth / Tan(radians)

                If Zb > ymax Then
                    'limit the entrainment height when plume (30 deg to vertical intercepts the wall of the shaft before reaching the ceiling)
                    Zb = ymax
                Else
                    'Stop
                End If
                'doesn't affect the calculated layer height in R2 ?
            End If

            'for determining any additional entrainment between the downstand and the spill edge of a balcony or soffit
            If spillbalconyprojection(sroom, room, vent) <= 0 Then
                Mb = Abs(Ms) 'no additional entrainment
            Else
                Mb = 0.89 * (VentHeight(sroom, room, vent) / W) ^ (-0.92) * RoomHeight(room) * Abs(Ms) / W
                'only applies where a horizontally flowing layer forms beneath the higher balcony
                'and the flow is channelled by screens or walls
                'elevation of balcony and fire compartment ceiling to be the same
                'check limits of applicability?
            End If

            If spillplumemodel(sroom, room, vent) = 1 Then
                '2D adhered plume
                mflow = 0.08 * (QEQ) ^ (1 / 3) * (W ^ (2 / 3)) * Zb + 1.34 * Abs(Mb)
            ElseIf spillplumemodel(sroom, room, vent) = 2 Then
                '2D balcony 
                mflow = 0.16 * (QEQ) ^ (1 / 3) * (W ^ (2 / 3)) * Zb + 1.34 * Abs(Mb)
            ElseIf spillplumemodel(sroom, room, vent) = 3 Then
                '3D adhered
                If Zb > ztrans Then
                    'use axisymmetric plume equation
                    mflow = 0.071 * (QEQ) ^ (1 / 3) * Zb ^ (5 / 3)
                ElseIf W / ds <= 13 Then
                    '3D adhered plume
                    mflow = 0.3 * (QEQ) ^ (1 / 3) * (W ^ (1 / 6)) * ds ^ (1 / 2) * Zb + 1.34 * Abs(Mb)
                Else
                    'use 2D adhered plume
                    mflow = 0.08 * (QEQ) ^ (1 / 3) * (W ^ (2 / 3)) * Zb + 1.34 * Abs(Mb)
                End If

            ElseIf spillplumemodel(sroom, room, vent) = 4 Then
                '3D channelled balcony

                If Zb <= ztrans Then
                    mflow = 0.16 * (QEQ) ^ (1 / 3) * (W ^ (2 / 3) + 1.56 * ds ^ (2 / 3)) * Zb + 1.34 * Mb
                    'all ok
                Else
                    'use axisymmetric plume equation
                    mflow = 0.071 * (QEQ) ^ (1 / 3) * Zb ^ (5 / 3)
                End If

            ElseIf spillplumemodel(sroom, room, vent) = 5 Then
                '3D unchannelled balcony
                Mb = Abs(Ms)

                If W >= 2 * spillbalconyprojection(sroom, room, vent) Then
                    mflow = 0.16 * (QEQ) ^ (1 / 3) * ((W + spillbalconyprojection(sroom, room, vent)) ^ (2 / 3) + 1.56 * ds ^ (2 / 3)) * Zb + 1.34 * Mb
                Else
                    'out of range
                    mflow = 0
                    'validate input on entering data so this shouldn't happen
                End If
            Else
            End If

            vent_entrain = mflow - Abs(Ms) 'the entrained portion 
            '(mass flow taken from the lower layer of dest. room and added to the upper layer of dest. room)

            If vent_entrain < 0 Then vent_entrain = 0

            If ventfire > 0 Then

                ventfire_new = O2_limit_cfast(sroom, Mass_Upper(sroom), ventfire, vent_entrain, TU(1), conu(1, 1), conl(1, 1), mw_upper(sroom), mw_lower(sroom), QPlume)
                If Abs(ventfire_new - ventfire) / ventfire > 1 Then
                    ventfire = ventfire_new
                    GoTo tryagain2
                End If
            End If

            'store entrainment into new variable
            vententrain(room, sroom, vent) = vent_entrain

            'here
            'track species etc into next room
            For h = 3 To NProd + 2
                prod1U(h) = prod1U(h) + conl(h - 2, 1) * vent_entrain
                prod1L(h) = prod1L(h) - conl(h - 2, 1) * vent_entrain
            Next h

            r1u = r1u + vent_entrain
            r1l = r1l - vent_entrain
            e1u = e1u + 1000 * cpl(1) * vent_entrain * TL(1) ' enthalpy
            e1l = e1l - 1000 * cpl(1) * vent_entrain * TL(1)


        ElseIf Ms >= 0 And room = fireroom Then
            'flow from 2-1 but no entrainment of lower layer
            vent_entrain = 0

        ElseIf Ms >= 0 And sroom = fireroom Then
            'flow from 2-1 but no entrainment of lower layer
            'UFLW2(1, 1, 2) > 0 means that flow is from room 2 to room 1
            vent_entrain = 0

        Else
        End If

    End Sub
    Public Sub spillplume_2014(ByVal minwidth As Single, ByVal dteps As Single, ByRef Mass_Upper() As Double, ByRef vententrain(,,) As Double, ByVal vent As Integer, ByVal sroom As Integer, ByVal NProd As Short, ByRef prod1L() As Double, ByRef prod2L() As Double, ByRef prod2U() As Double, ByRef prod1U() As Double, ByRef conu(,) As Double, ByRef conl(,) As Double, ByRef TU() As Double, ByRef TL() As Double, ByRef cpu() As Double, ByRef cpl() As Double, ByVal room As Short, ByRef e1l As Double, ByRef e1u As Double, ByRef e2l As Double, ByRef e2u As Double, ByRef r2u As Double, ByRef r2l As Double, ByRef r1u As Double, ByRef r1l As Double, ByRef ylay() As Double, ByRef UFLW2(,,) As Double, ByRef ventfire As Double, ByRef mw_upper() As Double, ByRef mw_lower() As Double, ByRef weighted_hc As Double, ByVal nelev As Integer, ByRef dp1m2() As Double, ByRef yelev() As Double)
        '=============================================================================
        ' Subroutine to calculate entrainment
        ' including ventfires
        '
        '   vent = identifies which vent (=m)
        '   sroom = source room 1 (=i)
        '   room = mass_destination room 2(=j)
        '
        '   this routine is only called where the destination room is an inside room 

        '   ylay(1) - variable holding layer height in the source room (m) relative to ref datum
        '   TU(1) - variable holding the upper layer temperature in the source room (K)
        '   cpu(1) - variable holding the specific heat of the upper layer gases in the source room (kJ/kgK)
        '=============================================================================

        Dim mflow, ventfire_new, QPlume As Double
        Dim W As Single
        Dim vent_entrain, QEQ As Double
        Dim h As Integer
        Dim Zb, ds, ztrans, Ms, Ms1, Ms2, Mb As Double
        Static warning As Boolean
        If stepcount = 1 Then warning = False

        'entrainment in door vent
        'we will only model entrainment when the layer receiving the flow is higher than in the source room

        If ylay(1) > ylay(2) Then
            'entrainment in room 2
            'flow passing through vent
            Ms = UFLW2(2, 1, 2) 'to/from upper layer of room 2

            If Ms < 0 Then 'flow to/from upper layer of room 2
                'flow leaving upper layer of room 2 (-ve) and entering room 1
                'vent flow will entrain air from the lower layer of room 1 into the upper layer of room 1
tryagain2:
                W = VentWidth(sroom, room, vent)

                'entrainment height in the destination room
                'measured from the spill edge to the underside of layer
                If spillbalconyprojection(sroom, room, vent) > 0 Then 'which side of vent is the projection, assume both sides here
                    'with projection 
                    Zb = ylay(1) - (RoomHeight(sroom) + FloorElevation(sroom))
                Else
                    'no projection
                    Zb = ylay(1) - (VentHeight(sroom, room, vent) + VentSillHeight(sroom, room, vent) + FloorElevation(sroom))
                End If

                ds = RoomHeight(room) + FloorElevation(room) - ylay(2) 'smoke layer depth in fire room

                ztrans = 3.4 * (W ^ (2 / 3) + 1.56 * ds ^ (2 / 3)) ^ (3 / 2)

                'enthalpy or convective flux associated with the mass flow leaving through the compartment opening
                QEQ = (cpu(2) * TU(2) - cpl(1) * TL(1)) * Abs(Ms) + ventfire 'kw

                If QEQ < 0 Then QEQ = 0

                'height of "plume"
                If QEQ > 0 Then
                    Dim angle As Single = 30
                    Dim radians As Single = angle * (PI / 180)
                    'minwidth is the lesser of the length or width of the room
                    'assuming the plume is close to one wall, ymax gives a upper limit estimate 
                    'of the height at which one side of the plume intercepts the wall 
                    'plume assumed to diverge at an angle of 30 deg from vertical
                    Dim ymax As Double = minwidth / Tan(radians)

                    If Zb > ymax Then
                        'limit the entrainment height when plume (30 deg to vertical intercepts the wall of the shaft before reaching the ceiling)
                        Zb = ymax
                    Else
                        'Stop
                    End If
                    'doesn't affect the calculated layer height in R2 ?
                End If

                'for determining any additional entrainment between the downstand and the spill edge of a balcony or soffit
                If spillbalconyprojection(sroom, room, vent) <= 0 Then
                    Mb = Abs(Ms) 'no additional entrainment
                Else
                    Mb = 0.89 * (VentHeight(sroom, room, vent) / W) ^ (-0.92) * RoomHeight(room) * Abs(Ms) / W
                    'only applies where a horizontally flowing layer forms beneath the higher balcony
                    'and the flow is channelled by screens or walls
                    'elevation of balcony and fire compartment ceiling to be the same
                    'check limits of applicability?
                End If

                If spillplumemodel(sroom, room, vent) = 1 Then
                    '2D adhered plume
                    mflow = 0.08 * (QEQ) ^ (1 / 3) * (W ^ (2 / 3)) * Zb + 1.34 * Abs(Mb)
                ElseIf spillplumemodel(sroom, room, vent) = 2 Then
                    '2D balcony 
                    mflow = 0.16 * (QEQ) ^ (1 / 3) * (W ^ (2 / 3)) * Zb + 1.34 * Abs(Mb)
                ElseIf spillplumemodel(sroom, room, vent) = 3 Then
                    '3D adhered
                    If Zb > ztrans Then
                        'use axisymmetric plume equation
                        mflow = 0.071 * (QEQ) ^ (1 / 3) * Zb ^ (5 / 3)
                    ElseIf W / ds <= 13 Then
                        '3D adhered plume
                        mflow = 0.3 * (QEQ) ^ (1 / 3) * (W ^ (1 / 6)) * ds ^ (1 / 2) * Zb + 1.34 * Abs(Mb)
                    Else
                        'use 2D adhered plume
                        mflow = 0.08 * (QEQ) ^ (1 / 3) * (W ^ (2 / 3)) * Zb + 1.34 * Abs(Mb)
                    End If

                ElseIf spillplumemodel(sroom, room, vent) = 4 Then
                    '3D channelled balcony

                    If Zb <= ztrans Then
                        mflow = 0.16 * (QEQ) ^ (1 / 3) * (W ^ (2 / 3) + 1.56 * ds ^ (2 / 3)) * Zb + 1.34 * Mb
                        'all ok
                    Else
                        'use axisymmetric plume equation
                        mflow = 0.071 * (QEQ) ^ (1 / 3) * Zb ^ (5 / 3)
                    End If

                ElseIf spillplumemodel(sroom, room, vent) = 5 Then
                    '3D unchannelled balcony
                    Mb = Abs(Ms)

                    If W >= 2 * spillbalconyprojection(sroom, room, vent) Then
                        mflow = 0.16 * (QEQ) ^ (1 / 3) * ((W + spillbalconyprojection(sroom, room, vent)) ^ (2 / 3) + 1.56 * ds ^ (2 / 3)) * Zb + 1.34 * Mb
                    Else
                        'out of range
                        mflow = 0
                        'validate input on entering data so this shouldn't happen
                    End If
                Else
                End If

                vent_entrain = mflow - Abs(Ms) 'the entrained portion 
                '(mass flow taken from the lower layer of dest. room and added to the upper layer of dest. room)

                If vent_entrain < 0 Then vent_entrain = 0

                If ventfire > 0 Then

                    ventfire_new = O2_limit_cfast(sroom, Mass_Upper(sroom), ventfire, vent_entrain, TU(1), conu(1, 1), conl(1, 1), mw_upper(sroom), mw_lower(sroom), QPlume)
                    If Abs(ventfire_new - ventfire) / ventfire > 1 Then
                        ventfire = ventfire_new
                        GoTo tryagain2
                    End If
                End If

                'store entrainment into new variable
                vententrain(room, sroom, vent) = vent_entrain

                'here
                'track species etc into next room
                For h = 3 To NProd + 2
                    prod1U(h) = prod1U(h) + conl(h - 2, 1) * vent_entrain
                    prod1L(h) = prod1L(h) - conl(h - 2, 1) * vent_entrain
                Next h

                r1u = r1u + vent_entrain
                r1l = r1l - vent_entrain
                e1u = e1u + 1000 * cpl(1) * vent_entrain * TL(1) ' enthalpy
                e1l = e1l - 1000 * cpl(1) * vent_entrain * TL(1)

            Else
                vent_entrain = 0
            End If

        ElseIf ylay(2) > ylay(1) Then
            'entrainment in room 2
            'flow passing through vent
            Ms = UFLW2(1, 1, 2) 'to/from upper layer of room 1

            If Ms < 0 Then 'flow to/from upper layer of room 1
                'flow leaving upper layer of room 1 (-ve) and entering room 2
                'vent flow will entrain air from the lower layer of room 2 into the upper layer of room 2
tryagain1:
                W = VentWidth(sroom, room, vent)

                'entrainment height in the destination room
                'measured from the spill edge to the underside of layer
                'ylay is relative to the ref datum not floor level
                If spillbalconyprojection(sroom, room, vent) > 0 Then
                    'with projection 
                    'rise from projection/ceiling level to underside of smoke layer
                    Zb = ylay(2) - (RoomHeight(sroom) + FloorElevation(sroom))
                Else
                    'no projection
                    'rise from top of opening to underside of smoke layer
                    Zb = ylay(2) - (VentHeight(sroom, room, vent) + VentSillHeight(sroom, room, vent) + FloorElevation(sroom))
                End If
              
                ds = RoomHeight(sroom) + FloorElevation(sroom) - ylay(1) 'smoke layer depth in source room

                ztrans = 3.4 * (W ^ (2 / 3) + 1.56 * ds ^ (2 / 3)) ^ (3 / 2)

                'enthalpy or convective flux associated with the mass flow leaving through the compartment opening
                QEQ = (cpu(1) * TU(1) - cpl(2) * TL(2)) * Abs(Ms) + ventfire 'kw

                If QEQ < 0 Then QEQ = 0

                'height of "plume"
                If QEQ > 0 Then
                    Dim angle As Single = 30
                    Dim radians As Single = angle * (PI / 180)
                    'minwidth is the lesser of the length or width of the room
                    'assuming the plume is close to one wall, ymax gives a upper limit estimate 
                    'of the height at which one side of the plume intercepts the wall 
                    'plume assumed to diverge at an angle of 30 deg from vertical
                    Dim ymax As Double = minwidth / Tan(radians)

                    If Zb > ymax Then
                        'limit the entrainment height when plume (30 deg to vertical intercepts the wall of the shaft before reaching the ceiling)
                        Zb = ymax
                    Else
                        'Stop
                    End If
                    'doesn't affect the calculated layer height in R2 ?
                End If

                'for determining any additional entrainment between the downstand and the spill edge of a balcony or soffit
                If spillbalconyprojection(sroom, room, vent) <= 0 Then
                    Mb = Abs(Ms) 'no additional entrainment
                Else
                    Mb = 0.89 * (VentHeight(sroom, room, vent) / W) ^ (-0.92) * RoomHeight(sroom) * Abs(Ms) / W
                    'only applies where a horizontally flowing layer forms beneath the higher balcony
                    'and the flow is channelled by screens or walls
                    'elevation of balcony and fire compartment ceiling to be the same
                    'check limits of applicability?
                End If

                If spillplumemodel(sroom, room, vent) = 1 Then
                    '2D adhered plume
                    mflow = 0.08 * (QEQ) ^ (1 / 3) * (W ^ (2 / 3)) * Zb + 1.34 * Abs(Mb)
                ElseIf spillplumemodel(sroom, room, vent) = 2 Then
                    '2D balcony 
                    mflow = 0.16 * (QEQ) ^ (1 / 3) * (W ^ (2 / 3)) * Zb + 1.34 * Abs(Mb)
                ElseIf spillplumemodel(sroom, room, vent) = 3 Then
                    '3D adhered
                    If Zb > ztrans Then
                        'use axisymmetric plume equation
                        mflow = 0.071 * (QEQ) ^ (1 / 3) * Zb ^ (5 / 3)
                    ElseIf W / ds <= 13 Then
                        '3D adhered plume
                        mflow = 0.3 * (QEQ) ^ (1 / 3) * (W ^ (1 / 6)) * ds ^ (1 / 2) * Zb + 1.34 * Abs(Mb)
                    Else
                        'use 2D adhered plume
                        mflow = 0.08 * (QEQ) ^ (1 / 3) * (W ^ (2 / 3)) * Zb + 1.34 * Abs(Mb)
                    End If

                ElseIf spillplumemodel(sroom, room, vent) = 4 Then
                    '3D channelled balcony

                    If Zb <= ztrans Then
                        mflow = 0.16 * (QEQ) ^ (1 / 3) * (W ^ (2 / 3) + 1.56 * ds ^ (2 / 3)) * Zb + 1.34 * Mb
                        'all ok
                    Else
                        'use axisymmetric plume equation
                        mflow = 0.071 * (QEQ) ^ (1 / 3) * Zb ^ (5 / 3)
                    End If

                ElseIf spillplumemodel(sroom, room, vent) = 5 Then
                    '3D unchannelled balcony
                    Mb = Abs(Ms)

                    If W >= 2 * spillbalconyprojection(sroom, room, vent) Then
                        mflow = 0.16 * (QEQ) ^ (1 / 3) * ((W + spillbalconyprojection(sroom, room, vent)) ^ (2 / 3) + 1.56 * ds ^ (2 / 3)) * Zb + 1.34 * Mb
                    Else
                        'out of range
                        mflow = 0
                        'validate input on entering data so this shouldn't happen
                    End If
                Else
                End If

                vent_entrain = mflow - Abs(Ms) 'the entrained portion 
                '(mass flow taken from the lower layer of dest. room and added to the upper layer of dest. room)

                If vent_entrain < 0 Then vent_entrain = 0

                If ventfire > 0 Then
                    ventfire_new = O2_limit_cfast(room, Mass_Upper(room), ventfire, vent_entrain, TU(2), conu(1, 2), conl(1, 2), mw_upper(room), mw_lower(room), QPlume)
                    If Abs(ventfire_new - ventfire) / ventfire > 1 Then
                        ventfire = ventfire_new
                        GoTo tryagain1
                    End If
                End If

                'store entrainment into new variable
                vententrain(sroom, room, vent) = vent_entrain

                'track species etc into next room
                For h = 3 To NProd + 2
                    prod2U(h) = prod2U(h) + conl(h - 2, 2) * vent_entrain
                    prod2L(h) = prod2L(h) - conl(h - 2, 2) * vent_entrain
                Next h

                r2u = r2u + vent_entrain
                r2l = r2l - vent_entrain
                e2u = e2u + 1000 * cpl(2) * vent_entrain * TL(2) ' enthalpy
                e2l = e2l - 1000 * cpl(2) * vent_entrain * TL(2)
            Else
                vent_entrain = 0
            End If

        End If

    End Sub
    Public Sub createsmokeviewdata()

        '===========================================================
        '   Saves results in a .csv file to be read by smokeview
        '===========================================================

        Dim s, Txt As String
        Dim room As Integer
        Dim k, j As Integer
        Dim count As Integer = 0
        Dim SaveBox As New SaveFileDialog()
        'Dim oFile As System.IO.File
        Dim oWrite As System.IO.StreamWriter

        Try
            If basefile = "" Then Exit Sub

            Dim SmokeDataFile = Mid(basefile, 1, Len(basefile) - 4) & "_zone.csv"

            SaveBox.FileName = SmokeDataFile
            'Display the Save as dialog box.

            If SaveBox.CheckFileExists = False Then
                'create the file
                My.Computer.FileSystem.WriteAllText(SaveBox.FileName, "", True)
            End If


            'oWrite = oFile.CreateText(SmokeDataFile)
            oWrite = IO.File.CreateText(SmokeDataFile)

            'define a format string
            's = "Scientific"
            s = "General Number"

            Txt = "s,"
            For i = 1 To NumberRooms
                Txt = Txt & "C, C, m, Pa, 1 / m, 1 / m,"
            Next
            For i = 1 To NumberObjects
                Txt = Txt & " kW, m, m, m ^ 2,"
            Next
            For i = 1 To number_vents
                Txt = Txt & " m ^ 2,"
            Next

            For room = 1 To NumberRooms + 1
                For i = 1 To NumberRooms + 1
                    'If room < i Then
                    If NumberCVents(i, room) > 0 Then
                        For k = 1 To NumberCVents(i, room)
                            Txt = Txt & " m ^ 2,"
                        Next
                    End If
                    'End If
                Next
            Next

            oWrite.WriteLine("{0,1}", Txt)

            Txt = "Time,"
            For i = 1 To NumberRooms
                Txt = Txt & "ULT_" & i.ToString & ",LLT_" & i.ToString & ",HGT_" & i.ToString & ",PRS_" & i.ToString & ",ULOD_" & i.ToString & ",LLOD_" & i.ToString & ","
            Next
            For i = 1 To NumberObjects
                Txt = Txt & "HRR_" & i.ToString & ",FLHGT_" & i.ToString & ",FBASE_" & i.ToString & ",FAREA_" & i.ToString & ","
            Next
            For i = 1 To number_vents
                Txt = Txt & "HVENT_" & i.ToString & ","
            Next

            Dim m As Integer = 0
            For room = 1 To NumberRooms + 1
                For i = 1 To NumberRooms + 1
                    'If room < i Then
                    If NumberCVents(i, room) > 0 Then
                        For k = 1 To NumberCVents(i, room)
                            m = m + 1
                            Txt = Txt & "VVENT_" & m.ToString & ","
                        Next
                    End If
                    'End If
                Next
            Next

            oWrite.WriteLine("{0,1}", Txt)

            For j = 1 To NumberTimeSteps + 1
                If Int(tim(j, 1) / ExcelInterval) - tim(j, 1) / ExcelInterval = 0 Then
                    Txt = VB6.Format(tim(j, 1), s) & ","
                    For i = 1 To NumberRooms
                        Txt = Txt & uppertemp(i, j) - 273 & ","
                        Txt = Txt & lowertemp(i, j) - 273 & ","
                        Txt = Txt & layerheight(i, j) & ","
                        Txt = Txt & RoomPressure(i, j) & ","
                        Txt = Txt & OD_upper(i, j) & ","
                        Txt = Txt & OD_lower(i, j) & ","
                    Next
                    For i = 1 To NumberObjects

                        Txt = Txt & HeatRelease(fireroom, j, 2) & ","
                        Txt = Txt & 0.235 * HeatRelease(fireroom, j, 2) ^ 0.4 - 1.2 * System.Math.Sqrt(4 * ObjLength(i) * ObjWidth(i) / PI) & "," 'FLHGT_
                        Txt = Txt & FireHeight(i) & "," 'FBASE_
                        Txt = Txt & ObjLength(i) * ObjWidth(i) & "," 'FAREA_
                        
                    Next

                    For room = 1 To NumberRooms
                        For i = 2 To NumberRooms + 1
                            If room < i Then
                                If NumberVents(room, i) > 0 Then
                                    For k = 1 To NumberVents(room, i)
                                        Txt = Txt & VentHeight(room, i, k) * VentWidth(room, i, k) & ","
                                    Next
                                End If
                            End If
                        Next
                    Next

                    For room = 1 To NumberRooms + 1
                        For i = 1 To NumberRooms + 1
                            'If room < i Then
                            If NumberCVents(i, room) > 0 Then
                                For k = 1 To NumberCVents(i, room)
                                    Txt = Txt & CVentArea(i, room, k) & ","
                                Next
                            End If
                            'End If
                        Next
                    Next

                    oWrite.WriteLine("{0,1}", Txt)

                End If
            Next

            oWrite.Close()
            oWrite.Dispose()

            'MsgBox("Data saved in " & SaveBox.FileName, MsgBoxStyle.Information + MsgBoxStyle.OkOnly)

            Exit Sub

        Catch ex As Exception
            oWrite.Close()
            oWrite.Dispose()
        End Try
    End Sub
    Public Sub createsmokeviewdata_copy()

        '===========================================================
        '   Saves results in a .csv file to be read by smokeview
        '===========================================================

        Dim s, Txt As String
        Dim room As Integer
        Dim k, j As Integer
        Dim count As Integer = 0
        Dim SaveBox As New SaveFileDialog()
        'Dim oFile As System.IO.File
        Dim oWrite As System.IO.StreamWriter

        Try
            If basefile = "" Then Exit Sub

            Dim SmokeDataFile = Mid(basefile, 1, Len(basefile) - 4) & "_zone.csv"

            SaveBox.FileName = SmokeDataFile
            'Display the Save as dialog box.

            If SaveBox.CheckFileExists = False Then
                'create the file
                My.Computer.FileSystem.WriteAllText(SaveBox.FileName, "", True)
            End If


            'oWrite = oFile.CreateText(SmokeDataFile)
            oWrite = IO.File.CreateText(SmokeDataFile)

            'define a format string
            's = "Scientific"
            s = "General Number"

            Txt = "s,"
            For i = 1 To NumberRooms
                Txt = Txt & "C, C, m, Pa, 1 / m, 1 / m,"
            Next
            For i = 1 To NumberObjects
                Txt = Txt & " kW, m, m, m ^ 2,"
            Next
            For i = 1 To number_vents
                Txt = Txt & " m ^ 2,"
            Next

            For room = 1 To NumberRooms + 1
                For i = 1 To NumberRooms + 1
                    'If room < i Then
                    If NumberCVents(i, room) > 0 Then
                        For k = 1 To NumberCVents(i, room)
                            Txt = Txt & " m ^ 2,"
                        Next
                    End If
                    'End If
                Next
            Next

            oWrite.WriteLine("{0,1}", Txt)

            Txt = "Time,"
            For i = 1 To NumberRooms
                Txt = Txt & "ULT_" & i.ToString & ",LLT_" & i.ToString & ",HGT_" & i.ToString & ",PRS_" & i.ToString & ",ULOD_" & i.ToString & ",LLOD_" & i.ToString & ","
            Next
            For i = 1 To NumberObjects
                Txt = Txt & "HRR_" & i.ToString & ",FLHGT_" & i.ToString & ",FBASE_" & i.ToString & ",FAREA_" & i.ToString & ","
            Next
            For i = 1 To number_vents
                Txt = Txt & "HVENT_" & i.ToString & ","
            Next

            Dim m As Integer = 0
            For room = 1 To NumberRooms + 1
                For i = 1 To NumberRooms + 1
                    'If room < i Then
                    If NumberCVents(i, room) > 0 Then
                        For k = 1 To NumberCVents(i, room)
                            m = m + 1
                            Txt = Txt & "VVENT_" & m.ToString & ","
                        Next
                    End If
                    'End If
                Next
            Next

            oWrite.WriteLine("{0,1}", Txt)

            For j = 1 To NumberTimeSteps + 1
                If Int(tim(j, 1) / ExcelInterval) - tim(j, 1) / ExcelInterval = 0 Then
                    Txt = VB6.Format(tim(j, 1), s) & ","
                    For i = 1 To NumberRooms
                        Txt = Txt & VB6.Format(uppertemp(i, j) - 273, s) & ","
                        Txt = Txt & VB6.Format(lowertemp(i, j) - 273, s) & ","
                        Txt = Txt & VB6.Format(layerheight(i, j), s) & ","
                        Txt = Txt & (RoomPressure(i, j)) & ","
                        Txt = Txt & (OD_upper(i, j)) & ","
                        Txt = Txt & OD_lower(i, j).ToString & ","
                    Next
                    For i = 1 To NumberObjects

                        Txt = Txt & VB6.Format(HeatRelease(fireroom, j, 2), s) & ","
                        Txt = Txt & VB6.Format(0.235 * HeatRelease(fireroom, j, 2) ^ 0.4 - 1.2 * System.Math.Sqrt(4 * ObjLength(i) * ObjWidth(i) / PI), s) & "," 'FLHGT_
                        Txt = Txt & VB6.Format(FireHeight(i), s) & "," 'FBASE_
                        Txt = Txt & VB6.Format(ObjLength(i) * ObjWidth(i), s) & "," 'FAREA_

                    Next

                    For room = 1 To NumberRooms
                        For i = 2 To NumberRooms + 1
                            If room < i Then
                                If NumberVents(room, i) > 0 Then
                                    For k = 1 To NumberVents(room, i)
                                        Txt = Txt & VB6.Format(VentHeight(room, i, k) * VentWidth(room, i, k), s) & ","
                                    Next
                                End If
                            End If
                        Next
                    Next

                    For room = 1 To NumberRooms + 1
                        For i = 1 To NumberRooms + 1
                            'If room < i Then
                            If NumberCVents(i, room) > 0 Then
                                For k = 1 To NumberCVents(i, room)
                                    Txt = Txt & CVentArea(i, room, k) & ","
                                Next
                            End If
                            'End If
                        Next
                    Next

                    oWrite.WriteLine("{0,1}", Txt)

                End If
            Next

            oWrite.Close()
            oWrite.Dispose()

            'MsgBox("Data saved in " & SaveBox.FileName, MsgBoxStyle.Information + MsgBoxStyle.OkOnly)

            Exit Sub

        Catch ex As Exception
            oWrite.Close()
            oWrite.Dispose()
        End Try
    End Sub
End Module