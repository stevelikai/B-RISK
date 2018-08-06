Option Strict Off
Option Explicit On
Module cvent
    Public Sub VENTCF2A(ByVal avent As Double, ByRef conl(,) As Double, ByRef conu(,) As Double, ByVal NProd As Integer, ByRef TU() As Double, ByRef TL() As Double, ByRef cpu() As Double, ByRef cpl() As Double, ByRef yref() As Double, ByRef yceil() As Double, ByRef ylay() As Double, ByRef YVent As Double, ByRef DPRef() As Double, ByVal epsp As Double, ByRef denl() As Double, ByRef denu() As Double, ByRef C(,) As Double, ByRef cvent(,) As Double, ByRef XML() As Double, ByRef XMU() As Double, ByRef XMVent() As Double, ByRef PL(,) As Double, ByRef PU(,) As Double, ByRef PVent(,) As Double, ByRef P() As Double, ByRef QL() As Double, ByRef QU() As Double, ByRef QVent() As Double, ByRef T() As Double, ByRef Tvent() As Double, ByRef DELP1 As Double, ByRef Den() As Double, ByRef DischargeCoeff As Double)
        '===========================================================================
        '   Calculation of the flow of mass, enthalpy, oxygen and other products of
        '   combustion through a horizontal vent joing an upper  space 1 to a lower
        '   space 2. The subroutine uses input data describing the two-layer environment
        '   of inside rooms and a uniform environment in outside spaces. This is a variation
        '   of VENTCF2 which smoothes the rate of extraction from vent-adjacent layers when such layers are relatively thin.
        '   ref L Y Cooper FSJ 28 (1997) 253-287
        ' 28 January 2008
        '===========================================================================
        Dim Rho, VHighFL, GR, DELPFL, EPS As Double
        Dim room, NP As Integer
        Dim GG, SRDELP, COEFF, coeff_X, X0, FNOISE, FF As Double
        Dim DELDEN, DENBAR, V, W, TBAR, XMEW, EPSDEN As Double
        Dim sigma2, VBLow, DPDDPFL, D, VBHigh, VEXMAX, XM3 As Double
        Dim VST(2) As Double
        Dim VB(2) As Double
        Dim DP(2) As Double
        Dim epscut As Double
        Dim sigma1 As Double
        Dim XM As Double
        Dim VVent(2) As Double
        Dim Denvnt(2) As Double
        Dim APOWER As Double
        Dim DEL(2) As Double
        Dim DELY(2) As Double
        Dim FRACM(2) As Double
        Dim FRACQ(2) As Double
        Dim FRACP(,) As Double
        Dim temp, factor1 As Double
        Dim cp(2) As Double
        Dim cpvent(2) As Double

        ReDim FRACP(NProd, 2)

        If avent < gcd_Machine_Error Then Exit Sub 'avent =0 & assuming avent is never -ve

        GR = 0
        VHighFL = 0
        DELPFL = 0
        sigma2 = 1.045
        XM3 = -0.707
        APOWER = 1

        For room = 1 To 2
            VST(room) = 0
            VB(room) = 0
            QVent(room) = 0
            XMVent(room) = 0
            XML(room) = 0
            XMU(room) = 0
            QL(room) = 0
            QU(room) = 0
            Den(room) = 0
            T(room) = 0
            cp(room) = 0
            For NP = 1 To NProd
                PVent(NP, room) = 0
                PL(NP, room) = 0
                PU(NP, room) = 0
                C(NP, room) = 0
            Next NP
        Next room

        'calulate the P(room), DELP1, the other properties adjacent to the two sides of the vent and DELDEN
        'Room1 is always the upper room, therefore the pressure immediately above the vent is - this should always be very close to zero!
        If System.Math.Abs(YVent - yref(1)) > gcd_Machine_Error Then
            DP(1) = -G * denl(1) * (YVent - yref(1))
        Else
            DP(1) = 0.0#
        End If
        'Room2 is always the lower room, therefore the pressure below the vent is taken as the pressure at the layer interface
        If System.Math.Abs(YVent - ylay(2)) > gcd_Machine_Error Then
            DP(2) = -G * denl(2) * (ylay(2) - yref(2))
        Else
            DP(2) = -G * denl(2) * (YVent - yref(2))
        End If
        For room = 1 To 2
            '        If (yref(room) >= YVent + gcd_Machine_Error) And (YVent <= ylay(room) + gcd_Machine_Error) Then
            '            'the vent is at or below reference elevation in space 1
            '            'if space is an inside room, then both the vent and the layer interface are at floor elevation
            '            DP(room) = -G * denl(room) * (YVent - yref(room))
            '        Else
            '            'the vent is above the reference elevation in space 2
            '            'if space is an inside room then the vent is at the ceiling
            '            DP(room) = -G * (denl(room) * (ylay(room) - yref(room)) + denu(room) * (YVent - ylay(room)))
            '        End If
            P(room) = DPRef(room) + DP(room) + Atm_Pressure * 1000
        Next room

        'DELP1 is the pressure immediately below the upper/hot layer below the vent less the pressure immediately above vent
        DELP1 = P(2) - P(1)

        D = System.Math.Sqrt(4 * avent / PI) 'assumes the vent is circular

        For room = 1 To 2
            If (yceil(room) > yref(room)) Then
                DEL(room) = (yceil(room) - yref(room)) / 2
                If DEL(room) > D / 2 Then DEL(room) = D / 2 'DEL= the lesser of the radius of the vent or 1/2 floor to ceiling height
                DELY(room) = System.Math.Abs(ylay(room) - YVent) 'distance of the layer interface from the vent
            Else
                DEL(room) = D / 2
                DELY(room) = 0
            End If

            'compute effective near vent properties
            If (DELY(room) > DEL(room)) Then 'no plug-holing is assumed to occur
                If room = 1 Then 'properties are taken as the same as the lower layer
                    Den(1) = denl(1)
                    T(1) = TL(1)
                    cp(1) = cpl(1)
                    For NP = 1 To NProd
                        C(NP, 1) = conl(NP, 1)
                    Next NP
                Else 'room = 2, then properties are takes as the same as the upper layer
                    Den(2) = denu(2)
                    T(2) = TU(2)
                    cp(2) = cpu(2)
                    For NP = 1 To NProd
                        C(NP, 2) = conu(NP, 2)
                    Next NP
                End If
            Else 'plug-holing is assumed to occur
                If room = 1 Then
                    factor1 = (DELY(1) / DEL(1)) ^ APOWER
                    Den(1) = denu(1) * (1 - factor1) + denl(1) * factor1
                    T(1) = TU(1) * (1 - factor1) + TL(1) * factor1
                    cp(1) = cpu(1) * (1 - factor1) + cpl(1) * factor1
                    '                T(1) = TU(1) * denu(1) / Den(1)
                    '                cp(1) = cpu(1)
                    For NP = 1 To NProd
                        C(NP, 1) = (conu(NP, 1) * denu(1) * (1 - factor1) + conl(NP, 1) * denl(1) * factor1) / Den(1)
                    Next NP
                Else
                    factor1 = (DELY(2) / DEL(2)) ^ APOWER
                    Den(2) = denl(2) * (1 - factor1) + denu(2) * factor1
                    T(2) = TL(2) * (1 - factor1) + TU(2) * factor1
                    cp(2) = cpl(2) * (1 - factor1) + cpu(2) * factor1
                    '                T(2) = TL(2) * denl(2) / Den(2)
                    '                cp(2) = cpu(2)
                    For NP = 1 To NProd
                        C(NP, 2) = (conl(NP, 2) * denl(2) * (1 - factor1) + conu(NP, 2) * denu(2) * factor1) / Den(2)
                    Next NP
                End If
            End If
        Next room

        'DELDEN is density immediately above the vent less density immediately below vent
        DELDEN = Den(1) - Den(2)

        'calculate the standard volume flow rate through the vent into space 1 or 2
        'calculate VST(room) if DELP=0
        If System.Math.Abs(DELP1) < gcd_Machine_Error Then 'DELP1 = 0
            DELP1 = 0
            VST(1) = 0
            VST(2) = 0
        Else 'calculate VST(room) for nonzero DELP1
            If DELP1 > 0 Then
                VST(2) = 0
                Rho = Den(2)
                EPS = DELP1 / P(2)
            Else 'DELP1 < 0 Then
                VST(1) = 0
                Rho = Den(1)
                EPS = -DELP1 / P(1)
            End If

            coeff_X = 1 - EPS
            COEFF = DischargeCoeff + 0.17 * EPS 'DC can be changed by user
            'COEFF = 0.68 + 0.17 * EPS
            'COEFF = 0.6 + 0.25 * EPS 'from cooper paper 1997, not sure why it was changed to above - doesn't seem to change result much.
            X0 = 1
            temp = DPRef(2)
            If DPRef(1) > temp Then temp = DPRef(1)
            If X0 > temp Then temp = X0
            epscut = epsp * temp
            epscut = System.Math.Sqrt(epscut)
            SRDELP = System.Math.Sqrt(System.Math.Abs(DELP1))
            If ((SRDELP / epscut) <= 130) Then
                FNOISE = 1 - System.Math.Exp(-SRDELP / epscut)
            Else
                FNOISE = 1
            End If
            If EPS < 0.00001 Then
                W = 1 - 0.75 * EPS / gamma
            Else
                If (EPS < 1 - gamcut) Then
                    GG = coeff_X ^ (1 / gamma)
                    FF = System.Math.Sqrt((2 * gamma / (gamma - 1)) * GG * GG * (1 - coeff_X / GG))
                Else
                    FF = gammax
                End If
                W = FF / System.Math.Sqrt(2 * EPS)
            End If
            V = FNOISE * COEFF * W * System.Math.Sqrt(2 / Rho) * avent * SRDELP
            If DELP1 > 0 Then
                VST(1) = V
            Else 'DELP1 < 0
                VST(2) = V
            End If
        End If
        'when cross vent density configuration is unstable ie DELDEN >0, then calculate
        'the vent flow according to: cooper, combined pressure and buoyancy driven
        'flow through a horizontal vent nistir 5384. for stable configuration calculate
        'flow with standard model
        If DELDEN > 0 Then
            'for unstable configuration, now calculate the combined pressure and buoyancy driven
            'volume vent flow rates from high to low and from the low to high sides of the vent,
            'vbhigh and vblow respectively. from these calculate the combine pressure and buoyancy
            'driven volume flow rates vb(room) through the vent into room.
            TBAR = (T(1) + T(2)) / 2
            DENBAR = (Den(1) + Den(2)) / 2
            XMEW = 0.000000004128 * DENBAR * (TBAR ^ 2.5) / (TBAR + 110.4)
            EPSDEN = DELDEN / DENBAR
            GR = 2 * G * (D ^ 3) * EPSDEN / ((XMEW / DENBAR) ^ 2)
            If DELP1 > 0 Then EPSDEN = -EPSDEN
            VHighFL = 0.1754 * avent * System.Math.Sqrt(2 * G * D * System.Math.Abs(EPSDEN)) * System.Math.Exp(0.5536 * EPSDEN)
            DELPFL = 0.2427 * (4 * G * System.Math.Abs(EPSDEN * DENBAR) * D) * (1 + EPSDEN / 2) * System.Math.Exp(1.1072 * EPSDEN)
            DPDDPFL = System.Math.Abs(DELP1) / DELPFL '(when > 1 expect unidirectional flow, otherwise mixed flow)
            sigma1 = FNOISE * COEFF * W / 0.178
            If DPDDPFL >= 1 Then ' expect unidirectional flow
                VBLow = 0
                VBHigh = VHighFL * (1 - (sigma2 ^ 2) + System.Math.Sqrt(sigma2 ^ 4 + (sigma1 ^ 2) * (DPDDPFL - 1)))
            Else 'expect mixed flow
                VEXMAX = 0.055 * (4 / PI) * avent * System.Math.Sqrt(G * D * System.Math.Abs(EPSDEN))
                XM = (sigma1 / sigma2) ^ 2 - 1
                VBLow = VEXMAX * (((1 + XM3 / 2) * ((1 - DPDDPFL) ^ 2) - (2 + XM3 / 2) * (1 - DPDDPFL)) ^ 2)
                If System.Math.Abs(DPDDPFL) < gcd_Machine_Error Then
                    VBHigh = VBLow
                Else
                    VBHigh = VBLow + (XM - System.Math.Sqrt(1 + (XM ^ 2 - 1) * (1 - DPDDPFL))) * VHighFL / (XM - 1)
                End If
            End If
            If DELP1 > 0 Then
                VB(1) = VBHigh
                VB(2) = VBLow
            Else
                VB(1) = VBLow
                VB(2) = VBHigh
            End If
            'calculate VVent(room), volume flow rate through the vent into room
            If ((0 < GR) And (GR < 20000000.0#)) Then
                VVent(1) = VST(1) + (GR / 20000000.0#) * (VB(1) - VST(1))
                VVent(2) = VST(2) + (GR / 20000000.0#) * (VB(2) - VST(2))
            Else
                VVent(1) = VB(1)
                VVent(2) = VB(2)
            End If
        Else 'DELDEN <= 0 Then
            'calculate VVent(room), volume flow rate through the vent into room
            VVent(1) = VST(1)
            VVent(2) = VST(2)
        End If

        'calculate the vent flow properties
        Denvnt(1) = Den(2) '* P(1) / P(2)
        Denvnt(2) = Den(1) '* P(2) / P(1)
        Tvent(1) = T(2)
        Tvent(2) = T(1)
        cpvent(1) = cp(2)
        cpvent(2) = cp(1)
        For NP = 1 To NProd
            cvent(NP, 1) = C(NP, 2)
            cvent(NP, 2) = C(NP, 1)
        Next NP

        'calculate the vent mass flow rates
        For room = 1 To 2
            XMVent(room) = Denvnt(room) * VVent(room)
        Next room

        'calculate the rest of the vent flow rates
        For room = 1 To 2
            QVent(room) = XMVent(room) * cpvent(room) * Tvent(room) * 1000
            For NP = 1 To NProd
                PVent(NP, room) = XMVent(room) * cvent(NP, room)
            Next NP
        Next room

        'calculate the rate at which the vent flows add mass enthalpy and
        'products to the layers of the spaces from which they are extracted
        'first treat layers of space 1 and then space 2
        For room = 1 To 2
            FRACM(room) = 1
            FRACQ(room) = 1
            For NP = 1 To NProd
                FRACP(NP, room) = 1
            Next NP

            If DELY(room) > DEL(room) Then 'upper layer thickness is greater than 1/2 vent diameter & flow extracted from vent adjacent layer
                If room = 1 Then 'flow extracted from the lower layer of upper room
                    XML(1) = -XMVent(2)
                    XMU(1) = 0
                    QL(1) = -QVent(2)
                    QU(1) = 0
                    For NP = 1 To NProd
                        PL(NP, 1) = -PVent(NP, 2)
                        PU(NP, 1) = 0
                    Next NP
                Else 'room = 2, and flow extracted from the upper of the lower room
                    XMU(2) = -XMVent(1)
                    XML(2) = 0
                    QU(2) = -QVent(1)
                    QL(2) = 0
                    For NP = 1 To NProd
                        PU(NP, 2) = -PVent(NP, 1)
                        PL(NP, 2) = 0
                    Next NP
                End If
            Else 'flow extracted from both layers
                If room = 1 Then
                    factor1 = (DELY(1) / DEL(1)) ^ APOWER
                    XMU(1) = -denu(1) * (1 - factor1) * VVent(2)
                    XML(1) = -denl(1) * factor1 * VVent(2)
                    If XMVent(2) > 0 Then FRACM(2) = XML(1) / XMVent(2)
                    QU(1) = XMU(1) * cpu(1) * TU(1) * 1000
                    QL(1) = XML(1) * cpl(1) * TL(1) * 1000
                    If QVent(2) > 0 Then FRACQ(2) = QL(1) / QVent(2)
                    For NP = 1 To NProd
                        PU(NP, 1) = XMU(1) * conu(NP, 1)
                        PL(NP, 1) = XML(1) * conl(NP, 1)
                        If PVent(NP, 2) > 0 Then FRACP(NP, 2) = PL(NP, 1) / PVent(NP, 2)
                    Next NP
                Else 'room = 2
                    factor1 = (DELY(2) / DEL(2)) ^ APOWER
                    XML(2) = -denl(2) * (1 - factor1) * VVent(1)
                    XMU(2) = -denu(2) * factor1 * VVent(1)
                    If XMVent(1) > 0 Then FRACM(1) = XMU(2) / XMVent(1)
                    QL(2) = XML(2) * cpl(2) * TL(2) * 1000
                    QU(2) = XMU(2) * cpu(2) * TU(2) * 1000
                    If QVent(1) > 0 Then FRACQ(1) = QU(2) / QVent(1)
                    For NP = 1 To NProd
                        PL(NP, 2) = XML(2) * conl(NP, 2)
                        PU(NP, 2) = XMU(2) * conu(NP, 2)
                        If PVent(NP, 1) > 0 Then FRACP(NP, 1) = PU(NP, 2) / PVent(NP, 1)
                    Next NP
                End If
            End If
        Next room
    End Sub
	
	Public Sub CV_Flogo(ByVal dteps As Double, ByVal ylayer As Double, ByVal yup As Double, ByVal ylow As Double, ByVal Tvent As Double, ByVal TU As Double, ByVal TL As Double, ByRef FU As Double)
		
		'no upper layer
		If ylayer >= yup Then
			If Tvent > TL + dteps Then
				FU = 1
				Exit Sub
			Else
				FU = 0
				Exit Sub
			End If
		End If
		
		'no lower layer
		If ylayer <= ylow Then
			If Tvent >= TU + dteps Then
				FU = 1
				Exit Sub
			Else
				FU = 0
				Exit Sub
			End If
		End If
		
		'upper layer temp>lower layer temp
		If TU > TL Then
			'determine FU
			If Tvent >= TU Then
				FU = 1
				Exit Sub
			ElseIf Tvent <= TL Then 
				FU = 0
				Exit Sub
			Else
				FU = (Tvent - TL) / (TU - TL)
				Exit Sub
			End If
			
			'upper layer temp = lower temp
		ElseIf TU <= TL Then 
			If Tvent > TU Then
				FU = 1
				Exit Sub
			Else
				FU = 0
				Exit Sub
			End If
		End If
	End Sub
End Module