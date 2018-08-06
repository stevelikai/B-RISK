Option Strict Off
Option Explicit On
Imports System.Math

Module CCFM
    Public Sub CommonWall(ByRef yflor() As Double, ByRef yceil() As Double, ByRef ylay() As Double, ByRef denu() As Double, ByRef denl() As Double, ByRef pflor() As Double, ByRef nelev As Integer, ByRef yelev() As Double, ByRef dp1m2() As Double, ByRef pab1() As Double, ByRef pab2() As Double)
        '===============================================================
        '   Setup for calculation of the flow through vents in the wall
        '   segment common to two rooms. This routine calculates room
        '   pressures and room to room pressure differences at certain elevations
        '   along common walls of adjacent rooms. These elevations are:
        '
        '   ymin    maximum of two rooms floor elevations
        '   ymax    minimum of two rooms ceiling elevations
        '   layer elevations between ymin and ymax
        '   neutral planes elevations between ymin and ymax
        '
        '   inputs
        '   yflor()   elevation of floor in each room (m)
        '   yceil()   elevation of ceiling in each room (m)
        '   ylay()    elevation of layer in each room (m)
        '   denu()    upper layer density in each room (kg/m3)
        '   denl()    lower layer density in each room (kg/m3)
        '   pflor()   pressure at floor of each room (Pa)
        '   pdatum    absolute pressure at reference elevation (Pa)
        '
        '   outputs
        '   nelev     number of elevations
        '   yelev()   elevation (m)
        '   dp1m2()   pressure difference at elevations in yelev (Pa)
        '   pab1()    array of absolute pressures at elevations of interest (Pa)
        '   pab2()    array of absolute pressures at elevations of interest (Pa)
        '
        '   Reference NISTIR 4344
        '   6 June 1998
        '==============================================================

        Dim ymin, ymax As Double
        Dim p1, y1, y2, p2 As Double
        Dim n, i As Integer
        Dim yinterp, dummy As Double

        Try


            'find the lowest point of the common wall segment
            If yflor(1) > yflor(2) Then
                ymin = yflor(1)
            Else
                ymin = yflor(2)
            End If

            'find the highest point of the common wall segment
            If yceil(1) < yceil(2) Then
                ymax = yceil(1)
            Else
                ymax = yceil(2)
            End If

            yelev(1) = yflor(1)
            yelev(2) = yflor(2)
            yelev(3) = ylay(1)
            yelev(4) = ylay(2)
            yelev(5) = yceil(1)
            yelev(6) = yceil(2)

            nelev = 0
            For i = 1 To 6
                'keep only those elevations which fall between ymin and ymax
                If ymin <= yelev(i) And yelev(i) <= ymax Then
                    nelev = nelev + 1
                    yelev(nelev) = yelev(i)
                    Call DELP(yelev(nelev), yflor, ylay, denl, denu, pflor, dp1m2(nelev), pab1(nelev), pab2(nelev))
                End If
            Next

            'sort the list of heights and remove any duplicates that are found
            Call HSort(yelev, dp1m2, pab1, pab2, nelev)
            Call Remove_Duplicates(yelev, dp1m2, pab1, pab2, nelev)
            n = nelev
            For i = 2 To nelev
                'look for neutral planes
                If dp1m2(i) * dp1m2(i - 1) < 0 Then
                    'change of sign, so interpolate to find the height where pressure diff = 0
                    y2 = yelev(i)
                    y1 = yelev(i - 1)
                    p2 = dp1m2(i)
                    p1 = dp1m2(i - 1)
                    yinterp = (y2 * p1 - y1 * p2) / (p1 - p2)

                    'call routine to calculate absolute pressures at neutral plane
                    Call DELP(yinterp, yflor, ylay, denl, denu, pflor, dummy, pab1(n + 1), pab2(n + 1))
                    yelev(n + 1) = yinterp
                    dp1m2(n + 1) = 0
                    n = n + 1
                End If
            Next

            're-sort list and remove duplicates
            If nelev <> n Then
                nelev = n
                Call HSort(yelev, dp1m2, pab1, pab2, nelev)
                Call Remove_Duplicates(yelev, dp1m2, pab1, pab2, nelev)
            End If

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in CCFM.vb CommonWall")
        End Try
    End Sub
	
	Public Sub Remove_Duplicates(ByRef X() As Double, ByRef Y() As Double, ByRef Z() As Double, ByRef W() As Double, ByRef n As Integer)
		'=======================================================================
		'   This routine removes duplicate entries from the sorted data quadruples
		'   x(i), y(i), z(i), w(i), i=1 ... N
		'   The parameter N is then adjusted to the new number of entries.
		'   Two entries i and i+1 are duplicates if x(i)=x(i+1).
		'
		'   28 Jan 2008
		'======================================================================
		Dim i, j, flag As Integer
		Dim diff As Double
		
		j = 1
		flag = 0
		
		For i = 2 To n
			diff = X(i) - X(j)
			If diff > gcd_Machine_Error Then
				j = j + 1
				If flag = 1 Then
					X(j) = X(i)
					Y(j) = Y(i)
					Z(j) = Z(i)
					W(j) = W(i)
				End If
			Else : flag = 1
			End If
		Next 
		
		n = j
		
	End Sub
	
	Public Sub HSort(ByRef RA() As Double, ByRef RB() As Double, ByRef rc() As Double, ByRef RD() As Double, ByRef n As Integer)
		'=======================================================================
		'   This routine sorts the array RA in ascending order using a heap sort.
		'   Corresponding changes are made in the arrays RB, RC, and RD.
		'
		'   28 Jan 2008
		'======================================================================
		Dim L, IR As Integer
		Dim RRA, RRB As Double
		Dim RRC, RRD As Double
		Dim i, j As Integer
        Try

            L = n \ 2 + 1
		IR = n
		
		Do 
			If L > 1 Then
				L = L - 1
				RRA = RA(L)
				RRB = RB(L)
				RRC = rc(L)
				RRD = RD(L)
			Else
				RRA = RA(IR)
				RRB = RB(IR)
				RA(IR) = RA(1)
				RB(IR) = RB(1)
				RRC = rc(IR)
				RRD = RD(IR)
				rc(IR) = rc(1)
				RD(IR) = RD(1)
				IR = IR - 1
				If IR = 1 Then
					RA(1) = RRA
					RB(1) = RRB
					rc(1) = RRC
					RD(1) = RRD
					Exit Sub
				End If
			End If
			i = L
			j = L + L
			
			Do While j <= IR
				If j < IR Then
					If RA(j) < RA(j + 1) Then
						j = j + 1
					End If
				End If
				
				If RRA < RA(j) Then
					RA(i) = RA(j)
					RB(i) = RB(j)
					rc(i) = rc(j)
					RD(i) = RD(j)
					i = j
					j = j + j
				Else
					j = IR + 1
				End If
			Loop 
			RA(i) = RRA
			RB(i) = RRB
			rc(i) = RRC
			RD(i) = RRD
		Loop

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in CCFM.vb Hsort")

        End Try
    End Sub
	
	
    Public Sub DELP(ByRef Y As Double, ByRef yflor() As Double, ByRef ylay() As Double, ByRef denl() As Double, ByRef denu() As Double, ByRef pflor() As Double, ByRef DP As Double, ByRef pab1 As Double, ByRef pab2 As Double)
        '=======================================================================
        '   This routine calculates the hydrostatic pressures at a specified
        '   elevation in each of two adjacent rooms and the pressure difference.
        '   The basic calculation involves determination and differencing of
        '   hydrostatic pressures above a specified datum pressure.
        '
        '   Y   height above datum elevation where pressure difference is to be calculated
        '   yflor   height of the floor in each room above datum elevation
        '   ylay    height of layer in each room above datum elevation
        '   denu()    upper layer density in each room (kg/m3)
        '   denl()    lower layer density in each room (kg/m3)
        '   pflor()   pressure at floor of each room (Pa)
        '   pdatum    absolute pressure at reference elevation (Pa)
        '
        '   6 June 1998
        '======================================================================
        Dim i As Integer
        Dim dp2, dp1 As Double
        Dim PROOM(2) As Double
        Try


            DP = 0

            For i = 1 To 2
                PROOM(i) = 0
                If yflor(i) <= Y And Y <= ylay(i) Then
                    'the height y is in the lower layer
                    PROOM(i) = -(Y - yflor(i)) * denl(i) * G
                Else
                    If Y > ylay(i) Then
                        'the height is in the upper layer
                        PROOM(i) = -(ylay(i) - yflor(i)) * denl(i) * G - (Y - ylay(i)) * denu(i) * G
                    End If
                End If
            Next

            'change in pressure is the difference in pressure in two rooms
            dp1 = pflor(1) + PROOM(1)
            dp2 = pflor(2) + PROOM(2)
            pab1 = Atm_Pressure * 1000 + dp1
            pab2 = Atm_Pressure * 1000 + dp2

            'If (Abs(dp1) + Abs(dp2)) > 1 Then
            '    temp = (Abs(dp1) + Abs(dp2)) * 0.001
            'Else
            '    temp = 0.001
            'End If
            'If Abs(dp1 - dp2) < temp Then
            '    DP = 0
            'Else
            DP = dp1 - dp2
            If Abs(DP) < gcd_Machine_Error Then DP = 0
            'End If

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in CCFM.vb DELP")
        End Try
    End Sub
	
    Public Sub VENTHP(ByRef DischargeCoeff As Double, ByRef yflor() As Double, ByRef ylay() As Double, ByRef TU() As Double, ByRef TL() As Double, ByRef cpu() As Double, ByRef cpl() As Double, ByRef denl() As Double, ByRef denu() As Double, ByRef pflor() As Double, ByVal yvtop As Double, ByVal yvbot As Double, ByVal avent As Double, ByRef yelev() As Double, ByRef dp1m2() As Double, ByRef nelev As Integer, ByRef pab1() As Double, ByRef pab2() As Double, ByRef conl(,) As Double, ByRef conu(,) As Double, ByVal NProd As Integer, ByVal epsp As Double, ByRef cslab(,) As Double, ByRef pslab(,) As Double, ByRef tslab() As Double, ByRef qslab() As Double, ByRef yslab() As Double, ByRef dirs12() As Double, ByRef dpv1m2() As Double, ByRef yvelev() As Double, ByRef xmslab() As Double, ByRef nvelev As Integer, ByRef nslab As Integer)
        '=======================================================================
        '   This routine calculates the flow of mass, enthalpy, oxygen and other
        '   products through a vertical constant-width vent in a wall segment comon
        '   to two rooms. The subroutine uses input data describing the two layer
        '   environment in each of the two rooms and other input data calculated
        '   in subroutine Commonwall.
        '
        '   6 June 1998
        ' 1 feb 2008 Amanda
        '======================================================================
        Dim y1, ymax, ymin, fact1, y2 As Double
        Dim P1RT, p2, p1, RNum, P2RT As Double
        Dim WW, FF, ZZ, GG, CC As Double
        Dim n, i, jroom As Integer
        Dim ptest As Double
        Dim VP1RT(4) As Double
        Dim VP2RT(4) As Double
        Dim prat As Double
        Dim ymid, EPS As Double
        Dim rslab(6) As Double
        Dim cpslab(6) As Double
        Dim iprod As Integer
        Dim pabst, r1, pabsb As Double
        Dim psum, X0, Xx, epscut, wp As Double
        Dim pm As Double
        Dim pabv1(9) As Double
        Dim pabv2(9) As Double

        Try

            nvelev = 0
        If yvbot >= yelev(1) Then
            ymin = yvbot
        Else
            ymin = yelev(1)
        End If
        If yvtop < yelev(nelev) Then
            ymax = yvtop
        Else
            ymax = yelev(nelev)
        End If

        'put yvbot into yvelev if it lies between ymin and ymax and calculate the pressure at yvbot
        If ymin <= yvbot And yvbot <= ymax Then
            nvelev = nvelev + 1
            yvelev(nvelev) = yvbot
            Call DELP(yvbot, yflor, ylay, denl, denu, pflor, dpv1m2(nvelev), pabv1(nvelev), pabv2(nvelev))
        End If

        For i = 1 To nelev
            If ymin <= yelev(i) And yelev(i) <= ymax Then
                nvelev = nvelev + 1
                yvelev(nvelev) = yelev(i)
                dpv1m2(nvelev) = dp1m2(i)
                pabv1(nvelev) = pab1(i)
                pabv2(nvelev) = pab2(i)
            End If
        Next

        'put yvtop into yvelev if it lies between ymin ymax and calculate the pressure at yvtop
        If ymin <= yvtop And yvtop <= ymax Then
            nvelev = nvelev + 1
            yvelev(nvelev) = yvtop
            Call DELP(yvtop, yflor, ylay, denl, denu, pflor, dpv1m2(nvelev), pabv1(nvelev), pabv2(nvelev))
        End If

        'sort the list and remove duplicates
        Call HSort(yvelev, dpv1m2, pabv1, pabv2, nvelev)
        Call Remove_Duplicates(yvelev, dpv1m2, pabv1, pabv2, nvelev)

        'the number of slabs is 1 less than the number of heights
        nslab = nvelev - 1

        For n = 1 To nslab
            'determine whether temperature and density properties should come from room 1 or room 2
            ptest = dpv1m2(n + 1) + dpv1m2(n)
            If ptest > gcd_Machine_Error ^ 2 Then 'average pressure of room 1 > average pressure of room 2
                'If ptest > 0 Then
                jroom = 1
                dirs12(n) = 1
                prat = pabv2(n) / pabv1(n)
            ElseIf ptest < -gcd_Machine_Error ^ 2 Then  'average pressure of room 2 > average pressure of room 1
                'ElseIf ptest < 0 Then 'average pressure of room 2 > average pressure of room 1
                dirs12(n) = -1
                jroom = 2
                prat = pabv1(n) / pabv2(n)
            Else 'avg P of Room 1 ~ avg P of Room 2, i.e. no pressure difference between rooms, therefore no pressure driven flow
                dirs12(n) = 0
                jroom = 1
                prat = 0 'as per original ccfm code
                'prat = 1 'cw 17/3/2010
            End If

            'determine whether temperature and density properties should come from upper or lower layer
            ymid = (yvelev(n) + yvelev(n + 1)) / 2
            If ymid <= ylay(jroom) Then 'mid point of slab is in lower layer
                tslab(n) = TL(jroom)
                rslab(n) = denl(jroom) * prat
                cpslab(n) = cpl(jroom)
                For iprod = 1 To NProd
                    cslab(n, iprod) = conl(iprod, jroom)
                Next
            Else 'mid point of slab is in upper layer
                tslab(n) = TU(jroom)
                rslab(n) = denu(jroom) * prat
                cpslab(n) = cpu(jroom)
                For iprod = 1 To NProd
                    cslab(n, iprod) = conu(iprod, jroom)
                Next
            End If

            'initialisation
            xmslab(n) = 0
            qslab(n) = 0
            For iprod = 1 To NProd
                pslab(n, iprod) = 0
            Next
            yslab(n) = ymid
            p1 = Abs(dpv1m2(n)) 'bottom
            p2 = Abs(dpv1m2(n + 1)) 'top
            P1RT = Sqrt(p1) 'bottom
            P2RT = Sqrt(p2) 'top

            'for non-zero flow slabs determine xmslab(n) and yslab(n)
            'if both cross pressures are 0 then there is no flow
            If p1 > 0 Or p2 > 0 Then 'flow exists
                r1 = rslab(n)
                y2 = yvelev(n + 1)
                y1 = yvelev(n)
                pabst = pabv1(n + 1)
                If pabv2(n + 1) > pabst Then
                    pabst = pabv2(n + 1)
                End If

                pabsb = pabv1(n)
                If pabv2(n) > pabsb Then
                    pabsb = pabv2(n)
                End If

                EPS = (p1 + p2) / (pabst + pabsb)
                Xx = 1 - EPS

                'discharge coefficient
                CC = DischargeCoeff + 0.17 * EPS
                If CC > 1 Then
                    CC = 1.0#
                End If

                X0 = 1

                If Abs(pflor(1)) >= Abs(pflor(2)) Then
                    pm = Abs(pflor(1))
                Else
                    pm = Abs(pflor(2))
                End If

                If X0 >= pm Then
                    epscut = epsp * X0
                Else
                    epscut = epsp * pm
                End If

                epscut = Sqrt(epscut)

                psum = P1RT + P2RT
                If Abs(psum) < gcd_Machine_Error Then
                    'If psum = 0 Then
                    wp = 0
                Else
                    wp = (p1 + (P1RT * P2RT) + p2) / psum
                End If
                'numerical damping factor
                If Abs(wp / epscut) <= 130 Then
                    ZZ = 1 - Exp(-wp / epscut)
                Else
                    ZZ = 1
                End If

                If EPS <= 0.00001 Then
                    WW = 1 - (3 / (4 * gamma)) * EPS
                Else
                    If EPS < (1 - gamcut) Then
                        GG = Xx ^ (1 / gamma)
                        FF = Sqrt((2 * gamma / (gamma - 1)) * GG * GG * (1 - Xx / GG))
                    Else
                        FF = gammax
                    End If
                    WW = FF / Sqrt(EPS + EPS)
                End If
                If r1 < 0 Then
                    Exit Sub
                End If
                xmslab(n) = CC * WW * Sqrt(8 * r1) * avent * (y2 - y1) / (ymax - ymin) * (p2 + P1RT * P2RT + p1) / (P2RT + P1RT)
                xmslab(n) = xmslab(n) * ZZ
                xmslab(n) = xmslab(n) / 3
                qslab(n) = cpslab(n) * xmslab(n) * tslab(n) * 1000 'Watts

                For iprod = 1 To NProd
                    pslab(n, iprod) = cslab(n, iprod) * xmslab(n)
                Next

                'if pressure are close together then set slab height at midpoint
                If Abs(p1 - p2) > 0.001 Then
                    VP1RT(1) = P1RT
                    VP2RT(1) = P2RT
                    For i = 2 To 4
                        VP1RT(i) = VP1RT(i - 1) * P1RT
                        VP2RT(i) = VP2RT(i - 1) * P2RT
                    Next
                    RNum = VP2RT(4) + VP1RT(4)
                    For i = 1 To 3
                        RNum = RNum + VP1RT(i) * VP2RT(4 - i)
                    Next
                    fact1 = RNum / (p1 + P1RT * P2RT + p2)
                    yslab(n) = (0.6 * fact1 * (y2 - y1) + (p2 * y1 - p1 * y2)) / (p2 - p1)
                    'Else yslab(n) = ymid, as set in initialisation
                End If
                'Else no flow exists
            End If
        Next
            System.Windows.Forms.Application.DoEvents()


        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in CCFM.vb VENTHP")
        End Try
    End Sub
	
    Public Sub FLogo2(ByRef uvol() As Double, ByRef WAfloor_flag As Short, ByRef dirs12() As Double, ByVal NProd As Integer, ByRef yslab1() As Double, ByRef yslab2() As Double, ByRef xmslab() As Double, ByRef tslab() As Double, ByRef nslab As Integer, ByRef TU() As Double, ByRef TL() As Double, ByRef yflor() As Double, ByRef yceil() As Double, ByRef ylay() As Double, ByRef qslab() As Double, ByRef rlam As Double, ByRef pslab(,) As Double, ByVal mxprd As Integer, ByVal mxslab As Integer, ByVal dteps As Double, ByRef QSLVNT() As Double, ByRef QSUVNT() As Double, ByRef UFLW2(,,) As Double)
        '==========================adouble=============================================
        '   This routine calculates the deposition of mass, enthalpy, oxygen and other
        '   products passing between two rooms through a vertical constant-width vent
        '
        '    28 January 2008
        '======================================================================
        'define indexes for use with product arrays
        Dim P, Q, U, L, m, O, IP As Short
        Dim ito, i, iprod, ifrom, ilay As Integer
        Dim n As Short
        Dim TTU, yup, ylayer, ylow, TTS, TTL As Double
        Dim FU, yslabf, yslabt, FL, TTR As Double
        Dim FF() As Double
        Dim xmterm, FQ, FS, Qterm As Double
        Dim no_flow_flag As Short

        ReDim FF(2)

        L = 1
        U = 2
        m = 1
        Q = 2
        O = 3
        P = 3

        'Initialise outputs
        For i = 1 To 2
            For iprod = 1 To NProd + 2
                UFLW2(i, iprod, L) = 0
                UFLW2(i, iprod, U) = 0
            Next
            QSUVNT(i) = 0
            QSLVNT(i) = 0
        Next

        WAfloor_flag = 0
        'Put each slab flow into appropriate layer of room ito and take slab flow
        'out of appropriate layer of room ifrom
        For n = 1 To nslab 'determine what room flow is coming from
            no_flow_flag = 0
            If dirs12(n) > gcd_Machine_Error Then 'i.e. dirs12(n) = 1
                ifrom = 1
                ito = 2
                If n = 1 Then WAfloor_flag = 1 'determine flow direction at floor level (-1,0,1)
            ElseIf dirs12(n) < -gcd_Machine_Error Then  'i.e. dirs12(n) = -1
                ifrom = 2
                ito = 1
                If n = 1 Then WAfloor_flag = -1 'determine flow direction at floor level (-1,0,1)
            Else 'If dirs12(n) = 0 Then 'no flow
                no_flow_flag = 1
            End If

            If no_flow_flag = 0 Then
                'apportion flow between layers of destination room
                ylayer = ylay(ito)
                ylow = yflor(ito)
                yup = yceil(ito)
                TTS = tslab(n)
                TTU = TU(ito)
                TTL = TL(ito)
                If ito = 1 Then
                    yslabt = yslab1(n)
                    yslabf = yslab2(n)
                Else
                    yslabt = yslab2(n)
                    yslabf = yslab1(n)
                End If

                If uvol(ito) <= Smallvalue + gcd_Machine_Error Then 'no upper layer - use criteria consistent with diff_eqns i.e min vol = smallvalue

                    If TTS > TTL + dteps Then 'slab temp must be at least 3K above lower layer temp to start formation of a layer in the next room
                        FU = 1
                        TTR = TTL
                    Else
                        FU = 0
                        TTR = TTL
                    End If
                ElseIf ylayer <= ylow Then  'no lower layer
                    'If TTS >= TTU + dteps Then ' <= 2008.1
                    If TTS >= TTU - dteps Then 'this is what CCFM says, if slab temp is within dteps of the upper layer temp, then do not allow the lower layer to grow
                        FU = 1
                        TTR = TTU
                    Else
                        FU = 0
                        TTR = TTU
                    End If
                ElseIf TTU > TTL Then  'upper layer temp>lower layer temp
                    'determine FU
                    If TTS >= TTU Then
                        FU = 1
                    ElseIf TTS <= TTL Then
                        FU = 0
                    Else
                        FU = (TTS - TTL) / (TTU - TTL)
                    End If

                    'determine ttr
                    If yslabt > ylayer Then
                        TTR = TTU
                    ElseIf yslabt < ylayer Then
                        TTR = TTL
                    Else
                        TTR = (TTU + TTL) / 2
                    End If
                ElseIf TTU <= TTL Then  'upper layer temp = lower temp
                    If TTS > TTU Then
                        FU = 1
                        TTR = TTU
                    ElseIf TTS < TTU Then
                        FU = 0
                        TTR = TTL
                    Else
                        If yslabt > ylayer Then
                            FU = 1
                            TTR = TTU
                        ElseIf yslabt < ylayer Then
                            FU = 0
                            TTR = TTL
                        Else
                            FU = 0.5
                            TTR = TTL
                        End If
                    End If
                End If

                FL = 1 - FU
                FF(L) = FL
                FF(U) = FU
                FQ = rlam * (1 - TTR / TTS) '=0
                FS = 1 - FQ '=1
                xmterm = xmslab(n)
                Qterm = qslab(n)
                QSUVNT(ito) = QSUVNT(ito) + FU * Qterm * FQ
                QSLVNT(ito) = QSLVNT(ito) + FL * Qterm * FQ

                For ilay = 1 To 2
                    UFLW2(ito, m, ilay) = UFLW2(ito, m, ilay) + FF(ilay) * xmterm
                    UFLW2(ito, Q, ilay) = UFLW2(ito, Q, ilay) + FF(ilay) * Qterm * FS
                    For iprod = 1 To NProd
                        IP = P + iprod - 1
                        UFLW2(ito, IP, ilay) = UFLW2(ito, IP, ilay) + FF(ilay) * pslab(n, iprod)
                    Next iprod
                Next ilay

                'flow entirely from upper layer of room ifrom
                If yslabf >= ylay(ifrom) Then
                    UFLW2(ifrom, m, U) = UFLW2(ifrom, m, U) - xmterm
                    UFLW2(ifrom, Q, U) = UFLW2(ifrom, Q, U) - Qterm
                    For iprod = 1 To NProd
                        IP = P + iprod - 1
                        UFLW2(ifrom, IP, U) = UFLW2(ifrom, IP, U) - pslab(n, iprod)
                    Next
                Else
                    'flow entirely from lower layer of room ifrom
                    UFLW2(ifrom, m, L) = UFLW2(ifrom, m, L) - xmterm
                    UFLW2(ifrom, Q, L) = UFLW2(ifrom, Q, L) - Qterm
                    For iprod = 1 To NProd
                        IP = P + iprod - 1
                        UFLW2(ifrom, IP, L) = UFLW2(ifrom, IP, L) - pslab(n, iprod)
                    Next
                End If
            End If
        Next
    End Sub
    Public Sub FLogo1(ByVal room1 As Integer, ByVal room2 As Integer, ByRef uvol() As Double, ByRef WAfloor_flag As Short, ByRef dirs12() As Double, ByVal NProd As Integer, ByRef yslab1() As Double, ByRef yslab2() As Double, ByRef xmslab() As Double, ByRef tslab() As Double, ByRef nslab As Integer, ByRef TU() As Double, ByRef TL() As Double, ByRef yflor() As Double, ByRef yceil() As Double, ByRef ylay() As Double, ByRef qslab() As Double, ByRef rlam As Double, ByRef pslab(,) As Double, ByVal mxprd As Integer, ByVal mxslab As Integer, ByVal dteps As Double, ByRef QSLVNT() As Double, ByRef QSUVNT() As Double, ByRef UFLW2(,,) As Double)
        '==========================adouble=============================================
        '   This routine calculates the deposition of mass, enthalpy, oxygen and other
        '   products passing between two rooms through a vertical constant-width vent
        '
        '    11 April 2008
        ' the rules in this routine are based on CFAST 3.17 source code
        ' upper layer to upper layer, and lower layer to lower layer
        '======================================================================
        'define indexes for use with product arrays
        Dim P, Q, U, L, m, O, IP As Short
        Dim ito, i, iprod, ifrom, ilay As Integer
        Dim n As Short
        Dim TTU, yup, ylayer, ylow, TTS, TTL As Double
        Dim FU, yslabf, yslabt, FL, TTR As Double
        Dim FF() As Double
        Dim xmterm, FQ, FS, Qterm As Double
        Dim no_flow_flag As Short

        ReDim FF(2)

        L = 1
        U = 2
        m = 1
        Q = 2
        O = 3
        P = 3

        'Initialise outputs
        For i = 1 To 2
            For iprod = 1 To NProd + 2
                UFLW2(i, iprod, L) = 0
                UFLW2(i, iprod, U) = 0
            Next
            QSUVNT(i) = 0
            QSLVNT(i) = 0
        Next

        WAfloor_flag = 0
        'Put each slab flow into appropriate layer of room ito and take slab flow
        'out of appropriate layer of room ifrom
        For n = 1 To nslab 'determine what room flow is coming from
            no_flow_flag = 0
            If dirs12(n) > gcd_Machine_Error Then 'i.e. dirs12(n) = 1
                ifrom = 1
                ito = 2
                If n = 1 Then WAfloor_flag = 1 'determine flow direction at floor level (-1,0,1)
            ElseIf dirs12(n) < -gcd_Machine_Error Then  'i.e. dirs12(n) = -1
                ifrom = 2
                ito = 1
                If n = 1 Then WAfloor_flag = -1 'determine flow direction at floor level (-1,0,1)
            Else 'If dirs12(n) = 0 Then 'no flow
                no_flow_flag = 1
            End If

            If no_flow_flag = 0 Then
                'apportion flow between layers of destination room
                ylayer = ylay(ito)
                ylow = yflor(ito)
                yup = yceil(ito)
                TTS = tslab(n)
                TTU = TU(ito)
                TTL = TL(ito)
                If ito = 1 Then
                    yslabt = yslab1(n)
                    yslabf = yslab2(n)
                Else
                    yslabt = yslab2(n)
                    yslabf = yslab1(n)
                End If

                '=========================
                'from CFAST 3.17
                If yslabf > ylay(ifrom) Then 'upper layer gases
                    FU = 1
                Else 'lower layer gases
                    FU = 0
                End If


                ''0502016 inflow from outside always goes into the lower layer in a choked flow situation
                'If yslabf > ylay(ito) Then 'upper layer gases '0502016
                '    FU = 1.0
                'Else 'lower layer gases
                '    FU = 0
                'End If


                If ito = 1 Then
                    If room1 <= NumberRooms Then
                        If TwoZones(room1) = False Then FU = 1 'if destination is single zone must put it in the upper layer
                    End If
                Else
                    If room2 <= NumberRooms Then
                        If TwoZones(room2) = False Then FU = 1 'if destination is single zone must put it in the upper layer
                    End If
                End If
                '=======================================

                FL = 1 - FU
                FF(L) = FL
                FF(U) = FU
                FQ = rlam * (1 - TTR / TTS) '=0 : also TTR is therrefore not needed/used, since we do not use rlam (heat loss fraction to surfaces), we calculate that separately
                FS = 1 - FQ '=1
                xmterm = xmslab(n)
                Qterm = qslab(n)
                QSUVNT(ito) = QSUVNT(ito) + FU * Qterm * FQ
                QSLVNT(ito) = QSLVNT(ito) + FL * Qterm * FQ

                For ilay = 1 To 2
                    UFLW2(ito, m, ilay) = UFLW2(ito, m, ilay) + FF(ilay) * xmterm
                    UFLW2(ito, Q, ilay) = UFLW2(ito, Q, ilay) + FF(ilay) * Qterm * FS
                    For iprod = 1 To NProd
                        IP = P + iprod - 1
                        UFLW2(ito, IP, ilay) = UFLW2(ito, IP, ilay) + FF(ilay) * pslab(n, iprod)
                    Next iprod
                Next ilay

                'flow entirely from upper layer of room ifrom
                If yslabf >= ylay(ifrom) Then
                    UFLW2(ifrom, m, U) = UFLW2(ifrom, m, U) - xmterm
                    UFLW2(ifrom, Q, U) = UFLW2(ifrom, Q, U) - Qterm
                    For iprod = 1 To NProd
                        IP = P + iprod - 1
                        UFLW2(ifrom, IP, U) = UFLW2(ifrom, IP, U) - pslab(n, iprod)
                    Next
                Else
                    'flow entirely from lower layer of room ifrom
                    UFLW2(ifrom, m, L) = UFLW2(ifrom, m, L) - xmterm
                    UFLW2(ifrom, Q, L) = UFLW2(ifrom, Q, L) - Qterm
                    For iprod = 1 To NProd
                        IP = P + iprod - 1
                        UFLW2(ifrom, IP, L) = UFLW2(ifrom, IP, L) - pslab(n, iprod)
                    Next
                End If
            End If
        Next
    End Sub
    Public Sub FLogo1A(ByRef uvol() As Double, ByRef WAfloor_flag As Short, ByRef dirs12() As Double, ByVal NProd As Integer, ByRef yslab1() As Double, ByRef yslab2() As Double, ByRef xmslab() As Double, ByRef tslab() As Double, ByRef nslab As Integer, ByRef TU() As Double, ByRef TL() As Double, ByRef yflor() As Double, ByRef yceil() As Double, ByRef ylay() As Double, ByRef qslab() As Double, ByRef rlam As Double, ByRef pslab(,) As Double, ByVal mxprd As Integer, ByVal mxslab As Integer, ByVal dteps As Double, ByRef QSLVNT() As Double, ByRef QSUVNT() As Double, ByRef UFLW2(,,) As Double)
        '==========================adouble=============================================
        '   This routine calculates the deposition of mass, enthalpy, oxygen and other
        '   products passing between two rooms through a vertical constant-width vent
        '
        '    11 April 2008
        ' the rules in this routine are based on CFAST 3.17 source code
        ' upper layer to upper layer, and lower layer to lower layer
        '======================================================================
        'define indexes for use with product arrays
        Dim P, Q, U, L, m, O, IP As Short
        Dim ito, i, iprod, ifrom, ilay As Integer
        Dim n As Short
        Dim TTU, yup, ylayer, ylow, TTS, TTL As Double
        Dim FU, yslabf, yslabt, FL, TTR As Double
        Dim FF() As Double
        Dim xmterm, FQ, FS, Qterm As Double
        Dim no_flow_flag As Short

        ReDim FF(2)

        L = 1
        U = 2
        m = 1
        Q = 2
        O = 3
        P = 3

        'Initialise outputs
        For i = 1 To 2
            For iprod = 1 To NProd + 2
                UFLW2(i, iprod, L) = 0
                UFLW2(i, iprod, U) = 0
            Next
            QSUVNT(i) = 0
            QSLVNT(i) = 0
        Next

        WAfloor_flag = 0
        'Put each slab flow into appropriate layer of room ito and take slab flow
        'out of appropriate layer of room ifrom
        For n = 1 To nslab 'determine what room flow is coming from
            no_flow_flag = 0
            If dirs12(n) > gcd_Machine_Error Then 'i.e. dirs12(n) = 1
                ifrom = 1
                ito = 2
                If n = 1 Then WAfloor_flag = 1 'determine flow direction at floor level (-1,0,1)
            ElseIf dirs12(n) < -gcd_Machine_Error Then  'i.e. dirs12(n) = -1
                ifrom = 2
                ito = 1
                If n = 1 Then WAfloor_flag = -1 'determine flow direction at floor level (-1,0,1)
            Else 'If dirs12(n) = 0 Then 'no flow
                no_flow_flag = 1
            End If

            If no_flow_flag = 0 Then
                'apportion flow between layers of destination room
                ylayer = ylay(ito)
                ylow = yflor(ito)
                yup = yceil(ito)
                TTS = tslab(n)
                TTU = TU(ito)
                TTL = TL(ito)
                If ito = 1 Then
                    yslabt = yslab1(n)
                    yslabf = yslab2(n)
                Else
                    yslabt = yslab2(n)
                    yslabf = yslab1(n)
                End If

                '=========================
                'from CFAST 3.17
                If yslabf > ylay(ifrom) Then 'upper layer gases
                    FU = 1
                Else 'lower layer gases
                    If ito = 1 Then
                        FU = 1 'adding flow to single zone room, add to upper layer
                    Else
                        FU = 0
                    End If

                End If

                '=======================================


                FL = 1 - FU
                FF(L) = FL
                FF(U) = FU
                FQ = rlam * (1 - TTR / TTS) '=0 : also TTR is therrefore not needed/used, since we do not use rlam (heat loss fraction to surfaces), we calculate that separately
                FS = 1 - FQ '=1
                xmterm = xmslab(n)
                Qterm = qslab(n)
                QSUVNT(ito) = QSUVNT(ito) + FU * Qterm * FQ
                QSLVNT(ito) = QSLVNT(ito) + FL * Qterm * FQ

                For ilay = 1 To 2
                    UFLW2(ito, m, ilay) = UFLW2(ito, m, ilay) + FF(ilay) * xmterm
                    UFLW2(ito, Q, ilay) = UFLW2(ito, Q, ilay) + FF(ilay) * Qterm * FS
                    For iprod = 1 To NProd
                        IP = P + iprod - 1
                        UFLW2(ito, IP, ilay) = UFLW2(ito, IP, ilay) + FF(ilay) * pslab(n, iprod)
                    Next iprod
                Next ilay

                'flow entirely from upper layer of room ifrom
                If yslabf >= ylay(ifrom) Then
                    UFLW2(ifrom, m, U) = UFLW2(ifrom, m, U) - xmterm
                    UFLW2(ifrom, Q, U) = UFLW2(ifrom, Q, U) - Qterm
                    For iprod = 1 To NProd
                        IP = P + iprod - 1
                        UFLW2(ifrom, IP, U) = UFLW2(ifrom, IP, U) - pslab(n, iprod)
                    Next
                Else
                    'flow entirely from lower layer of room ifrom
                    UFLW2(ifrom, m, L) = UFLW2(ifrom, m, L) - xmterm
                    UFLW2(ifrom, Q, L) = UFLW2(ifrom, Q, L) - Qterm
                    For iprod = 1 To NProd
                        IP = P + iprod - 1
                        UFLW2(ifrom, IP, L) = UFLW2(ifrom, IP, L) - pslab(n, iprod)
                    Next
                End If
            End If
        Next
    End Sub
    Public Sub FLogo1B(ByRef uvol() As Double, ByRef WAfloor_flag As Short, ByRef dirs12() As Double, ByVal NProd As Integer, ByRef yslab1() As Double, ByRef yslab2() As Double, ByRef xmslab() As Double, ByRef tslab() As Double, ByRef nslab As Integer, ByRef TU() As Double, ByRef TL() As Double, ByRef yflor() As Double, ByRef yceil() As Double, ByRef ylay() As Double, ByRef qslab() As Double, ByRef rlam As Double, ByRef pslab(,) As Double, ByVal mxprd As Integer, ByVal mxslab As Integer, ByVal dteps As Double, ByRef QSLVNT() As Double, ByRef QSUVNT() As Double, ByRef UFLW2(,,) As Double)
        '==========================adouble=============================================
        '   This routine calculates the deposition of mass, enthalpy, oxygen and other
        '   products passing between two rooms through a vertical constant-width vent
        '
        '    11 April 2008
        ' the rules in this routine are based on CFAST 3.17 source code
        ' upper layer to upper layer, and lower layer to lower layer
        '======================================================================
        'define indexes for use with product arrays
        Dim P, Q, U, L, m, O, IP As Short
        Dim ito, i, iprod, ifrom, ilay As Integer
        Dim n As Short
        Dim TTU, yup, ylayer, ylow, TTS, TTL As Double
        Dim FU, yslabf, yslabt, FL, TTR As Double
        Dim FF() As Double
        Dim xmterm, FQ, FS, Qterm As Double
        Dim no_flow_flag As Short

        ReDim FF(2)

        L = 1
        U = 2
        m = 1
        Q = 2
        O = 3
        P = 3

        'Initialise outputs
        For i = 1 To 2
            For iprod = 1 To NProd + 2
                UFLW2(i, iprod, L) = 0
                UFLW2(i, iprod, U) = 0
            Next
            QSUVNT(i) = 0
            QSLVNT(i) = 0
        Next

        WAfloor_flag = 0
        'Put each slab flow into appropriate layer of room ito and take slab flow
        'out of appropriate layer of room ifrom
        For n = 1 To nslab 'determine what room flow is coming from
            no_flow_flag = 0
            If dirs12(n) > gcd_Machine_Error Then 'i.e. dirs12(n) = 1
                ifrom = 1
                ito = 2
                If n = 1 Then WAfloor_flag = 1 'determine flow direction at floor level (-1,0,1)
            ElseIf dirs12(n) < -gcd_Machine_Error Then  'i.e. dirs12(n) = -1
                ifrom = 2
                ito = 1
                If n = 1 Then WAfloor_flag = -1 'determine flow direction at floor level (-1,0,1)
            Else 'If dirs12(n) = 0 Then 'no flow
                no_flow_flag = 1
            End If

            If no_flow_flag = 0 Then
                'apportion flow between layers of destination room
                ylayer = ylay(ito)
                ylow = yflor(ito)
                yup = yceil(ito)
                TTS = tslab(n)
                TTU = TU(ito)
                TTL = TL(ito)
                If ito = 1 Then
                    yslabt = yslab1(n)
                    yslabf = yslab2(n)
                Else
                    yslabt = yslab2(n)
                    yslabf = yslab1(n)
                End If

                '=========================
                'from CFAST 3.17
                If yslabf > ylay(ifrom) Then 'upper layer gases
                    FU = 1
                Else 'lower layer gases
                    If ito = 2 Then
                        FU = 1 'adding flow to single zone room, add to upper layer
                    Else
                        FU = 0
                    End If

                End If

                '=======================================


                FL = 1 - FU
                FF(L) = FL
                FF(U) = FU
                FQ = rlam * (1 - TTR / TTS) '=0 : also TTR is therrefore not needed/used, since we do not use rlam (heat loss fraction to surfaces), we calculate that separately
                FS = 1 - FQ '=1
                xmterm = xmslab(n)
                Qterm = qslab(n)
                QSUVNT(ito) = QSUVNT(ito) + FU * Qterm * FQ
                QSLVNT(ito) = QSLVNT(ito) + FL * Qterm * FQ

                For ilay = 1 To 2
                    UFLW2(ito, m, ilay) = UFLW2(ito, m, ilay) + FF(ilay) * xmterm
                    UFLW2(ito, Q, ilay) = UFLW2(ito, Q, ilay) + FF(ilay) * Qterm * FS
                    For iprod = 1 To NProd
                        IP = P + iprod - 1
                        UFLW2(ito, IP, ilay) = UFLW2(ito, IP, ilay) + FF(ilay) * pslab(n, iprod)
                    Next iprod
                Next ilay

                'flow entirely from upper layer of room ifrom
                If yslabf >= ylay(ifrom) Then
                    UFLW2(ifrom, m, U) = UFLW2(ifrom, m, U) - xmterm
                    UFLW2(ifrom, Q, U) = UFLW2(ifrom, Q, U) - Qterm
                    For iprod = 1 To NProd
                        IP = P + iprod - 1
                        UFLW2(ifrom, IP, U) = UFLW2(ifrom, IP, U) - pslab(n, iprod)
                    Next
                Else
                    'flow entirely from lower layer of room ifrom
                    UFLW2(ifrom, m, L) = UFLW2(ifrom, m, L) - xmterm
                    UFLW2(ifrom, Q, L) = UFLW2(ifrom, Q, L) - Qterm
                    For iprod = 1 To NProd
                        IP = P + iprod - 1
                        UFLW2(ifrom, IP, L) = UFLW2(ifrom, IP, L) - pslab(n, iprod)
                    Next
                End If
            End If
        Next
    End Sub
    Public Sub FLogo1C(ByRef uvol() As Double, ByRef WAfloor_flag As Short, ByRef dirs12() As Double, ByVal NProd As Integer, ByRef yslab1() As Double, ByRef yslab2() As Double, ByRef xmslab() As Double, ByRef tslab() As Double, ByRef nslab As Integer, ByRef TU() As Double, ByRef TL() As Double, ByRef yflor() As Double, ByRef yceil() As Double, ByRef ylay() As Double, ByRef qslab() As Double, ByRef rlam As Double, ByRef pslab(,) As Double, ByVal mxprd As Integer, ByVal mxslab As Integer, ByVal dteps As Double, ByRef QSLVNT() As Double, ByRef QSUVNT() As Double, ByRef UFLW2(,,) As Double)
        '==========================adouble=============================================
        '   This routine calculates the deposition of mass, enthalpy, oxygen and other
        '   products passing between two rooms through a vertical constant-width vent
        '
        '    11 April 2008
        ' the rules in this routine are based on CFAST 3.17 source code
        ' upper layer to upper layer, and lower layer to lower layer
        '======================================================================
        'define indexes for use with product arrays
        Dim P, Q, U, L, m, O, IP As Short
        Dim ito, i, iprod, ifrom, ilay As Integer
        Dim n As Short
        Dim TTU, yup, ylayer, ylow, TTS, TTL As Double
        Dim FU, yslabf, yslabt, FL, TTR As Double
        Dim FF() As Double
        Dim xmterm, FQ, FS, Qterm As Double
        Dim no_flow_flag As Short

        ReDim FF(2)

        L = 1
        U = 2
        m = 1
        Q = 2
        O = 3
        P = 3

        'Initialise outputs
        For i = 1 To 2
            For iprod = 1 To NProd + 2
                UFLW2(i, iprod, L) = 0
                UFLW2(i, iprod, U) = 0
            Next
            QSUVNT(i) = 0
            QSLVNT(i) = 0
        Next

        WAfloor_flag = 0
        'Put each slab flow into appropriate layer of room ito and take slab flow
        'out of appropriate layer of room ifrom
        For n = 1 To nslab 'determine what room flow is coming from
            no_flow_flag = 0
            If dirs12(n) > gcd_Machine_Error Then 'i.e. dirs12(n) = 1
                ifrom = 1
                ito = 2
                If n = 1 Then WAfloor_flag = 1 'determine flow direction at floor level (-1,0,1)
            ElseIf dirs12(n) < -gcd_Machine_Error Then  'i.e. dirs12(n) = -1
                ifrom = 2
                ito = 1
                If n = 1 Then WAfloor_flag = -1 'determine flow direction at floor level (-1,0,1)
            Else 'If dirs12(n) = 0 Then 'no flow
                no_flow_flag = 1
            End If

            If no_flow_flag = 0 Then
                'apportion flow between layers of destination room
                ylayer = ylay(ito)
                ylow = yflor(ito)
                yup = yceil(ito)
                TTS = tslab(n)
                TTU = TU(ito)
                TTL = TL(ito)
                If ito = 1 Then
                    yslabt = yslab1(n)
                    yslabf = yslab2(n)
                Else
                    yslabt = yslab2(n)
                    yslabf = yslab1(n)
                End If

                '=========================
                'from CFAST 3.17
                ' If yslabf > ylay(ifrom) Then 'upper layer gases
                FU = 1 'both rooms are single zone
                'Else 'lower layer gases
                '    If ito = 2 Then
                '        FU = 1 'adding flow to single zone room, add to upper layer
                '    Else
                '        FU = 0
                '    End If

                'End If

                '=======================================


                FL = 1 - FU
                FF(L) = FL
                FF(U) = FU
                FQ = rlam * (1 - TTR / TTS) '=0 : also TTR is therrefore not needed/used, since we do not use rlam (heat loss fraction to surfaces), we calculate that separately
                FS = 1 - FQ '=1
                xmterm = xmslab(n)
                Qterm = qslab(n)
                QSUVNT(ito) = QSUVNT(ito) + FU * Qterm * FQ
                QSLVNT(ito) = QSLVNT(ito) + FL * Qterm * FQ

                For ilay = 1 To 2
                    UFLW2(ito, m, ilay) = UFLW2(ito, m, ilay) + FF(ilay) * xmterm
                    UFLW2(ito, Q, ilay) = UFLW2(ito, Q, ilay) + FF(ilay) * Qterm * FS
                    For iprod = 1 To NProd
                        IP = P + iprod - 1
                        UFLW2(ito, IP, ilay) = UFLW2(ito, IP, ilay) + FF(ilay) * pslab(n, iprod)
                    Next iprod
                Next ilay

                'flow entirely from upper layer of room ifrom
                If yslabf >= ylay(ifrom) Then
                    UFLW2(ifrom, m, U) = UFLW2(ifrom, m, U) - xmterm
                    UFLW2(ifrom, Q, U) = UFLW2(ifrom, Q, U) - Qterm
                    For iprod = 1 To NProd
                        IP = P + iprod - 1
                        UFLW2(ifrom, IP, U) = UFLW2(ifrom, IP, U) - pslab(n, iprod)
                    Next
                Else
                    'flow entirely from lower layer of room ifrom
                    UFLW2(ifrom, m, L) = UFLW2(ifrom, m, L) - xmterm
                    UFLW2(ifrom, Q, L) = UFLW2(ifrom, Q, L) - Qterm
                    For iprod = 1 To NProd
                        IP = P + iprod - 1
                        UFLW2(ifrom, IP, L) = UFLW2(ifrom, IP, L) - pslab(n, iprod)
                    Next
                End If
            End If
        Next
    End Sub
End Module