Option Strict Off
Option Explicit On
Imports System.Math
Imports CenterSpace.NMath.Core
Imports CenterSpace.NMath.Analysis


Module KineticModelCode
    Dim Y_pyrol() As Double
    Dim DYDX_pyrol() As Double
    Sub Implicit_Temps_Ceil_kinetic(ByVal room As Integer, ByVal i As Integer, ByRef CeilingNode(,,) As Double)
        '*  ================================================================
        '*      This function updates the surface temperatures, using an
        '*      implicit finite difference method.
        '*  ================================================================

        Try
            Dim ceilinglayersremaining As Integer
            Dim k, j As Integer
            Dim ceilingnodestemp As Integer
            Dim NLC As Integer
            Dim ier As Short
            Dim ceilingnodeadjust As Integer
            Dim moisturecontent, temp As Double
            Dim prop_k As Double 'W/mK
            Dim char_alpha, char_c, char_fourier As Double
            Dim wood_alpha, wood_c, wood_fourier As Double
            Dim Coutbiot, WoodDensity As Double
            Dim kwood As Double

            kwood = CeilingConductivity(room)
            WoodDensity = CeilingDensity(room)

            Dim mf_init As Double = 0.1 'initial mc


            'chardensity = 0.63 * WoodDensity / (1 + moisturecontent)

            If CLTceilingpercent > 0 Then
                NLC = CeilingThickness(room) / 1000 / Lamella 'number of lamella - in two places also in main_program2

                ceilinglayersremaining = NLC - Lamella1 / Lamella + 1

                If ceilinglayersremaining = 0 Then
                    If ceilinglayersremaining = 0 Then Dim Message As String = CStr(tim(stepcount, 1)) & " sec. Ceiling layer at " & Format(Lamella1, "0.000") & " m delaminates. Simulation terminated. "

                    flagstop = 1 'do not continue
                    Exit Sub
                End If

                ceilingnodestemp = (Ceilingnodes - 1) * (ceilinglayersremaining) + 1 'remove one layer and recalc number of nodes
                ceilingnodeadjust = (Ceilingnodes - 1) * NLC + 1 - ceilingnodestemp

                'Find DeltaX
                CeilingDeltaX(room) = ceilinglayersremaining * Lamella / (ceilingnodestemp - 1)
            Else
                'clt model is on but not for ceiling
                NLC = 1
                ceilinglayersremaining = 1
                ceilingnodestemp = Ceilingnodes
                ceilingnodeadjust = 0
            End If

            Dim CeilingNodeTemp(ceilingnodestemp) As Double
            Dim CeilingNodeStatus(ceilingnodestemp) As Integer  'array to identify charred element

            For k = 1 To ceilingnodestemp
                CeilingNodeTemp(k) = CeilingNode(room, k + ceilingnodeadjust, i)

                'array to identify charred element
                If CeilingNodeTemp(k) > 300 + 273 Then CeilingNodeStatus(k) = 1

            Next

            Dim CeilingNodeUnExposed As Double = CeilingNodeTemp(ceilingnodestemp)
            Dim CeilingNodeExposed As Double = CeilingNodeTemp(1)


            If CeilingNodeTemp(ceilingnodestemp) <= 473 Then
                prop_k = kwood

            ElseIf CeilingNodeTemp(ceilingnodestemp) <= 663 Then
                prop_k = -0.617 + 0.0038 * CeilingNodeTemp(ceilingnodestemp) - 0.000004 * CeilingNodeTemp(ceilingnodestemp) ^ 2
            Else
                prop_k = 0.04429 + 0.0001477 * CeilingNodeTemp(ceilingnodestemp)
            End If

            'Find Biot Numbers -exterior side
            Coutbiot = OutsideConvCoeff * CeilingDeltaX(room) / prop_k

            Dim UC(ceilingnodestemp, ceilingnodestemp) As Double
            Dim CX(ceilingnodestemp, 1) As Double

            moisturecontent = CeilingElementMF(ceilingnodestemp - 1, 0, i) * mf_init 'mass fraction of the original wet wood
            WoodDensity = CeilingApparentDensity(ceilingnodestemp - 1, i)
            'chardensity = 0.63 * WoodDensity / (1 + moisturecontent)

            'Find Fourier Numbers -exterior side
            If CeilingNodeStatus(ceilingnodestemp) = 1 Then
                'char
                temp = CeilingNodeTemp(ceilingnodestemp) - 273 'node temp in deg C
                char_c = 714 + 2.3 * temp - 0.0008 * temp ^ 2 - 0.00000037 * temp ^ 3
                char_alpha = prop_k / (char_c * WoodDensity)
                char_fourier = char_alpha * Timestep / (CeilingDeltaX(room)) ^ 2

                'exterior boundary conditions
                CX(ceilingnodestemp, 1) = 2 * char_fourier * Coutbiot * ((ExteriorTemp - CeilingNodeUnExposed) - Surface_Emissivity(1, room) / OutsideConvCoeff * StefanBoltzmann * (CeilingNodeUnExposed ^ 4 - ExteriorTemp ^ 4)) + CeilingNodeUnExposed
                UC(ceilingnodestemp, ceilingnodestemp - 1) = -2 * char_fourier
                UC(ceilingnodestemp, ceilingnodestemp) = 1 + 2 * char_fourier
                UC(1, 1) = 1 + 2 * char_fourier
                UC(1, 2) = -2 * char_fourier
            Else
                'wood
                wood_c = 101.3 + 3.867 * CeilingNodeTemp(ceilingnodestemp)
                wood_c = (wood_c + 4187 * moisturecontent) / (1 + moisturecontent) + (23.55 * (CeilingNodeTemp(ceilingnodestemp) - 273) - 1326 * moisturecontent + 2417) * moisturecontent

                'correct for the latent heat of water over the temp range 90-110 C
                If CeilingNodeTemp(ceilingnodestemp) >= 273 + 80 And CeilingNodeTemp(ceilingnodestemp) <= 273 + 100 Then
                    wood_c = wood_c + mf_init * 2257 / 20 * 1000 'J/kgK
                End If

                wood_alpha = prop_k / (wood_c * WoodDensity)
                wood_fourier = wood_alpha * Timestep / (CeilingDeltaX(room)) ^ 2

                'exterior boundary conditions
                CX(ceilingnodestemp, 1) = 2 * wood_fourier * Coutbiot * ((ExteriorTemp - CeilingNodeUnExposed) - Surface_Emissivity(1, room) / OutsideConvCoeff * StefanBoltzmann * (CeilingNodeUnExposed ^ 4 - ExteriorTemp ^ 4)) + CeilingNodeUnExposed
                UC(ceilingnodestemp, ceilingnodestemp - 1) = -2 * wood_fourier
                UC(ceilingnodestemp, ceilingnodestemp) = 1 + 2 * wood_fourier
                UC(1, 1) = 1 + 2 * wood_fourier
                UC(1, 2) = -2 * wood_fourier
            End If

            'exposed side
            If CeilingNodeTemp(2) <= 473 Then
                prop_k = kwood
            ElseIf CeilingNodeTemp(2) <= 663 Then
                prop_k = -0.617 + 0.0038 * CeilingNodeTemp(2) - 0.000004 * CeilingNodeTemp(2) ^ 2
            Else
                prop_k = 0.04429 + 0.0001477 * CeilingNodeTemp(2)
            End If

            moisturecontent = CeilingElementMF(1, 0, i) * mf_init
            WoodDensity = CeilingApparentDensity(1, i)
            'chardensity = 0.63 * WoodDensity / (1 + moisturecontent)

            If CeilingNodeStatus(2) = 1 Then 'char
                temp = CeilingNodeTemp(2) - 273 'node temp in deg C
                char_c = 714 + 2.3 * temp - 0.0008 * temp ^ 2 - 0.00000037 * temp ^ 3
                char_alpha = prop_k / (char_c * WoodDensity)
                char_fourier = char_alpha * Timestep / (CeilingDeltaX(room)) ^ 2
                UC(1, 1) = 1 + 2 * char_fourier
                UC(1, 2) = -2 * char_fourier
                'interior boundary conditions
                CX(1, 1) = -2 * QCeiling(room, i) * 1000 * char_fourier * CeilingDeltaX(room) / prop_k + CeilingNodeExposed
                For k = 2 To ceilingnodestemp - 1
                    CX(k, 1) = CeilingNodeTemp(k)
                Next k
            Else
                'wood
                wood_c = 101.3 + 3.867 * CeilingNodeTemp(2)
                wood_c = (wood_c + 4187 * moisturecontent) / (1 + moisturecontent) + (23.55 * (CeilingNodeTemp(2) - 273) - 1326 * moisturecontent + 2417) * moisturecontent

                'correct for the latent heat of water over the temp range 90-110 C
                If CeilingNodeTemp(2) >= 273 + 80 And CeilingNodeTemp(2) <= 273 + 100 Then
                    wood_c = wood_c + mf_init * 2257 / 20 * 1000 'J/kgK
                End If

                wood_alpha = prop_k / (wood_c * WoodDensity)
                wood_fourier = wood_alpha * Timestep / (CeilingDeltaX(room)) ^ 2
                UC(1, 1) = 1 + 2 * wood_fourier
                UC(1, 2) = -2 * wood_fourier
                'interior boundary conditions
                CX(1, 1) = -2 * QCeiling(room, i) * 1000 * wood_fourier * CeilingDeltaX(room) / prop_k + CeilingNodeExposed
                For k = 2 To ceilingnodestemp - 1
                    CX(k, 1) = CeilingNodeTemp(k)
                Next k
            End If


            'inside nodes
            k = 2
            For j = 2 To ceilingnodestemp - 1
                If CeilingNodeTemp(j + 1) <= 473 Then
                    prop_k = kwood
                ElseIf CeilingNodeTemp(j + 1) <= 663 Then
                    prop_k = -0.617 + 0.0038 * CeilingNodeTemp(j + 1) - 0.000004 * CeilingNodeTemp(j + 1) ^ 2
                Else
                    prop_k = 0.04429 + 0.0001477 * CeilingNodeTemp(j + 1)
                End If

                moisturecontent = CeilingElementMF(j, 0, i) * mf_init
                WoodDensity = CeilingApparentDensity(j, i)
                'chardensity = 0.63 * WoodDensity / (1 + moisturecontent)

                If CeilingNodeStatus(j + 1) = 1 Then
                    temp = CeilingNodeTemp(j + 1) - 273
                    char_c = 714 + 2.3 * temp - 0.0008 * temp ^ 2 - 0.00000037 * temp ^ 3
                    char_alpha = prop_k / (char_c * WoodDensity)
                    char_fourier = char_alpha * Timestep / (CeilingDeltaX(room)) ^ 2
                    UC(j, k - 1) = -char_fourier
                    UC(j, k) = 1 + 2 * char_fourier
                    UC(j, k + 1) = -char_fourier
                Else
                    wood_c = 101.3 + 3.867 * CeilingNodeTemp(j + 1)
                    wood_c = (wood_c + 4187 * moisturecontent) / (1 + moisturecontent) + (23.55 * (CeilingNodeTemp(j + 1) - 273) - 1326 * moisturecontent + 2417) * moisturecontent

                    'correct for the latent heat of water over the temp range 90-110 C
                    If CeilingNodeTemp(j + 1) >= 273 + 80 And CeilingNodeTemp(j + 1) <= 273 + 100 Then
                        wood_c = wood_c + mf_init * 2257 / 20 * 1000 'J/kgK
                    End If

                    wood_alpha = prop_k / (wood_c * WoodDensity)
                    wood_fourier = wood_alpha * Timestep / (CeilingDeltaX(room)) ^ 2
                    UC(j, k - 1) = -wood_fourier
                    UC(j, k) = 1 + 2 * wood_fourier
                    UC(j, k + 1) = -wood_fourier
                End If

                k = k + 1
            Next j

            If frmOptions1.optLUdecom.Checked = True Then
                'find surface temperatures for the next timestep
                'using method of LU decomposition (preferred)
                Call MatSol(UC, CX, ceilingnodestemp) 'ceiling

            Else
                'find surface temperatures for the next timestep
                'using method of Gauss-Jordan elimination
                Call LINEAR2(ceilingnodestemp, UC, CX, ier)

                If ier = 1 Then MsgBox("singular matrix in implicit_surface_temps_CLTC_char")
            End If

            For j = 1 To ceilingnodestemp
                CeilingNode(room, j + ceilingnodeadjust, i + 1) = CX(j, 1)

            Next j
            For j = 1 To ceilingnodeadjust
                CeilingNode(room, j, i + 1) = chartemp + 273 + 1
            Next

            'store surface temps at next timestep in another array
            CeilingTemp(room, i + 1) = CeilingNode(room, 1 + ceilingnodeadjust, i + 1)

            UnexposedCeilingtemp(room, i + 1) = CeilingNode(room, ceilingnodestemp, i + 1)

            Erase CX
            Erase UC

        Catch ex As Exception
            MsgBox(Err.Description & " Line " & Err.Erl, MsgBoxStyle.Exclamation, "Exception in Implicit_Temps_Ceil_kinetic")
        End Try
    End Sub
    Sub Implicit_Surface_Temps_LWall(ByVal room As Integer, ByVal i As Integer, ByRef FloorNode(,,) As Double)
        '*  ================================================================
        '*      This function updates the surface temperatures, using an
        '*      implicit finite difference method.
        '*      for the lower wall when kinetic CLT model is used, lower wall does not participate in the CLT functionality.
        '*  ================================================================

        Try

            Dim k, j As Integer
            Dim ier As Short
            Dim LF(Wallnodes, Wallnodes) As Double
            Dim FX(Wallnodes, 1) As Double
            Dim wallFourier2, wallOutsideBiot2 As Double

            If HaveWallSubstrate(room) = False Then
                ReDim LF(Wallnodes, Wallnodes)
                ReDim FX(Wallnodes, 1)
            Else
                ReDim LF(2 * Wallnodes - 1, 2 * Wallnodes - 1)
                ReDim FX(2 * Wallnodes - 1, 1)

                wallFourier2 = WallSubConductivity(room) * Timestep / (((WallSubThickness(room) / 1000) / (Wallnodes - 1)) ^ 2 * WallSubDensity(room) * WallSubSpecificHeat(room))
                wallOutsideBiot2 = OutsideConvCoeff * ((WallSubThickness(room) / 1000) / (Wallnodes - 1)) / WallSubConductivity(room)
            End If

            LF(1, 1) = 1 + 2 * WallFourier(room)
            LF(1, 2) = -2 * WallFourier(room)

            If HaveWallSubstrate(room) = False Then
                k = 2
                For j = 2 To Wallnodes - 1
                    LF(j, k - 1) = -WallFourier(room)
                    LF(j, k) = 1 + 2 * WallFourier(room)
                    LF(j, k + 1) = -WallFourier(room)
                    k = k + 1
                Next j
            Else
                k = 2
                For j = 2 To Wallnodes - 1
                    LF(j, k - 1) = -WallFourier(room)
                    LF(j, k) = 1 + 2 * WallFourier(room)
                    LF(j, k + 1) = -WallFourier(room)
                    k = k + 1
                Next j

                j = Wallnodes
                LF(j, k - 1) = -(room)
                LF(j, k) = 1 + WallFourier(room) + WallFourier2
                LF(j, k + 1) = -WallFourier2
                k = k + 1

                For j = Wallnodes + 1 To 2 * Wallnodes - 2
                    LF(j, k - 1) = -wallFourier2
                    LF(j, k) = 1 + 2 * wallFourier2
                    LF(j, k + 1) = -wallFourier2
                    k = k + 1
                Next j
            End If

            If HaveWallSubstrate(room) = False Then
                LF(Wallnodes, Wallnodes - 1) = -2 * WallFourier(room)
                LF(Wallnodes, Wallnodes) = 1 + 2 * WallFourier(room)
            Else
                LF(2 * Wallnodes - 1, 2 * Wallnodes - 2) = -2 * wallFourier2
                LF(2 * Wallnodes - 1, 2 * Wallnodes - 1) = 1 + 2 * wallFourier2
            End If

            FX(1, 1) = -2 * QLowerWall(room, i) * 1000 * WallFourier(room) * WallDeltaX(room) / WallConductivity(room) + LWallNode(room, 1, i)
            If HaveWallSubstrate(room) = False Then
                For k = 2 To Wallnodes - 1
                    FX(k, 1) = LWallNode(room, k, i)
                Next k
            Else
                For k = 2 To 2 * Wallnodes - 2
                    FX(k, 1) = LWallNode(room, k, i)
                Next k
            End If

            If HaveWallSubstrate(room) = False Then
                FX(Wallnodes, 1) = 2 * WallFourier(room) * WallOutsideBiot(room) * ((ExteriorTemp - LWallNode(room, Wallnodes, i)) - Surface_Emissivity(3, room) / OutsideConvCoeff * StefanBoltzmann * (LWallNode(room, Wallnodes, i) ^ 4 - ExteriorTemp ^ 4)) + LWallNode(room, Wallnodes, i)
            Else
                FX(2 * Wallnodes - 1, 1) = 2 * wallFourier2 * wallOutsideBiot2 * ((ExteriorTemp - LWallNode(room, 2 * Wallnodes - 1, i)) - Surface_Emissivity(3, room) / OutsideConvCoeff * StefanBoltzmann * (LWallNode(room, 2 * Wallnodes - 1, i) ^ 4 - ExteriorTemp ^ 4)) + LWallNode(room, 2 * Wallnodes - 1, i)
            End If

            If frmOptions1.optLUdecom.Checked = True Then
                'find surface temperatures for the next timestep
                'using method of LU decomposition (preferred)
                If HaveWallSubstrate(room) = False Then
                    Call MatSol(LF, FX, Wallnodes)
                Else
                    Call MatSol(LF, FX, 2 * Wallnodes - 1)
                End If
            Else
                'find surface temperatures for the next timestep
                'using method of Gauss-Jordan elimination
                If HaveWallSubstrate(room) = False Then
                    Call LINEAR2(Wallnodes, LF, FX, ier)
                Else
                    Call LINEAR2(2 * Wallnodes - 1, LF, FX, ier)
                End If
                If ier = 1 Then MsgBox("Singular matrix in implicit_surface_temps_LWall")
            End If

            If HaveWallSubstrate(room) = False Then
                For j = 1 To Wallnodes
                    LWallNode(room, j, i + 1) = FX(j, 1)
                Next j
            Else
                For j = 1 To 2 * Wallnodes - 1
                    LWallNode(room, j, i + 1) = FX(j, 1)
                Next j
            End If

            LowerWallTemp(room, i + 1) = LWallNode(room, 1, i + 1)

            If HaveWallSubstrate(room) = True Then
                UnexposedLowerwalltemp(room, i + 1) = LWallNode(room, 2 * Wallnodes - 1, i + 1)
            Else
                UnexposedLowerwalltemp(room, i + 1) = LWallNode(room, Wallnodes, i + 1)
            End If

            Erase LF
            Erase FX

        Catch ex As Exception
            MsgBox(Err.Description & " Line " & Err.Erl, MsgBoxStyle.Exclamation, "Exception in Implicit_Surface_Temp_LWall ")
        End Try
    End Sub
    Sub Implicit_Temps_UWall_kinetic(ByVal room As Integer, ByVal i As Integer, ByRef UWallNode(,,) As Double)
        '*  ================================================================
        '*      This function updates the surface temperatures, using an
        '*      implicit finite difference method.
        '*  ================================================================

        Try
            Dim walllayersremaining As Integer
            Dim k, j As Integer
            Dim wallnodestemp As Integer
            Dim NLW As Integer
            Dim ier As Short
            Dim wallnodeadjust As Integer
            Dim moisturecontent, temp As Double
            Dim prop_k As Double 'W/mK
            Dim char_alpha, char_c, char_fourier As Double
            Dim wood_alpha, wood_c, wood_fourier As Double
            Dim Coutbiot, WoodDensity As Double
            Dim kwood As Double

            kwood = WallConductivity(room)
            WoodDensity = WallDensity(room)

            Dim mf_init As Double = 0.1 'initial mc


            If CLTwallpercent > 0 Then
                NLW = WallThickness(room) / 1000 / Lamella 'number of lamella - in two places also in main_program2

                walllayersremaining = NLW - Lamella2 / Lamella + 1

                If walllayersremaining = 0 Then
                    If walllayersremaining = 0 Then Dim Message As String = CStr(tim(stepcount, 1)) & " sec. Wall layer at " & Format(Lamella2, "0.000") & " m delaminates. Simulation terminated. "

                    flagstop = 1 'do not continue
                    Exit Sub
                End If

                wallnodestemp = (Ceilingnodes - 1) * (walllayersremaining) + 1 'remove one layer and recalc number of nodes
                wallnodeadjust = (Wallnodes - 1) * NLW + 1 - wallnodestemp

                'Find DeltaX
                WallDeltaX(room) = walllayersremaining * Lamella / (wallnodestemp - 1)
            Else
                'clt model is on but not for wall
                NLW = 1
                walllayersremaining = 1
                wallnodestemp = Wallnodes
                wallnodeadjust = 0
            End If

            Dim wallNodeTemp(wallnodestemp) As Double
            Dim wallNodeStatus(wallnodestemp) As Integer  'array to identify charred element

            For k = 1 To wallnodestemp
                wallNodeTemp(k) = UWallNode(room, k + wallnodeadjust, i)

                'array to identify charred element
                If wallNodeTemp(k) > 300 + 273 Then wallNodeStatus(k) = 1

            Next

            Dim wallNodeUnExposed As Double = wallNodeTemp(wallnodestemp)
            Dim wallNodeExposed As Double = wallNodeTemp(1)


            If wallNodeTemp(wallnodestemp) <= 473 Then
                prop_k = kwood

            ElseIf wallNodeTemp(wallnodestemp) <= 663 Then
                prop_k = -0.617 + 0.0038 * wallNodeTemp(wallnodestemp) - 0.000004 * wallNodeTemp(wallnodestemp) ^ 2
            Else
                prop_k = 0.04429 + 0.0001477 * wallNodeTemp(wallnodestemp)
            End If

            'Find Biot Numbers -exterior side
            Coutbiot = OutsideConvCoeff * WallDeltaX(room) / prop_k

            Dim UW(wallnodestemp, wallnodestemp) As Double
            Dim WX(wallnodestemp, 1) As Double

            moisturecontent = UWallElementMF(wallnodestemp - 1, 0, i) * mf_init 'mass fraction of the original wet wood
            WoodDensity = WallApparentDensity(wallnodestemp - 1, i)

            'Find Fourier Numbers -exterior side
            If wallNodeStatus(wallnodestemp) = 1 Then
                'char
                temp = wallNodeTemp(wallnodestemp) - 273 'node temp in deg C
                char_c = 714 + 2.3 * temp - 0.0008 * temp ^ 2 - 0.00000037 * temp ^ 3
                char_alpha = prop_k / (char_c * WoodDensity)
                char_fourier = char_alpha * Timestep / (WallDeltaX(room)) ^ 2

                'exterior boundary conditions
                WX(wallnodestemp, 1) = 2 * char_fourier * Coutbiot * ((ExteriorTemp - wallNodeUnExposed) - Surface_Emissivity(2, room) / OutsideConvCoeff * StefanBoltzmann * (wallNodeUnExposed ^ 4 - ExteriorTemp ^ 4)) + wallNodeUnExposed
                UW(wallnodestemp, wallnodestemp - 1) = -2 * char_fourier
                UW(wallnodestemp, wallnodestemp) = 1 + 2 * char_fourier
                UW(1, 1) = 1 + 2 * char_fourier
                UW(1, 2) = -2 * char_fourier
            Else
                'wood
                wood_c = 101.3 + 3.867 * wallNodeTemp(wallnodestemp)
                wood_c = (wood_c + 4187 * moisturecontent) / (1 + moisturecontent) + (23.55 * (wallNodeTemp(wallnodestemp) - 273) - 1326 * moisturecontent + 2417) * moisturecontent

                'correct for the latent heat of water over the temp range 90-110 C
                If wallNodeTemp(wallnodestemp) >= 273 + 80 And wallNodeTemp(wallnodestemp) <= 273 + 100 Then
                    wood_c = wood_c + mf_init * 2257 / 20 * 1000 'J/kgK
                End If

                wood_alpha = prop_k / (wood_c * WoodDensity)
                wood_fourier = wood_alpha * Timestep / (WallDeltaX(room)) ^ 2

                'exterior boundary conditions
                WX(wallnodestemp, 1) = 2 * wood_fourier * Coutbiot * ((ExteriorTemp - wallNodeUnExposed) - Surface_Emissivity(2, room) / OutsideConvCoeff * StefanBoltzmann * (wallNodeUnExposed ^ 4 - ExteriorTemp ^ 4)) + wallNodeUnExposed
                UW(wallnodestemp, wallnodestemp - 1) = -2 * wood_fourier
                UW(wallnodestemp, wallnodestemp) = 1 + 2 * wood_fourier
                UW(1, 1) = 1 + 2 * wood_fourier
                UW(1, 2) = -2 * wood_fourier
            End If

            'exposed side
            If wallNodeTemp(2) <= 473 Then
                prop_k = kwood
            ElseIf wallNodeTemp(2) <= 663 Then
                prop_k = -0.617 + 0.0038 * wallNodeTemp(2) - 0.000004 * wallNodeTemp(2) ^ 2
            Else
                prop_k = 0.04429 + 0.0001477 * wallNodeTemp(2)
            End If

            moisturecontent = UWallElementMF(1, 0, i) * mf_init
            WoodDensity = WallApparentDensity(1, i)

            If wallNodeStatus(2) = 1 Then 'char
                temp = wallNodeTemp(2) - 273 'node temp in deg C
                char_c = 714 + 2.3 * temp - 0.0008 * temp ^ 2 - 0.00000037 * temp ^ 3
                char_alpha = prop_k / (char_c * WoodDensity)
                char_fourier = char_alpha * Timestep / (WallDeltaX(room)) ^ 2
                UW(1, 1) = 1 + 2 * char_fourier
                UW(1, 2) = -2 * char_fourier
                'interior boundary conditions
                WX(1, 1) = -2 * QUpperWall(room, i) * 1000 * char_fourier * WallDeltaX(room) / prop_k + wallNodeExposed
                For k = 2 To wallnodestemp - 1
                    WX(k, 1) = wallNodeTemp(k)
                Next k
            Else
                'wood
                wood_c = 101.3 + 3.867 * wallNodeTemp(2)
                wood_c = (wood_c + 4187 * moisturecontent) / (1 + moisturecontent) + (23.55 * (wallNodeTemp(2) - 273) - 1326 * moisturecontent + 2417) * moisturecontent

                'correct for the latent heat of water over the temp range 90-110 C
                If wallNodeTemp(2) >= 273 + 80 And wallNodeTemp(2) <= 273 + 100 Then
                    wood_c = wood_c + mf_init * 2257 / 20 * 1000 'J/kgK
                End If

                wood_alpha = prop_k / (wood_c * WoodDensity)
                wood_fourier = wood_alpha * Timestep / (WallDeltaX(room)) ^ 2
                UW(1, 1) = 1 + 2 * wood_fourier
                UW(1, 2) = -2 * wood_fourier
                'interior boundary conditions
                WX(1, 1) = -2 * QUpperWall(room, i) * 1000 * wood_fourier * WallDeltaX(room) / prop_k + wallNodeExposed
                For k = 2 To wallnodestemp - 1
                    WX(k, 1) = wallNodeTemp(k)
                Next k
            End If


            'inside nodes
            k = 2
            For j = 2 To wallnodestemp - 1
                If wallNodeTemp(j + 1) <= 473 Then
                    prop_k = kwood
                ElseIf wallNodeTemp(j + 1) <= 663 Then
                    prop_k = -0.617 + 0.0038 * wallNodeTemp(j + 1) - 0.000004 * wallNodeTemp(j + 1) ^ 2
                Else
                    prop_k = 0.04429 + 0.0001477 * wallNodeTemp(j + 1)
                End If

                moisturecontent = UWallElementMF(j, 0, i) * mf_init
                WoodDensity = WallApparentDensity(j, i)

                If wallNodeStatus(j + 1) = 1 Then
                    temp = wallNodeTemp(j + 1) - 273
                    char_c = 714 + 2.3 * temp - 0.0008 * temp ^ 2 - 0.00000037 * temp ^ 3
                    char_alpha = prop_k / (char_c * WoodDensity)
                    char_fourier = char_alpha * Timestep / (WallDeltaX(room)) ^ 2
                    UW(j, k - 1) = -char_fourier
                    UW(j, k) = 1 + 2 * char_fourier
                    UW(j, k + 1) = -char_fourier
                Else
                    wood_c = 101.3 + 3.867 * wallNodeTemp(j + 1)
                    wood_c = (wood_c + 4187 * moisturecontent) / (1 + moisturecontent) + (23.55 * (wallNodeTemp(j + 1) - 273) - 1326 * moisturecontent + 2417) * moisturecontent

                    'correct for the latent heat of water over the temp range 90-110 C
                    If wallNodeTemp(j + 1) >= 273 + 80 And wallNodeTemp(j + 1) <= 273 + 100 Then
                        wood_c = wood_c + mf_init * 2257 / 20 * 1000 'J/kgK
                    End If

                    wood_alpha = prop_k / (wood_c * WoodDensity)
                    wood_fourier = wood_alpha * Timestep / (WallDeltaX(room)) ^ 2
                    UW(j, k - 1) = -wood_fourier
                    UW(j, k) = 1 + 2 * wood_fourier
                    UW(j, k + 1) = -wood_fourier
                End If

                k = k + 1
            Next j

            If frmOptions1.optLUdecom.Checked = True Then
                'find surface temperatures for the next timestep
                'using method of LU decomposition (preferred)
                Call MatSol(UW, WX, wallnodestemp) 'ceiling

            Else
                'find surface temperatures for the next timestep
                'using method of Gauss-Jordan elimination
                Call LINEAR2(wallnodestemp, UW, WX, ier)

                If ier = 1 Then MsgBox("Singular matrix in implicit_Temps_UWall_kinetic")
            End If

            For j = 1 To wallnodestemp
                UWallNode(room, j + wallnodeadjust, i + 1) = WX(j, 1)

            Next j
            For j = 1 To wallnodeadjust
                UWallNode(room, j, i + 1) = chartemp + 273 + 1
            Next

            'store surface temps at next timestep in another array
            Upperwalltemp(room, i + 1) = UWallNode(room, 1 + wallnodeadjust, i + 1)

            UnexposedUpperwalltemp(room, i + 1) = UWallNode(room, wallnodestemp, i + 1)

            Erase WX
            Erase UW

        Catch ex As Exception
            MsgBox(Err.Description & " Line " & Err.Erl, MsgBoxStyle.Exclamation, "Exception in implicit_Temps_UWall_kinetic")
        End Try
    End Sub
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
            Dim DensityInitial As Double = 515 'kg/m3
            'Dim chardensity As Double = 85 'kg/m3
            Dim chardensity As Double = 150 'kg/m3
            Dim DelamDuration As Double = 120 'seconds

            chardensity = DensityInitial * 0.63 / (1 + mf_init(0))

            Dim rmw As Double

            'the ceiling
            'CeilingNode(room, node, timestep) contains the temperature at each node at each timestep
            'CeilingElementMF (element,timsetep) contains the residual mass fraction of each component (relative to its initial value = 1) 

            ReDim Zstart(0 To 3)
            elements = maxceilingnodes - 1
            CeilingWoodMLR_tot(i + 1) = 0


            Dim ceilingexposedpercent As Double = CLTceilingpercent
            Dim ElapsedTime As Double
            Dim DT As Double = CLTceildelamT + flashover_time 'time of delamination
            If DT > flashover_time + 1 Then
                ElapsedTime = tim(i, 1) - DT
                If ElapsedTime < DelamDuration Then 'within 60 s of when delamination happened
                    ceilingexposedpercent = CLTceilingpercent / DelamDuration * (tim(i, 1) - DT)
                End If
            End If

            For count = 1 To elements 'loop through each finite difference element in the ceiling
                If i = 1 Then
                    For m = 1 To 3
                        CeilingResidualMass(count, i) = DensityInitial * mf_init(m) 'initialise
                    Next
                    CeilingApparentDensity(count, i) = DensityInitial
                End If

                For m = 0 To 3
                    Zstart(m) = CeilingElementMF(count, m, i)
                Next
                elementcounter = count

                Call ODE_Solver_Pyrolysis(Zstart, i, "C")

                For m = 0 To 3
                    CeilingElementMF(count, m, i + 1) = Max(Min(Zstart(m), 1), 0) 'residual mass fraction at the next time step
                Next

                'total mass fraction of char residue in this element
                CeilingCharResidue(count, i + 1) = (1 - CeilingElementMF(count, 1, i + 1)) * mf_init(1) * CharYield + (1 - CeilingElementMF(count, 2, i + 1)) * mf_init(2) * CharYield + (1 - CeilingElementMF(count, 3, i + 1)) * mf_init(3) * CharYield

                'total mass (per unit vol) of residual fuel (cellulose, hemicellulose, lignin) in this element 'kg/m3
                CeilingResidualMass(count, i + 1) = DensityInitial * (CeilingElementMF(count, 1, i + 1) * mf_init(1) + CeilingElementMF(count, 2, i + 1) * mf_init(2) + CeilingElementMF(count, 3, i + 1) * mf_init(3)) 'kg/m3

                'mass loss rate of wood fuel over this timestep 'kg/s
                If i > 1 Then CeilingWoodMLR(count, i + 1) = -(CeilingResidualMass(count, i + 1) - CeilingResidualMass(count, i)) / Timestep 'kg/(s.m3)


                CeilingWoodMLR_tot(i + 1) = CeilingWoodMLR_tot(i + 1) + CeilingWoodMLR(count, i + 1) * ceilingexposedpercent / 100 * RoomFloorArea(fireroom) * CeilingThickness(fireroom) / 1000 / elements 'kg/s

                'residual mass of water in this element 'kg/m3
                rmw = CeilingElementMF(count, 0, i + 1) * DensityInitial * mf_init(0) 'kg/m3

                'apparent density of this element 'kg/m3 
                CeilingApparentDensity(count, i + 1) = rmw + CeilingCharResidue(count, i + 1) * DensityInitial + CeilingResidualMass(count, i + 1) 'water + char + solids

                'put a lower limit on the apparent density
                If CeilingApparentDensity(count, i + 1) < chardensity Then CeilingApparentDensity(count, i + 1) = chardensity
            Next

            'the wall
            'UWallNode(room, node, timestep) contains the temperature at each node at each timestep
            'WallElementMF (element,timsetep) contains the residual mass fraction of each component (relative to its initial value = 1) 

            elements = maxwallnodes - 1
            WallWoodMLR_tot(i + 1) = 0

            Dim wallexposedpercent As Double = CLTwallpercent
            DT = CLTwalldelamT + flashover_time 'time of delamination
            If DT > flashover_time + 1 Then
                ElapsedTime = tim(i, 1) - DT
                If ElapsedTime < DelamDuration Then 'within 60 s of when delamination happened
                    wallexposedpercent = CLTwallpercent / DelamDuration * (tim(i, 1) - DT)
                End If
            End If

            For count = 1 To elements 'loop through each finite difference element in the ceiling
                If i = 1 Then
                    For m = 1 To 3
                        WallResidualMass(count, i) = DensityInitial * mf_init(m) 'initialise
                    Next
                    WallApparentDensity(count, i) = DensityInitial
                End If

                For m = 0 To 3
                    Zstart(m) = UWallElementMF(count, m, i)
                Next
                elementcounter = count

                Call ODE_Solver_Pyrolysis(Zstart, i, "W")

                For m = 0 To 3
                    UWallElementMF(count, m, i + 1) = Max(Min(Zstart(m), 1), 0) 'residual mass fraction at the next time step
                Next

                'total mass fraction of char residue in this element
                UWallCharResidue(count, i + 1) = (1 - UWallElementMF(count, 1, i + 1)) * mf_init(1) * CharYield + (1 - UWallElementMF(count, 2, i + 1)) * mf_init(2) * CharYield + (1 - UWallElementMF(count, 3, i + 1)) * mf_init(3) * CharYield

                'total mass (per unit vol) of residual fuel (cellulose, hemicellulose, lignin) in this element 'kg/m3
                WallResidualMass(count, i + 1) = DensityInitial * (UWallElementMF(count, 1, i + 1) * mf_init(1) + UWallElementMF(count, 2, i + 1) * mf_init(2) + UWallElementMF(count, 3, i + 1) * mf_init(3)) 'kg/m3

                'mass loss rate of wood fuel over this timestep 'kg/s
                If i > 1 Then WallWoodMLR(count, i + 1) = -(WallResidualMass(count, i + 1) - WallResidualMass(count, i)) / Timestep 'kg/(s.m3)

                WallWoodMLR_tot(i + 1) = WallWoodMLR_tot(i + 1) + WallWoodMLR(count, i + 1) * wallexposedpercent / 100 * (RoomLength(fireroom) + RoomWidth(fireroom)) * 2 * RoomHeight(fireroom) * WallThickness(fireroom) / 1000 / elements 'kg/s

                'residual mass of water in this element 'kg/m3
                rmw = UWallElementMF(count, 0, i + 1) * DensityInitial * mf_init(0) 'kg/m3

                'apparent density of this element 'kg/m3 
                WallApparentDensity(count, i + 1) = rmw + UWallCharResidue(count, i + 1) * DensityInitial + WallResidualMass(count, i + 1) 'water + char + solids

                'put a lower limit on the apparent density
                If WallApparentDensity(count, i + 1) < chardensity Then WallApparentDensity(count, i + 1) = chardensity
            Next

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in " & Err.Source & " Line " & Err.Erl)
        End Try


    End Sub
    Sub ODE_Solver_Pyrolysis(ByRef Zstart() As Double, i As Integer, ByVal Mat As String)

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

            If Mat = "C" Then
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
            ElseIf Mat = "W" Then
                ' Construct the delegate representing our system of differential equations...
                Dim odeFunctionPyrol As New Func(Of Double, DoubleVector, DoubleVector)(AddressOf RigidPyrol2)
                Dim Soln As RungeKutta45OdeSolver.Solution(Of DoubleMatrix) = Solver.Solve(odeFunctionPyrol, TimeSpan, y0, SolverOptions)

                index = 0

                For var = 1 To Nvariables
                    Y_pyrol(var - 1) = Soln.Y.Col(var - 1).Last
                    index = index + 1
                Next

            End If

            Zstart = Y_pyrol.Clone

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in " & Err.Source & " Line " & Err.Erl)
            flagstop = 1
        End Try

    End Sub
    Function RigidPyrol2(ByVal T As Double, ByVal Y As DoubleVector) As DoubleVector

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
            ElementTemp = (UWallNode(fireroom, elementcounter, stepcount) + UWallNode(fireroom, elementcounter + 1, stepcount)) / 2 'use average temperature of the two adjacent nodes

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

        Static Tsave, mwallsave, mceilingsave, MLRsave, mplume As Double

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

                mplume = (massplumeflow(stepcount - 1, fireroom) - massplumeflow(stepcount - 2, fireroom)) * (T - tim(stepcount - 2, 1)) / Timestep + massplumeflow(stepcount - 2, fireroom)
                Qtemp = mplume * O2MassFraction(fireroom, stepcount - 1, 2) * 13100
                'Qtemp = massplumeflow(stepcount - 1, fireroom) * O2MassFraction(fireroom, stepcount - 1, 2) * 13100
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
            mceiling = (CeilingWoodMLR_tot(stepcount) - CeilingWoodMLR_tot(stepcount - 1)) * (T - tim(stepcount - 1, 1)) / Timestep + CeilingWoodMLR_tot(stepcount - 1)
            mwall = (WallWoodMLR_tot(stepcount) - WallWoodMLR_tot(stepcount - 1)) * (T - tim(stepcount - 1, 1)) / Timestep + WallWoodMLR_tot(stepcount - 1)

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
