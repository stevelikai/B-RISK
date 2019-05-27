Option Strict Off
Option Explicit On
Imports CenterSpace.NMath.Core
Imports CenterSpace.NMath.Stats
Imports System.Math
Imports System.Collections.Generic

Module DFG
    Public vent_space(,) As Single
    Public number_vents As Integer
    Public number_cvents As Integer
    Public seed As Integer = Convert.ToInt16(frmInputs.txtSeed.Text)
    Public tempdatain() As Double
    Public mc_energyyield(,) As Single
    Public mc_sootyield(,) As Single
    Public mc_co2yield(,) As Single
    Public ventclearance As Single
    Public gridsize As Single

    Public Sub populate_items(ByRef n_max As Integer, ByVal itcounter As Integer, ByRef fireload_count As Double, ByRef firemass_count As Double)

        'code to populate the fire room with fuel items
        'coded by C. Wade based on flowchart by A. Robbins
        Try

            Dim item_id As Single = 1.001
            Dim j_max As Integer = 1    'maximum number of types of items in the database NOT USED YET 
            Dim n_database As Integer    'maximum number of individual entries in the database 
            Dim xx As Single
            Dim yy As Single
            Dim j, k, m, n As Integer
            Dim available_items As Integer 'the number of unique items in database that are available for selection
            Dim room As Integer = fireroom
            Dim x_max As Single = CSng(RoomLength(room)) 'room x-dimension
            Dim y_max As Single = CSng(RoomWidth(room))  'room y-dimension

            Dim delta_x As Single = gridsize 'grid size in metres, objects will be centred on a grid coordinate 
            If delta_x = 0 Then Exit Sub

            'initalise
            number_vents = 0 'module level variable

            'get the number of vents for this room
            For j = 1 To NumberRooms + 1
                If j <> room Then
                    number_vents = number_vents + NumberVents(room, j)
                End If
            Next j

            'now we know the number of vents, dimension the arrays
            ReDim vent_space(0 To number_vents, 0 To 3)
            Dim xvent_lower(0 To number_vents) As Single
            Dim xvent_upper(0 To number_vents) As Single
            Dim yvent_lower(0 To number_vents) As Single
            Dim yvent_upper(0 To number_vents) As Single

            For j = 1 To NumberRooms + 1
                If j <> room Then
                    m = 0
                    For k = 1 To NumberVents(room, j)
                        m = m + 1
                        'vents are assumed centrally located in the wall segments for now.
                        'set variables that define the vent locations
                        If VentFace(room, j, m) = 0 Then
                            'front face (bottom)
                            xvent_lower(m) = x_max / 2 - VentWidth(room, j, m) / 2
                            xvent_upper(m) = x_max / 2 + VentWidth(room, j, m) / 2
                            yvent_lower(m) = 0
                            yvent_upper(m) = 0
                        ElseIf VentFace(room, j, m) = 1 Then
                            'right face
                            xvent_lower(m) = x_max
                            xvent_upper(m) = x_max
                            yvent_lower(m) = y_max / 2 - VentWidth(room, j, m) / 2
                            yvent_upper(m) = y_max / 2 + VentWidth(room, j, m) / 2
                        ElseIf VentFace(room, j, m) = 2 Then
                            'rear face (top)
                            xvent_lower(m) = x_max / 2 - VentWidth(room, j, m) / 2
                            xvent_upper(m) = x_max / 2 + VentWidth(room, j, m) / 2
                            yvent_lower(m) = y_max
                            yvent_upper(m) = y_max
                        Else
                            'left face
                            xvent_lower(m) = 0
                            xvent_upper(m) = 0
                            yvent_lower(m) = y_max / 2 - VentWidth(room, j, m) / 2
                            yvent_upper(m) = y_max / 2 + VentWidth(room, j, m) / 2
                        End If
                    Next
                End If
            Next j

            If mc_go = False Then 'only sample if we are not running the simulation
                'find number of items in the room based on fire load density
                Call sample_fireload() 'sample a single value for the FLED based on assigned distributions for this iteration only
            End If

            Dim counter As Integer
            Dim spaces As Integer
            Dim n1 As Integer
            Dim x As Single
            Dim y As Single
            Dim random_value As Double
            Dim in_clear  'Flag for when item is in vent space
            Dim room_fit As Integer 'flag to check whether the item fits into the room

            Dim PE(,) As Single
            Dim orientation As Integer  '=1 if item length dimension is in the room x direction; =2 if item length dimension is in the room y direction
            'Dim ventclearspace As Single = CSng(frmPopulate.txtVentClearance.Text) 'm space in front of vent to keep clear 
            Dim ventclearspace As Single = ventclearance 'm space in front of vent to keep clear 

            'locate regions adjacent to vents where items will not be permitted
            For vent = 1 To number_vents
                If xvent_lower(vent) = xvent_upper(vent) Then 'vent is located on a wall in the y-direction
                    If Abs(xvent_lower(vent)) < delta_x Then 'vent is located on the y-axis
                        vent_space(vent, 0) = xvent_lower(vent)
                        vent_space(vent, 1) = xvent_upper(vent) + ventclearspace
                        vent_space(vent, 2) = yvent_lower(vent)
                        vent_space(vent, 3) = yvent_upper(vent)
                    Else  'vent is located at y-max
                        vent_space(vent, 0) = xvent_lower(vent) - ventclearspace
                        vent_space(vent, 1) = xvent_upper(vent)
                        vent_space(vent, 2) = yvent_lower(vent)
                        vent_space(vent, 3) = yvent_upper(vent)
                    End If
                Else  'vent is located on a wall in the x-direction
                    If Abs(yvent_lower(vent)) < delta_x Then 'vent is located on the x-axis
                        vent_space(vent, 0) = xvent_lower(vent)
                        vent_space(vent, 1) = xvent_upper(vent)
                        vent_space(vent, 2) = yvent_lower(vent)
                        vent_space(vent, 3) = yvent_upper(vent) + ventclearspace
                    Else
                        vent_space(vent, 0) = xvent_lower(vent)
                        vent_space(vent, 1) = xvent_upper(vent)
                        vent_space(vent, 2) = yvent_lower(vent) - ventclearspace
                        vent_space(vent, 3) = yvent_upper(vent)
                    End If
                End If
            Next

            Dim items_count As Integer = 0

            Dim oitemdistributions As List(Of oDistribution)
            oitemdistributions = ItemDB.GetItemDistributions

            Dim oItems As New List(Of oItem)
            oItems = ItemDB.GetItemsv2() 'get list of items that can be selected from
            n_database = oItems.Count 'number of distinct items 
            available_items = n_database 'number of distinct items that are available

            If usepowerlawdesignfire = True Then
                available_items = 1
                n_max = 1
                n_database = 1
            End If
            'available_items = 0
            'For Each oItem In oItems
            '    available_items = oItem.maxnumitem + available_items
            'Next

            Dim centre_w(0 To n_max - 1) As Single 'half width of item
            Dim centre_l(0 To n_max - 1) As Single 'half length of item
            Dim CW(0 To n_max - 1) As Single
            Dim CL(0 To n_max - 1) As Single
            Dim item_space(0 To 3, 0 To n_max - 1) As Single
            Dim itemcounter(0 To n_database) As Integer 'to keep track of how many of a specific item have been added to the room
            Dim thisitem As Integer

            'If n_database < 2 Then
            '    MsgBox("The populate room module requires more than one item.", MsgBoxStyle.Exclamation)
            '    Exit Sub
            'End If


            items_count = 1
            n = 1

            'exit loop when max fire load in room is reached.
            'randomly select item from collection
            'calc cum. fire load  -  mass x hoc 
            'do we have max number of a specific item per room? eg. tv
            'individual items may have a probability of being present in the room
            Do While (fireload_count < FLED * x_max * y_max) And available_items > 0  'MJ
                'get new item
                If n_database > 1 Then
                    Dim Rui As New RandGenUniform(1, n_database) 'uniform distribution to generate random numbers between two values - use Ru.next to generate the next number

                    n = items_count
                    If firstitem > n_database Then firstitem = 1

                    If n = 1 And firstitem > 0 Then
                        thisitem = firstitem 'specified first ignited item
                        firstitemtemp = thisitem
                    ElseIf n = 1 Then
                        thisitem = Rui.Next 'select an item id at random from the item list
                        firstitemtemp = thisitem
                    Else
                        thisitem = Rui.Next 'select an item id at random from the item list
                    End If

                    Do While itemcounter(thisitem) + 1 > oItems(thisitem - 1).maxnumitem
                        thisitem = Rui.Next 'select an item id at random from the item list

                    Loop

                Else
                    n = items_count
                    thisitem = 1
                End If

                itemcounter(thisitem) = itemcounter(thisitem) + 1 'increment by 1
                If itemcounter(thisitem) = oItems(thisitem - 1).maxnumitem Then
                    available_items = available_items - 1
                End If

                ReDim Preserve centre_w(0 To n - 1)
                ReDim Preserve centre_l(0 To n - 1)
                ReDim Preserve CW(0 To n - 1)
                ReDim Preserve CL(0 To n - 1)
                ReDim Preserve item_space(0 To 3, 0 To n - 1)
                ReDim Preserve item_location(0 To 8, 0 To n) 'array defined in Global.vb

                centre_w(n - 1) = 0.5 * oItems(thisitem - 1).width 'half width of item
                centre_l(n - 1) = 0.5 * oItems(thisitem - 1).length 'half length of item

                counter = 1 'counts all valid locations on the grid for the current item
                'Flag to check whether the item fits, room_fit = 0 indicates item has not been fitted or does not fit, room_fit = 1 indicates item fits
                room_fit = 0

                ReDim PE(0 To 4, 0 To counter)

                For orientation = 1 To 2
                    'switching the orientation allows us to try and fit the item into available spaces either way
                    If orientation = 1 Then 'item length dimension is in the room x direction
                        CW(n - 1) = centre_w(n - 1)
                        CL(n - 1) = centre_l(n - 1)
                    Else 'item length dimension is in the room y direction
                        CW(n - 1) = centre_l(n - 1)
                        CL(n - 1) = centre_w(n - 1)
                    End If

                    'the next 4 lines define a rectangular zone within the room where the centre of an item can be placed
                    Dim xlow As Single = CL(n - 1)
                    Dim xhigh As Single = (x_max - CL(n - 1))
                    Dim ylow As Single = CW(n - 1)
                    Dim yhigh As Single = (y_max - CW(n - 1))

                    'iterate through grid points within the valid zone
                    For x = 0 To x_max Step delta_x 'for the width of the item orientated in the x-direction
                        If x >= xlow And x <= xhigh Then

                            For y = 0 To y_max Step delta_x 'for the length of the item orientated in the y-direction
                                If y >= ylow And y <= yhigh Then

                                    xx = Round(x * 1000) * 0.001 'unresolved issues with rounding in for loop step counter ????
                                    yy = Round(y * 1000) * 0.001 'fix it by creating new var to avoid excess floating point digits in the x,y values

                                    'Calculate the four points that defines the area the item takes up, if located in that space
                                    item_space(0, n - 1) = xx - CL(n - 1)
                                    item_space(1, n - 1) = xx + CL(n - 1)
                                    item_space(2, n - 1) = yy - CW(n - 1)
                                    item_space(3, n - 1) = yy + CW(n - 1)

                                    in_clear = 0
                                    spaces = 0

                                    Dim X1(0 To spaces) As Single
                                    Dim X2(0 To spaces) As Single
                                    Dim Y1(0 To spaces) As Single
                                    Dim Y2(0 To spaces) As Single

                                    For vent = 1 To number_vents
                                        spaces = spaces + 1
                                        ReDim Preserve X1(0 To spaces + 1)
                                        ReDim Preserve X2(0 To spaces + 1)
                                        ReDim Preserve Y1(0 To spaces + 1)
                                        ReDim Preserve Y2(0 To spaces + 1)
                                        'define a rectangular clear zones reserved for each vent
                                        X1(spaces) = vent_space(vent, 0)
                                        X2(spaces) = vent_space(vent, 1)
                                        Y1(spaces) = vent_space(vent, 2)
                                        Y2(spaces) = vent_space(vent, 3)
                                    Next

                                    If n > 1 Then
                                        'define a rectangular clear zones reserved for other items
                                        'check item can fit around other existing items
                                        For other_item = 1 To n - 1
                                            spaces = spaces + 1
                                            ReDim Preserve X1(0 To spaces + 1)
                                            ReDim Preserve X2(0 To spaces + 1)
                                            ReDim Preserve Y1(0 To spaces + 1)
                                            ReDim Preserve Y2(0 To spaces + 1)
                                            X1(spaces) = item_location(3, other_item)
                                            X2(spaces) = item_location(4, other_item)
                                            Y1(spaces) = item_location(5, other_item)
                                            Y2(spaces) = item_location(6, other_item)
                                        Next
                                    End If

                                    'check each reserved block occupied by vent or other item
                                    For n1 = 1 To spaces

                                        If (item_space(0, n - 1) < X2(n1)) Then
                                            'could be blocked
                                            If (item_space(1, n - 1) < X1(n1)) Then
                                                'all clear
                                            Else
                                                'could be blocked
                                                If (item_space(2, n - 1) > Y2(n1)) Then
                                                    'all clear
                                                Else
                                                    'could be blocked
                                                    If (item_space(3, n - 1) < Y1(n1)) Then
                                                        'all clear
                                                    Else
                                                        'is blocked
                                                        in_clear = 1
                                                    End If
                                                End If
                                            End If
                                        Else
                                            'all clear
                                        End If

                                        If in_clear = 1 Then ' this location no good - move to next potential item centre location and check if item fits
                                            Exit For
                                        End If
                                    Next

                                    If in_clear = 1 Then 'still no good - move to next potential item centre location position and repeat checks
                                        'continue looping
                                    Else
                                        'position is valid
                                        room_fit = 1 'Flag indicating that there was an orientation and location where the item fitted into the room

                                        'If (orientation = 1 And ((yy <= CW(n - 1)) Or (yy >= y_max - CW(n - 1)))) Or (orientation = 2 And ((xx <= CL(n - 1)) Or (xx >= x_max - CL(n - 1)))) Then
                                        If (orientation = 1 And ((yy <= CW(n - 1) + delta_x) Or (yy >= y_max - CW(n - 1) + -delta_x))) Or (orientation = 2 And ((xx <= CL(n - 1) + delta_x) Or (xx >= x_max - CL(n - 1) - delta_x))) Then
                                            'item is located close to a wall 
                                            If counter = 1 Then
                                                PE(1, counter) = oItems(thisitem - 1).prob * 10
                                            Else
                                                'calculated effective cumulative probability density value for the item centre location, located close to a wall
                                                PE(1, counter) = PE(1, counter - 1) + oItems(thisitem - 1).prob * 10
                                            End If
                                        Else
                                            'item is not located close to a wall
                                            If counter = 1 Then
                                                PE(1, counter) = (1 - oItems(thisitem - 1).prob) * 10
                                            Else
                                                PE(1, counter) = PE(1, counter - 1) + (1 - oItems(thisitem - 1).prob) * 10 'calculated effective cumulative probability density value for the item centre location, not located close to a wall
                                            End If
                                        End If

                                        PE(2, counter) = xx 'a valid centre location, x-value
                                        PE(3, counter) = yy 'a valid centre location, y-value
                                        PE(4, counter) = orientation
                                        counter = counter + 1  'only counts valid positions
                                        ReDim Preserve PE(0 To 4, 0 To counter)
                                    End If
                                End If 'new y
                            Next 'y increment
                        End If 'new x
                    Next 'x increment 
                Next 'orientation

                If room_fit = 1 Then 'item fits into room in at least 1 orientation

                    'Calculate a random number between 0 and PE(counter-1,1)
                    If PE(1, counter - 1) > 0 Then
                        Dim Ru As New RandGenUniform(0, PE(1, counter - 1)) 'uniform distribution to generate random numbers between two values - use Ru.next to generate the next number
                        random_value = Ru.Next
                    Else
                        random_value = 0
                    End If

                    For counter_no = 1 To counter - 1
                        If random_value < PE(1, counter_no) Then 'have found the centre location for the item
                            item_location(1, n) = PE(2, counter_no) 'storing the item x-centre location 
                            item_location(2, n) = PE(3, counter_no) 'storing the item y-centre location 
                            item_location(7, n) = PE(4, counter_no)  'storing the orientation

                            If item_location(7, n) = 1 Then 'item length dimension is in the room x direction
                                CW(n - 1) = centre_w(n - 1)
                                CL(n - 1) = centre_l(n - 1)
                            Else 'item length dimension is in the room y direction
                                CW(n - 1) = centre_l(n - 1)
                                CL(n - 1) = centre_w(n - 1)
                            End If

                            item_location(3, n) = PE(2, counter_no) - CL(n - 1)   'x1 stored to reduce possible calculations
                            item_location(4, n) = PE(2, counter_no) + CL(n - 1)   'x2 stored to reduce possible calculations  
                            item_location(5, n) = PE(3, counter_no) - CW(n - 1)   'y1 stored to reduce possible calculations
                            item_location(6, n) = PE(3, counter_no) + CW(n - 1)   'y2 stored to reduce possible calculations  

                            item_location(8, n) = thisitem 'storing the item database ID

                            NumberObjects = n
                            Resize_Objects()

                            ObjLabel(n) = oItems(thisitem - 1).userlabel
                            ObjLength(n) = oItems(thisitem - 1).length
                            ObjWidth(n) = oItems(thisitem - 1).width
                            ObjHeight(n) = oItems(thisitem - 1).height

                            ObjDimX(n) = CDec(item_location(1, n) - ObjLength(n) / 2)
                            ObjDimY(n) = CDec(item_location(2, n) - ObjWidth(n) / 2)

                            If n = 1 Then
                                'save the location of the first item in each iteration, to use for calculating sprinkler radial distance later
                                'Item1X(itcounter - 1) = item_location(2, n) 'error
                                'Item1Y(itcounter - 1) = item_location(3, n) 'error
                                Item1X(itcounter - 1) = item_location(1, n)
                                Item1Y(itcounter - 1) = item_location(2, n)
                            End If

                            ObjElevation(n) = oItems(thisitem - 1).elevation
                            FireHeight(n) = oItems(thisitem - 1).elevation 'taking base on the fire to be at the elevation height
                            ObjCRF(n) = oItems(thisitem - 1).critflux
                            ObjCRFauto(n) = oItems(thisitem - 1).critfluxauto
                            ObjFTPindexpilot(n) = oItems(thisitem - 1).ftpindexpilot
                            ObjFTPindexauto(n) = oItems(thisitem - 1).ftpindexauto
                            ObjFTPlimitpilot(n) = oItems(thisitem - 1).ftplimitpilot
                            ObjFTPlimitauto(n) = oItems(thisitem - 1).ftplimitauto
                            ObjectDescription(n) = oItems(thisitem - 1).description
                            ObjectMass(n) = oItems(thisitem - 1).mass
                            ObjectItemID(n) = oItems(thisitem - 1).id
                            ObjectRLF(n) = oItems(thisitem - 1).radlossfrac
                            ObjectLHoG(n) = oItems(thisitem - 1).LHoG
                            ObjectMLUA(0, n) = oItems(thisitem - 1).constantA
                            ObjectMLUA(1, n) = oItems(thisitem - 1).constantB
                            ObjectMLUA(2, n) = oItems(thisitem - 1).HRRUA
                            ObjectWindEffect(n) = oItems(thisitem - 1).windeffect
                            ObjectPyrolysisOption(n) = oItems(thisitem - 1).pyrolysisoption
                            ObjectPoolDensity(n) = oItems(thisitem - 1).pooldensity
                            ObjectPoolDiameter(n) = oItems(thisitem - 1).pooldiameter
                            ObjectPoolFBMLR(n) = oItems(thisitem - 1).poolFBMLR
                            ObjectPoolRamp(n) = oItems(thisitem - 1).poolramp
                            ObjectPoolVolume(n) = oItems(thisitem - 1).poolvolume
                            ObjectPoolVapTemp(n) = oItems(thisitem - 1).poolvaptemp

                            FireLocation(n) = 0 'centre
                            If item_location(3, n) < delta_x Then
                                If item_location(5, n) < delta_x Or item_location(6, n) > y_max - delta_x Then
                                    FireLocation(n) = 2 'corner
                                Else
                                    FireLocation(n) = 1 'wall
                                End If
                            End If
                            If item_location(4, n) > x_max - delta_x Then
                                If item_location(5, n) < delta_x Or item_location(6, n) > y_max - delta_x Then
                                    FireLocation(n) = 2 'corner
                                Else
                                    FireLocation(n) = 1 'wall
                                End If
                            End If
                            If item_location(5, n) < delta_x Then
                                If item_location(3, n) < delta_x Or item_location(4, n) > x_max - delta_x Then
                                    FireLocation(n) = 2 'corner
                                Else
                                    FireLocation(n) = 1 'wall
                                End If
                            End If
                            If item_location(6, n) > y_max - delta_x Then
                                If item_location(3, n) < delta_x Or item_location(4, n) > x_max - delta_x Then
                                    FireLocation(n) = 2 'corner
                                Else
                                    FireLocation(n) = 1 'wall
                                End If
                            End If

                            Dim str_data As String() = oItems(thisitem - 1).hrr.Split(New [Char]() {" "c, ","c, ":"c, CChar(vbCrLf)})
                            Dim s, q, numberpoints As Integer
                            q = str_data.Count - 1
                            If q >= 0 Then
                                s = 0
                                numberpoints = 0
                                NumberDataPoints(n) = 0
                                Do While (numberpoints < q And numberpoints < 1000)
                                    numberpoints = numberpoints + 2
                                    If IsNumeric(str_data(numberpoints - 2)) And IsNumeric(str_data(numberpoints - 1)) Then
                                        HeatReleaseData(1, NumberDataPoints(n) + 1, n) = CDbl(str_data(numberpoints - 2)) 'time
                                        HeatReleaseData(2, NumberDataPoints(n) + 1, n) = CDbl(str_data(numberpoints - 1)) 'hrr
                                    End If
                                    s = s + 1
                                    NumberDataPoints(n) = CShort(NumberDataPoints(n) + 1)
                                Loop
                            End If

                            'Call SampleFireData2_LHS(oitemdistributions, oItems, thisitem, n, itcounter)
                            If mc_item_hoc IsNot Nothing Then
                                'EnergyYield(n) = mc_energyyield(itcounter - 1, thisitem - 1)
                                'SootYield(n) = mc_sootyield(itcounter - 1, thisitem - 1)
                                'CO2Yield(n) = mc_co2yield(itcounter - 1, thisitem - 1)
                                EnergyYield(n) = mc_item_hoc(thisitem - 1, itcounter - 1)
                                SootYield(n) = mc_item_soot(thisitem - 1, itcounter - 1)
                                CO2Yield(n) = mc_item_co2(thisitem - 1, itcounter - 1)

                                ObjectLHoG(n) = mc_item_lhog(thisitem - 1, itcounter - 1)
                                ObjectRLF(n) = mc_item_RLF(thisitem - 1, itcounter - 1)
                                ObjectMLUA(2, n) = mc_item_hrrua(thisitem - 1, itcounter - 1)

                                'HCNuserYield(n) = 0
                                'fireload_count = oItems(thisitem - 1).mass * oItems(thisitem - 1).hoc + fireload_count 'MJ   
                                'fireload_count = oItems(thisitem - 1).mass * mc_energyyield(itcounter - 1, thisitem - 1) + fireload_count 'MJ   
                                If usepowerlawdesignfire = False Then
                                    fireload_count = oItems(thisitem - 1).mass * EnergyYield(n) + fireload_count 'MJ   
                                Else
                                    fireload_count = FLED * x_max * y_max + fireload_count 'MJ   
                                End If

                            Else
                                For Each oDistribution In oitemdistributions
                                    If oDistribution.varname = "heat of combustion" And oDistribution.id = thisitem Then
                                        If usepowerlawdesignfire = False Then
                                            fireload_count = oItems(thisitem - 1).mass * oDistribution.varvalue + fireload_count 'MJ   
                                        Else
                                            fireload_count = FLED * x_max * y_max + fireload_count 'MJ   
                                        End If
                                        'Exit For
                                    End If
                                    If oDistribution.varname = "soot yield" And oDistribution.id = thisitem Then
                                        SootYield(n) = oDistribution.varvalue
                                    End If
                                    If oDistribution.varname = "co2 yield" And oDistribution.id = thisitem Then
                                        CO2Yield(n) = oDistribution.varvalue
                                    End If
                                    If oDistribution.varname = "Latent Heat of Gasification" And oDistribution.id = thisitem Then
                                        ObjectLHoG(n) = oDistribution.varvalue
                                    End If
                                    If oDistribution.varname = "Radiant Loss Fraction" And oDistribution.id = thisitem Then
                                        ObjectRLF(n) = oDistribution.varvalue
                                    End If
                                    If oDistribution.varname = "HRRUA" And oDistribution.id = thisitem Then
                                        ObjectMLUA(2, n) = oDistribution.varvalue
                                    End If
                                Next
                            End If

                            If usepowerlawdesignfire = False Then
                                firemass_count = oItems(thisitem - 1).mass + firemass_count
                            Else
                                firemass_count = fireload_count / EnergyYield(1)
                            End If

                            items_count = items_count + 1

                            Exit For 'no need to check through any more of P(counter_no,1) values, so exit the For-loop

                        Else
                            'Stop
                        End If


                    Next
                    'Stop

                Else
                    'item doesn't fit in room, so prevent further sampling of it by changing the itemcounter to its max value
                    n = n - 1
                    'oItems(thisitem - 1).maxnumitem = 0
                    itemcounter(thisitem) = oItems(thisitem - 1).maxnumitem
                    available_items = available_items - 1

                End If
                'next item
            Loop
            'save actual fire load

            n = ObjectIgnTime.Count - 1

            n_max = n

            'calculate length of midpoint vectors from source to target items

            If n_max > 1 Then
                ReDim vectorlength(0 To n_max, 0 To n_max) 'target, source
                Dim x1, x2, y1, y2 As Single
                Dim s As Integer

                For i = 1 To n_max  'i is the target item
                    For s = 1 To n_max
                        x1 = item_location(1, s) 'source item
                        y1 = item_location(2, s) 'source item
                        x2 = item_location(1, i)
                        y2 = item_location(2, i)
                        ' vectorlength(i, 0) = Sqrt((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
                        vectorlength(i, s) = Sqrt((x2 - x1) ^ 2 + (y2 - y1) ^ 2)

                        'less 1/2 of the max target dimension
                        'If ObjLength(i + 2) > ObjWidth(i + 2) Then
                        '    vectorlength(i, 1) = vectorlength(i, 0) - 0.5 * ObjLength(i + 2)
                        'Else
                        '    vectorlength(i, 1) = vectorlength(i, 0) - 0.5 * ObjWidth(i + 2)
                        'End If

                        'If frmInputs.chkIgniteTargets.Checked = True Then
                        'ObjectIgnTime(i + 2) = vectorlength(i, 0) * 100 'temp, fictitious ignition time
                        'Else
                    Next 'source
                    If i > 1 Then ObjectIgnTime(i) = SimTime
                    'End If
                Next 'target


            End If

            'If 9 * (n_max + 1) <> item_location.Length Then Stop

            If FlagSimStop = True Then
                frmPopulate.txtNumber_Items.Text = n_max.ToString 'display number of items
                'frmPopulate.txtFLED.Text = FLED.ToString 'display FLED
                frmPopulate.txtFLED.Text = fireload_count / (x_max * y_max)
                'frmPopulate.Refresh() 'redraw graphics
            End If
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in DFG.vb populate_items")
        End Try
    End Sub
    Public Sub populate_items_manual(ByRef n_max As Integer, ByVal itcounter As Integer, ByRef fireload_count As Double, ByRef firemass_count As Double)

        'code to populate the fire room with fuel items where the position of the item is specified
        Try

            Dim item_id As Single = 1.001
            Dim j_max As Integer = 1    'maximum number of types of items in the database NOT USED YET 
            Dim n_database As Integer    'maximum number of individual entries in the database 
            Dim j, n As Integer
            Dim available_items As Integer
            Dim room As Integer = fireroom
            Dim x_max As Single = CSng(RoomLength(room)) 'room x-dimension
            Dim y_max As Single = CSng(RoomWidth(room))  'room y-dimension

            'check for valid inputs

            'If IsNumeric(frmPopulate.txtVentClearance.Text) = False Then
            '    MsgBox("Vent clearance value invalid")
            '    Exit Sub
            'End If
            'If IsNumeric(frmPopulate.txtGridSize.Text) = False Then
            '    MsgBox("Grid size value invalid")
            '    Exit Sub
            'End If

            Dim delta_x As Single = gridsize 'grid size in metres, objects will be centred on a grid coordinate 

            'initalise
            number_vents = 0 'module level variable

            'get the number of vents for this room
            For j = 1 To NumberRooms + 1
                If j <> room Then
                    number_vents = number_vents + NumberVents(room, j)
                End If
            Next j

            'now we know the number of vents, dimension the arrays
            ReDim vent_space(0 To number_vents, 0 To 3)
            Dim xvent_lower(0 To number_vents) As Single
            Dim xvent_upper(0 To number_vents) As Single
            Dim yvent_lower(0 To number_vents) As Single
            Dim yvent_upper(0 To number_vents) As Single


            Dim counter As Integer
            Dim room_fit As Integer 'flag to check whether the item fits into the room

            'Dim fireload_count As Double = 0
            Dim items_count As Integer = 0

            Dim oitemdistributions As New List(Of oDistribution)
            oitemdistributions = ItemDB.GetItemDistributions
            Dim oItems As New List(Of oItem)
            oItems = ItemDB.GetItemsv2() 'get list of items that can be selected from

            n_database = oItems.Count
            available_items = n_database 'number of distinct items that are available

            If usepowerlawdesignfire = True Then
                available_items = 1
                n_max = 1
                n_database = 1
            End If

            Dim centre_w(0 To n_max - 1) As Single 'half width of item
            Dim centre_l(0 To n_max - 1) As Single 'half length of item
            Dim CW(0 To n_max - 1) As Single
            Dim CL(0 To n_max - 1) As Single
            Dim item_space(0 To 3, 0 To n_max - 1) As Single
            Dim itemcounter(0 To n_database) As Integer 'to keep track of how many of a specific item have been added to the room
            Dim thisitem As Integer

            If n_database < 2 Then
                'MsgBox("The populate room module requires more than one item.", MsgBoxStyle.Exclamation)
                'Exit Sub
            End If


            items_count = 1
            n = 1

            Do While items_count <= n_database
                'get new item
                If n_database > 1 Then
                    Dim Rui As New RandGenUniform(1, n_database) 'uniform distribution to generate random numbers between two values - use Ru.next to generate the next number

                    n = items_count
                    thisitem = n

                    If n = 1 And firstitem > 0 Then
                        thisitem = firstitem 'specified first ignited item
                        firstitemtemp = thisitem
                    ElseIf n = 1 Then
                        thisitem = Rui.Next 'select an item id at random from the item list
                        firstitemtemp = thisitem
                    Else
                        thisitem = Rui.Next 'select an item id at random from the item list
                    End If

                    Do While itemcounter(thisitem) + 1 > 1
                        thisitem = Rui.Next 'select an item id at random from the item list
                    Loop
                Else
                    thisitem = 1
                End If

                itemcounter(thisitem) = itemcounter(thisitem) + 1
                If itemcounter(thisitem) = 1 Then
                    available_items = available_items - 1
                End If

                ReDim Preserve item_location(0 To 8, 0 To n) 'array defined in Global.vb

                counter = 1 'counts all valid locations on the grid for the current item
                'Flag to check whether the item fits, room_fit = 0 indicates item has not been fitted or does not fit, room_fit = 1 indicates item fits
                room_fit = 0

                item_location(1, n) = oItems(thisitem - 1).xleft + oItems(thisitem - 1).length / 2 'centre point x
                item_location(2, n) = oItems(thisitem - 1).ybottom + oItems(thisitem - 1).width / 2 'centre point y
                item_location(3, n) = oItems(thisitem - 1).xleft  'x1 stored to reduce possible calculations
                item_location(4, n) = oItems(thisitem - 1).xleft + oItems(thisitem - 1).length    'x2 stored to reduce possible calculations  
                item_location(5, n) = oItems(thisitem - 1).ybottom  'y1 stored to reduce possible calculations
                item_location(6, n) = oItems(thisitem - 1).ybottom + oItems(thisitem - 1).width   'y2 stored to reduce possible calculations  

                item_location(8, n) = thisitem 'storing the item database ID

                NumberObjects = n
                Resize_Objects()

                ObjLabel(n) = oItems(thisitem - 1).userlabel
                ObjLength(n) = oItems(thisitem - 1).length
                ObjWidth(n) = oItems(thisitem - 1).width
                ObjHeight(n) = oItems(thisitem - 1).height

                ObjDimX(n) = CDec(item_location(1, n) - ObjLength(n) / 2)
                ObjDimY(n) = CDec(item_location(2, n) - ObjWidth(n) / 2)

                If n = 1 Then
                    'save the location of the first item in each iteration, to use for calculating sprinkler radial distance later
                    Item1X(itcounter - 1) = item_location(1, n)
                    Item1Y(itcounter - 1) = item_location(2, n)
                End If

                ObjElevation(n) = oItems(thisitem - 1).elevation
                FireHeight(n) = oItems(thisitem - 1).elevation 'taking base on the fire to be at the elevation height
                ObjCRF(n) = oItems(thisitem - 1).critflux
                ObjCRFauto(n) = oItems(thisitem - 1).critfluxauto
                ObjFTPindexpilot(n) = oItems(thisitem - 1).ftpindexpilot
                ObjFTPindexauto(n) = oItems(thisitem - 1).ftpindexauto
                ObjFTPlimitpilot(n) = oItems(thisitem - 1).ftplimitpilot
                ObjFTPlimitauto(n) = oItems(thisitem - 1).ftplimitauto
                ObjectDescription(n) = oItems(thisitem - 1).description
                ObjectItemID(n) = oItems(thisitem - 1).id
                ObjectRLF(n) = oItems(thisitem - 1).radlossfrac
                ObjectLHoG(n) = oItems(thisitem - 1).LHoG
                ObjectMLUA(0, n) = oItems(thisitem - 1).constantA
                ObjectMLUA(1, n) = oItems(thisitem - 1).constantB
                ObjectMLUA(2, n) = oItems(thisitem - 1).HRRUA
                ObjectWindEffect(n) = oItems(thisitem - 1).windeffect
                ObjectPyrolysisOption(n) = oItems(thisitem - 1).pyrolysisoption
                ObjectPoolDensity(n) = oItems(thisitem - 1).pooldensity
                ObjectPoolDiameter(n) = oItems(thisitem - 1).pooldiameter
                ObjectPoolFBMLR(n) = oItems(thisitem - 1).poolFBMLR
                ObjectPoolRamp(n) = oItems(thisitem - 1).poolramp
                ObjectPoolVolume(n) = oItems(thisitem - 1).poolvolume
                ObjectPoolVapTemp(n) = oItems(thisitem - 1).poolvaptemp
                ObjectMass(n) = oItems(thisitem - 1).mass

                FireLocation(n) = 0 'centre
                If item_location(3, n) < delta_x Then
                    If item_location(5, n) < delta_x Or item_location(6, n) > y_max - delta_x Then
                        FireLocation(n) = 2 'corner
                    Else
                        FireLocation(n) = 1 'wall
                    End If
                End If
                If item_location(4, n) > x_max - delta_x Then
                    If item_location(5, n) < delta_x Or item_location(6, n) > y_max - delta_x Then
                        FireLocation(n) = 2 'corner
                    Else
                        FireLocation(n) = 1 'wall
                    End If
                End If
                If item_location(5, n) < delta_x Then
                    If item_location(3, n) < delta_x Or item_location(4, n) > x_max - delta_x Then
                        FireLocation(n) = 2 'corner
                    Else
                        FireLocation(n) = 1 'wall
                    End If
                End If
                If item_location(6, n) > y_max - delta_x Then
                    If item_location(3, n) < delta_x Or item_location(4, n) > x_max - delta_x Then
                        FireLocation(n) = 2 'corner
                    Else
                        FireLocation(n) = 1 'wall
                    End If
                End If

                Dim s, q, numberpoints As Integer
                Dim str_data As String()


                If FuelResponseEffects = True Then
                    str_data = oItems(thisitem - 1).mlrfreeburn.Split(New [Char]() {" "c, ","c, ":"c, CChar(vbCrLf)})
                    q = str_data.Count - 1
                    If q >= 0 Then
                        s = 0
                        numberpoints = 0
                        NumberDataPoints(n) = 0
                        Do While (numberpoints < q And numberpoints < 1000)
                            numberpoints = numberpoints + 2
                            If IsNumeric(str_data(numberpoints - 2)) And IsNumeric(str_data(numberpoints - 1)) Then
                                MLRData(1, NumberDataPoints(n) + 1, n) = CDbl(str_data(numberpoints - 2)) 'time
                                MLRData(2, NumberDataPoints(n) + 1, n) = CDbl(str_data(numberpoints - 1)) 'mlr
                            End If
                            s = s + 1
                            NumberDataPoints(n) = CShort(NumberDataPoints(n) + 1)
                        Loop
                    End If
                Else
                    str_data = oItems(thisitem - 1).hrr.Split(New [Char]() {" "c, ","c, ":"c, CChar(vbCrLf)})
                    q = str_data.Count - 1
                    If q >= 0 Then
                        s = 0
                        numberpoints = 0
                        NumberDataPoints(n) = 0
                        Do While (numberpoints < q And numberpoints < 1000)
                            numberpoints = numberpoints + 2
                            If IsNumeric(str_data(numberpoints - 2)) And IsNumeric(str_data(numberpoints - 1)) Then
                                HeatReleaseData(1, NumberDataPoints(n) + 1, n) = CDbl(str_data(numberpoints - 2)) 'time
                                HeatReleaseData(2, NumberDataPoints(n) + 1, n) = CDbl(str_data(numberpoints - 1)) 'hrr
                            End If
                            s = s + 1
                            NumberDataPoints(n) = CShort(NumberDataPoints(n) + 1)
                        Loop
                    End If
                End If

                'Call SampleFireData2_LHS(oitemdistributions, oItems, thisitem,)
                If mc_item_hoc IsNot Nothing Then
                    EnergyYield(n) = mc_item_hoc(thisitem - 1, itcounter - 1)
                    SootYield(n) = mc_item_soot(thisitem - 1, itcounter - 1)
                    CO2Yield(n) = mc_item_co2(thisitem - 1, itcounter - 1)
                    ObjectLHoG(n) = mc_item_lhog(thisitem - 1, itcounter - 1)
                    ObjectRLF(n) = mc_item_RLF(thisitem - 1, itcounter - 1)
                    ObjectMLUA(2, n) = mc_item_hrrua(thisitem - 1, itcounter - 1)
                    'HCNuserYield(n) = 0
                    'fireload_count = oItems(thisitem - 1).mass * EnergyYield(n) + fireload_count 'MJ   
                    If usepowerlawdesignfire = False Then
                        fireload_count = oItems(thisitem - 1).mass * EnergyYield(n) + fireload_count 'MJ   
                    Else
                        fireload_count = FLED * x_max * y_max + fireload_count 'MJ   
                    End If
                Else
                    For Each oDistribution In oitemdistributions
                        If oDistribution.varname = "heat of combustion" And oDistribution.id = thisitem Then
                            EnergyYield(n) = oDistribution.varvalue
                            If usepowerlawdesignfire = False Then
                                fireload_count = oItems(thisitem - 1).mass * oDistribution.varvalue + fireload_count 'MJ   
                            Else
                                fireload_count = FLED * x_max * y_max + fireload_count 'MJ   
                            End If
                            'Exit For
                        End If
                        If oDistribution.varname = "soot yield" And oDistribution.id = thisitem Then
                            SootYield(n) = oDistribution.varvalue
                        End If
                        If oDistribution.varname = "co2 yield" And oDistribution.id = thisitem Then
                            CO2Yield(n) = oDistribution.varvalue
                        End If
                        If oDistribution.varname = "Latent Heat of Gasification" And oDistribution.id = thisitem Then
                            ObjectLHoG(n) = oDistribution.varvalue
                        End If
                        If oDistribution.varname = "Radiant Loss Fraction" And oDistribution.id = thisitem Then
                            ObjectRLF(n) = oDistribution.varvalue
                        End If
                        If oDistribution.varname = "HRRUA" And oDistribution.id = thisitem Then
                            ObjectMLUA(2, n) = oDistribution.varvalue
                        End If
                    Next
                End If

                If usepowerlawdesignfire = False Then
                    firemass_count = oItems(thisitem - 1).mass + firemass_count
                Else
                    firemass_count = fireload_count / EnergyYield(1)
                End If
                items_count = items_count + 1
            Loop

            n_max = n

            'calculate length of midpoint vectors from source to target items
            If n_max > 1 Then
                ReDim vectorlength(0 To n_max, 0 To n_max) 'target, source
                Dim x1, x2, y1, y2 As Single
                Dim s As Integer

                For i = 1 To n_max  'i is the target item
                    For s = 1 To n_max
                        x1 = item_location(1, s) 'source item
                        y1 = item_location(2, s) 'source item
                        x2 = item_location(1, i)
                        y2 = item_location(2, i)
                        vectorlength(i, s) = Sqrt((x2 - x1) ^ 2 + (y2 - y1) ^ 2)

                    Next 'source
                    If i > 1 Then ObjectIgnTime(i) = SimTime

                Next 'target

            End If

            If FlagSimStop = True Then

                frmPopulate.txtNumber_Items.Text = n_max.ToString 'display number of items
                frmPopulate.txtFLED.Text = FLED.ToString 'display FLED
                frmPopulate.txtFLED.Text = fireload_count / (x_max * y_max)

            End If
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in DFG.vb populate_items_manual")
        End Try
    End Sub

    Public Sub secondary_targets(ByVal rcnone As Boolean, ByVal thistimestep As Integer, ByVal thisiteration As Integer, ByRef ItemFTP_sum_pilot() As Double, ByRef ItemFTP_sum_auto() As Double, ByRef ItemFTP_sum_wall() As Double, ByRef ItemFTP_sum_ceiling() As Double)
        'code to determine the ignition time of secondary target items
        'runs once per time step
        Try

            Dim r As Double, Q As Double, QPS As Double, zs As Single, radfromlayer As Double
            Dim emissivity As Double, Lf As Double

            'rate of heat release of source at current time
            Q = Get_HRR(1, tim(thistimestep, 1), 0, Qburner, 0, 0, 0) 'kW/m2
            If rcnone = True Then
                'this is not a flame spread simulation
                Qburner = Q
            End If

            If NumberObjects > 1 Then
                For i = 2 To NumberObjects 'a potential target

                    'only do this if object has not previously ignited
                    If ObjectIgnTime(i) = SimTime Then

                        'radial distance from source to be used in point source calculation
                        r = targetdistance(1, i)

                        'incident heat flux from point source to nearest vertical surface on target
                        'QPS = RadiantLossFraction * Q / (4 * PI * r ^ 2) 'kW/m2
                        QPS = ObjectRLF(1) * Qburner / (4 * PI * r ^ 2) 'kW/m2

                        'add heat flux from all other burning items
                        For j = 2 To NumberObjects
                            If i <> j Then
                                If ObjectIgnTime(j) < SimTime Then

                                    r = targetdistance(j, i)

                                    Q = Get_HRR(j, tim(thistimestep, 1), 0, 0, 0, 0, 0) 'kW/m2

                                    'incident heat flux from point source to nearest vertical surface on target
                                    QPS = QPS + ObjectRLF(j) * Q / (4 * PI * r ^ 2) 'kW/m2

                                End If
                            End If
                        Next

                        'height of top surface of secondary item relative to floor
                        zs = ObjElevation(i) + ObjHeight(i)

                        'radiation from underside of hot layer to top surface of secondary object
                        'radfromlayer
                        Call Radiation_to_Surface(fireroom, layerheight(fireroom, thistimestep), uppertemp(fireroom, thistimestep), CO2MassFraction(fireroom, thistimestep, 1), H2OMassFraction(fireroom, thistimestep, 1), radfromlayer, UpperVolume(fireroom, thistimestep), OD_upper(fireroom, thistimestep - 1), zs, emissivity)

                        'point source piloted ignition
                        If QPS >= ObjCRF(i) Then 'vertical surface on target
                            ItemFTP_sum_pilot(i) = ItemFTP_sum_pilot(i) + (QPS - ObjCRF(i)) ^ ObjFTPindexpilot(i) * Timestep
                        End If
                        ObjectRad(0, i, stepcount) = QPS 'vert surface =0 horizont surface=1

                        'hot layer auto ignition
                        If radfromlayer >= ObjCRFauto(i) Then 'horizontal surface on target
                            ItemFTP_sum_auto(i) = ItemFTP_sum_auto(i) + (radfromlayer - ObjCRFauto(i)) ^ ObjFTPindexauto(i) * Timestep
                        End If
                        ObjectRad(1, i, stepcount) = radfromlayer 'vert surface =0 horizont surface=1

                        'ignition time of secondary objects
                        If (ItemFTP_sum_pilot(i) > ObjFTPlimitpilot(i)) And ObjectIgnMode(i) = "" Then
                            ObjectIgnTime(i) = tim(thistimestep, 1)
                            frmInputs.rtb_log.Text = ObjectIgnTime(i).ToString & " sec. Item " & ObjectItemID(i).ToString & " " & ObjectDescription(i) & " ignited (vertical surface)." & Chr(13) & frmInputs.rtb_log.Text
                            ObjectIgnMode(i) = "P"
                        End If

                        'ignition time of secondary objects
                        If (ItemFTP_sum_auto(i) > ObjFTPlimitauto(i)) And ObjectIgnMode(i) = "" Then
                            ObjectIgnTime(i) = tim(thistimestep, 1)
                            frmInputs.rtb_log.Text = ObjectIgnTime(i).ToString & " sec. Item " & ObjectItemID(i).ToString & " " & ObjectDescription(i) & " ignited (horizontal surface)." & Chr(13) & frmInputs.rtb_log.Text
                            ObjectIgnMode(i) = "A"
                        End If

                    End If
                Next
            End If

            If rcnone = False Then 'room lining fire growth model applies 
                'Dim id As Integer = 1 'first item ignited
                'flame spread model, wall ignition
                For id = 2 To NumberObjects
                    If ObjectIgnTime(id) <= tim(thistimestep, 1) Then

                        If WallIgniteFlag(fireroom) = 0 Then

                            'rate of heat release of source at current time
                            Q = Get_HRR(id, tim(thistimestep, 1), Qburner, 0, 0, 0, 0) 'kW/m2

                            If FireLocation(id) = 0 Then
                                'only do this if wall has not previously ignited and it is not in corner or wall

                                'incident heat flux from point source to nearest vertical surface on target
                                QPS = ObjectRLF(id) * Q / (4 * PI * itemtowalldistance(id) ^ 2) 'kW/m2 first item ignited

                            ElseIf FireLocation(id) = 1 Then 'wall
                                ' against wall
                                'assumes wall corner surfaces are in contact with burner flame
                                QPS = 200 * (1 - Exp(-0.09 * Q ^ (1 / 3)))

                            ElseIf FireLocation(id) = 2 Then 'corner
                                'sfpe 3rd ed 2-272
                                QPS = 120 * (1 - Exp(-4 * 0.5 * (ObjLength(id) + ObjWidth(id)))) '=60kw/m2 for 0.17 m burner
                            End If

                            QPS = QPS - QUpperWall(fireroom, thistimestep) 'add the room re-radiation effects 

                            If QPS >= WallQCrit(fireroom) Then
                                ItemFTP_sum_wall(id) = ItemFTP_sum_wall(id) + (QPS - WallQCrit(fireroom)) ^ Walln(fireroom) * Timestep
                            End If

                            'ignition time of wall by item
                            If (ItemFTP_sum_wall(id) > WallFTP(fireroom)) Then
                                WallIgniteTime(fireroom) = tim(thistimestep, 1)
                                WallIgniteFlag(fireroom) = 1
                                WallIgniteStep(fireroom) = thistimestep
                                WallIgniteObject(fireroom) = id

                                frmInputs.rtb_log.Text = WallIgniteTime(fireroom).ToString & " sec. Item " & ObjectItemID(id).ToString & " " & ObjectDescription(id) & " ignites wall." & Chr(13) & frmInputs.rtb_log.Text

                                If burner_id = 0 Then burner_id = id
                                BurnerWidth = (ObjWidth(id) + ObjLength(id)) / 2
                                BurnerFlameLength = Get_BurnerFlameLength(id, Q)

                                Y_pyrolysis(fireroom, thistimestep) = 0.4 * BurnerFlameLength + FireHeight(id)
                                X_pyrolysis(fireroom, thistimestep) = BurnerWidth
                                Y_burnout(fireroom, thistimestep) = FireHeight(id)

                            End If
                        End If

                        If CeilingIgniteFlag(fireroom) = 0 Then
                            'rate of heat release of source at current time
                            Q = Get_HRR(id, tim(thistimestep, 1), Qburner, 0, 0, 0, 0) 'kW/m2

                            If burner_id = 0 Then burner_id = id
                            Dim bw As Single = (ObjLength(id) + ObjWidth(id)) / 2

                            If FireLocation(id) = 0 Then
                                'burner in centre of room
                                QPS = 0.28 * Q ^ (5 / 6) * (RoomHeight(fireroom) - FireHeight(id)) ^ (-7 / 3)
                                'eqn 15b Fire Technology Vol 21 No 4 p267 - mostly convective

                            ElseIf FireLocation(id) = 1 Then 'wall
                                Lf = 0.23 * Q ^ (2 / 5) - 1.02 * bw

                                If Lf < RoomHeight(fireroom) - FireHeight(id) Then
                                    'flame not touching the ceiling
                                    If Lf > 0 Then QPS = 20 * ((RoomHeight(fireroom) - FireHeight(id)) / Lf) ^ (-5 / 3)
                                Else
                                    'flame is touching the ceiling
                                    QPS = 200 * (1 - Exp(-0.09 * Q ^ (1 / 3)))   'kW/m2
                                End If

                            ElseIf FireLocation(id) = 2 Then 'corner

                                Lf = bw * 5.9 * Sqrt(Qburner / 1110 / bw ^ (5 / 2))

                                If (RoomHeight(fireroom) - FireHeight(id)) / Lf <= 0.52 Then
                                    QPS = 120
                                Else
                                    QPS = 13 * ((RoomHeight(fireroom) - FireHeight(id)) / Lf) ^ (-3.5)
                                End If
                            End If

                            QPS = QPS - QCeiling(fireroom, thistimestep) 'add the room re-radiation effects 

                            If QPS >= CeilingQCrit(fireroom) Then
                                ItemFTP_sum_ceiling(id) = ItemFTP_sum_ceiling(id) + (QPS - CeilingQCrit(fireroom)) ^ Ceilingn(fireroom) * Timestep
                            End If

                            'ignition time of ceiling by item
                            If (ItemFTP_sum_ceiling(id) > CeilingFTP(fireroom)) Then
                                CeilingIgniteTime(fireroom) = tim(thistimestep, 1)
                                CeilingIgniteFlag(fireroom) = 1
                                CeilingIgniteStep(fireroom) = thistimestep
                                CeilingIgniteObject(fireroom) = id

                                frmInputs.rtb_log.Text = CeilingIgniteTime(fireroom).ToString & " sec. Item " & ObjectItemID(id).ToString & " " & ObjectDescription(id) & " ignites ceiling." & Chr(13) & frmInputs.rtb_log.Text

                                If RoomHeight(fireroom) + 2 * BurnerWidth >= Y_pyrolysis(fireroom, thistimestep) Then
                                    Y_pyrolysis(fireroom, thistimestep) = RoomHeight(fireroom) + 2 * bw
                                End If
                                If FireHeight(id) > Y_burnout(fireroom, thistimestep) Then
                                    Y_burnout(fireroom, thistimestep) = FireHeight(id)
                                End If
                                If FireLocation(id) = 0 Then
                                    Y_burnout(fireroom, id) = RoomHeight(fireroom)
                                End If

                                'Xstart(fireroom, 1) = Y_pyrolysis(fireroom, i)
                                'Xstart(fireroom, 3) = Y_burnout(fireroom, i)

                            End If
                        End If

                    End If
                Next
            End If

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in DFG.vb secondary_targets")
        End Try
    End Sub
    Public Sub Target_distance()
        'run once before simulations to save target distances in an array
        Dim i, j, k, l As Integer
        Dim r, x, y As Single
        Dim x_max As Single = CSng(RoomLength(fireroom)) 'room x-dimension
        Dim y_max As Single = CSng(RoomWidth(fireroom))  'room y-dimension

        'NumberObjects = 2

        ReDim vectorlength(0 To NumberObjects, 0 To NumberObjects) 'target, source
        ReDim targetdistance(0 To NumberObjects, 0 To NumberObjects) 'source, target
        ReDim targetxpoint(0 To NumberObjects, 0 To NumberObjects) 'source, target
        ReDim targetypoint(0 To NumberObjects, 0 To NumberObjects) 'source, target
        ReDim itemtowalldistance(0 To NumberObjects) 'source

        Try

            For k = 1 To NumberObjects 'source
                Dim sourceitem As Integer = item_location(8, k)
                
                For j = 2 To NumberObjects 'target
                    If k <> j Then

                        Dim rmin As Single = 1000
                        Dim xsav As Single = 0
                        Dim ysav As Single = 0

                        Dim thisitem As Integer = item_location(8, j)

                        'If thisitem = 4 And sourceitem = 8 Then Stop 'mdf4 ignites tv
                        'If thisitem = 8 And sourceitem = 4 Then Stop 'tv ignites mdf4


                        Dim x1 As Single = item_location(3, j)
                        Dim x2 As Single = item_location(4, j)
                        Dim y1 As Single = item_location(5, j)
                        Dim y2 As Single = item_location(6, j)

                        '4 corners of target
                        x1 = Round((x1) * 1000) * 0.001
                        x2 = Round((x2) * 1000) * 0.001
                        y1 = Round((y1) * 1000) * 0.001
                        y2 = Round((y2) * 1000) * 0.001

                        Dim deltax As Single = Abs((x2 - x1) / 10)
                        Dim deltay As Single = Abs((y2 - y1) / 10)

                        'divide each side of the target object into 10ths, and find which point is closest to the centre of the source object, save it
                        x = x1
                        For i = 1 To 11
                            y = y1
                            For l = 1 To 11
                                If x = x1 Or x = x2 Or y = y1 Or y = y2 Then
                                    r = Sqrt((x - item_location(1, k)) ^ 2 + (y - item_location(2, k)) ^ 2)
                                    'r = Sqrt((x - 1) ^ 2 + (y - 1) ^ 2)
                                    If r < rmin Then
                                        rmin = r
                                        xsav = x
                                        ysav = y
                                    End If


                                End If
                                y = Round((y + deltay) * 1000) * 0.001
                            Next
                            x = Round((x + deltax) * 1000) * 0.001
                        Next
                        targetdistance(k, j) = rmin
                        targetxpoint(k, j) = xsav
                        targetypoint(k, j) = ysav

                        vectorlength(j, k) = Sqrt((item_location(1, j) - item_location(1, k)) ^ 2 + (item_location(2, j) - item_location(2, k)) ^ 2)
                        'For x = x1 To x2 Step deltax
                        '    'x = Round(x * 1000) * 0.001 'unresolved issues with rounding in for loop step counter ????

                        '    For y = y1 To y2 Step deltay
                        '        'y = Round(y * 1000) * 0.001 'unresolved issues with rounding in for loop step counter ????

                        '        If x = x1 Or x = x2 Or y = y1 Or y = y2 Then
                        '            r = Sqrt((x - item_location(1, j)) ^ 2 + (y - item_location(2, j)) ^ 2)
                        '            If r < rmin Then rmin = r
                        '        End If
                        '    Next y
                        'Next x

                        'targetdistance(k, j) = r
                    End If
                Next j

                'find distance to nearest wall to use with flame spread model
                Dim sx1 As Single = item_location(3, k)
                Dim sx2 As Single = item_location(4, k)
                Dim sy1 As Single = item_location(5, k)
                Dim sy2 As Single = item_location(6, k)

                'find the least distance from any wall to the centre of the fuel item
                itemtowalldistance(k) = Min(sx1 + (sx2 - sx1) / 2, Min(sy1 + (sy2 - sy1) / 2, Min(x_max - sx2 + (sx2 - sx1) / 2, y_max - sy2 + (sy2 - sy1) / 2)))
                itemtowalldistance(k) = Round((itemtowalldistance(k)) * 1000) * 0.001 'get rid of excess digits

            Next k

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in DFG.vb Target_distance")
        End Try
    End Sub
    Public Sub SampleFireData2_LHS(ByVal oitemdistributions, ByVal oitems, ByVal thisitem, ByVal nx, ByVal thisiteration)

        'thisitem id in unique object list
        'nx this item id in room


        'sample data prior to runs using centerspace nmath.core library
        Dim counter As Integer
        Dim i As Integer
        Dim X As Integer
        Dim N As Integer = CInt(frmInputs.txtNumberIterations.Text)
        If N = 1 Then GoTo skipline
        Dim Rand1N As New RandGenUniform(1, N) 'uniform distribution to generate random numbers between 1 and N - use Ru.next to generate the next number

skipline:

        Dim cdf_upper As Double
        Dim cdf_lower As Double
        Dim Ru As New RandGenUniform(0, 1, seed) 'uniform distribution to generate random numbers between 0 and 1 - use Ru.next to generate the next number
        Dim iTaken(0 To N) As Integer 'array to store 0 or 1 indicating whether an iteration has already been sampled = 1
        Dim mc_variable(0 To N - 1) As Double 'temporary array storing a monte carlo variable at each iteration 
        ReDim tempdatain(0 To N - 1)

        ReDim Preserve mc_energyyield(0 To N - 1, 0 To nx - 1)
        ReDim Preserve mc_sootyield(0 To N - 1, 0 To nx - 1)
        ReDim Preserve mc_co2yield(0 To N - 1, 0 To nx - 1)

        Try

            For i = 1 To N
                tempdatain(i - 1) = (1 / N) * Ru.Next + (i - 1) / N
            Next i

            For Each odistribution In oitemdistributions
                If odistribution.id = thisitem Then
                    If odistribution.varname = "heat of combustion" Then

                        If odistribution.distribution = "Normal" Then
                            'define a normal distribution
                            Dim Distribution1 As New CenterSpace.NMath.Stats.NormalDistribution(CDbl(odistribution.mean), CDbl(odistribution.variance))
                            'truncated normal, check to see if any tempdatain values are out of bounds
                            cdf_upper = Distribution1.CDF(odistribution.ubound)
                            cdf_lower = Distribution1.CDF(odistribution.lbound)
                            For i = 0 To N - 1
                                If tempdatain(i) < cdf_lower Then tempdatain(i) = cdf_lower 'set lower bound
                                If tempdatain(i) > cdf_upper Then tempdatain(i) = cdf_upper 'set upper bound
                            Next
                            For counter = 0 To N - 1
                                EnergyYield(nx) = Distribution1.InverseCDF(tempdatain(counter))
                                mc_variable(counter) = EnergyYield(nx)
                            Next
                        ElseIf odistribution.distribution = "Uniform" Then
                            Dim Distribution1 As New CenterSpace.NMath.Stats.UniformDistribution(CDbl(odistribution.lbound), CDbl(odistribution.ubound))

                            For counter = 0 To N - 1
                                EnergyYield(nx) = Distribution1.InverseCDF(tempdatain(counter))
                                mc_variable(counter) = EnergyYield(nx)
                            Next
                        ElseIf odistribution.distribution = "Triangular" Then
                            Dim Distribution1 As New CenterSpace.NMath.Stats.TriangularDistribution(CDbl(odistribution.lbound), CDbl(odistribution.ubound), CDbl(odistribution.mode))

                            For counter = 0 To N - 1
                                EnergyYield(nx) = Distribution1.InverseCDF(tempdatain(counter))
                                mc_variable(counter) = EnergyYield(nx)
                            Next
                        ElseIf odistribution.distribution = "Log Normal" Then
                            Dim Distribution1 As New CenterSpace.NMath.Stats.LognormalDistribution(CDbl(odistribution.mean), Math.Sqrt(CDbl(odistribution.variance)))

                            For counter = 0 To N - 1
                                EnergyYield(nx) = Distribution1.InverseCDF(tempdatain(counter))
                                mc_variable(counter) = EnergyYield(nx)
                            Next
                        ElseIf odistribution.distribution = "Gamma" Then
                            'Dim Distribution1 As New CenterSpace.NMath.Stats.GammaDistribution

                        Else
                            EnergyYield(nx) = CSng(odistribution.varvalue)
                            For counter = 0 To N - 1
                                mc_variable(counter) = EnergyYield(nx)
                            Next
                        End If

                    End If
                End If
            Next

            'need to random sort the vector for independent variables
            'ReDim iTaken(0 To N)
            'i = 1
            'While i <= N
            '    X = Math.Round(N * Ru.Next) 'randomly select an iteration number between 1 and N

            '    If iTaken(X) = 0 And X < N Then
            '        mc_energyyield(nx - 1, i - 1) = mc_variable(X)
            '        iTaken(X) = 1
            '        i = i + 1
            '    End If
            'End While

            ReDim iTaken(0 To N)
            i = 1
            While i <= N
                If N = 1 Then
                    X = 1
                Else
                    X = Rand1N.Next 'randomly select an iteration number between 1 and N
                End If
                If iTaken(X) = 0 And X <= N Then
                    mc_energyyield(i - 1, nx - 1) = mc_variable(X - 1)
                    iTaken(X) = 1
                    i = i + 1
                End If
            End While

            'soot yield
            For i = 1 To N
                tempdatain(i - 1) = (1 / N) * Ru.Next + (i - 1) / N
            Next i

            For Each odistribution In oitemdistributions
                If odistribution.id = thisitem Then
                    If odistribution.varname = "soot yield" Then
                        If odistribution.distribution = "Normal" Then
                            'define a normal distribution
                            Dim Distribution2 As New CenterSpace.NMath.Stats.NormalDistribution(CDbl(odistribution.mean), CDbl(odistribution.variance))
                            'truncated normal, check to see if any tempdatain values are out of bounds
                            cdf_upper = Distribution2.CDF(odistribution.ubound)
                            cdf_lower = Distribution2.CDF(odistribution.lbound)
                            For i = 0 To N - 1
                                If tempdatain(i) < cdf_lower Then tempdatain(i) = cdf_lower 'set lower bound
                                If tempdatain(i) > cdf_upper Then tempdatain(i) = cdf_upper 'set upper bound
                            Next
                            For counter = 0 To N - 1
                                SootYield(nx) = Distribution2.InverseCDF(tempdatain(counter))
                                mc_variable(counter) = SootYield(nx)
                            Next
                        ElseIf odistribution.distribution = "Uniform" Then
                            Dim Distribution2 As New CenterSpace.NMath.Stats.UniformDistribution(CDbl(odistribution.lbound), CDbl(odistribution.ubound))

                            For counter = 0 To N - 1
                                SootYield(nx) = Distribution2.InverseCDF(tempdatain(counter))
                                mc_variable(counter) = SootYield(nx)
                            Next
                        ElseIf odistribution.distribution = "Triangular" Then
                            Dim Distribution2 As New CenterSpace.NMath.Stats.TriangularDistribution(CDbl(odistribution.lbound), CDbl(odistribution.ubound), CDbl(odistribution.mode))

                            For counter = 0 To N - 1
                                SootYield(nx) = Distribution2.InverseCDF(tempdatain(counter))
                                mc_variable(counter) = SootYield(nx)
                            Next
                        ElseIf odistribution.distribution = "Log Normal" Then
                            Dim Distribution2 As New CenterSpace.NMath.Stats.LognormalDistribution(CDbl(odistribution.mean), Math.Sqrt(CDbl(odistribution.variance)))

                            For counter = 0 To N - 1
                                SootYield(nx) = Distribution2.InverseCDF(tempdatain(counter))
                                mc_variable(counter) = SootYield(nx)
                            Next
                        Else
                            SootYield(nx) = CSng(odistribution.varvalue)
                            For counter = 0 To N - 1
                                mc_variable(counter) = SootYield(nx)
                            Next
                        End If

                    End If
                End If
            Next
            'need to random sort the vector for independent variables
            'ReDim iTaken(0 To N)
            'i = 1
            'While i <= N
            '    X = Math.Round(N * Ru.Next) 'randomly select an iteration number between 1 and N

            '    If iTaken(X) = 0 And X < N Then
            '        mc_sootyield(nx - 1, i - 1) = mc_variable(X)
            '        iTaken(X) = 1
            '        i = i + 1
            '    End If
            'End While
            ReDim iTaken(0 To N)
            i = 1
            While i <= N
                If N = 1 Then
                    X = 1
                Else
                    X = Rand1N.Next 'randomly select an iteration number between 1 and N
                End If
                If iTaken(X) = 0 And X <= N Then
                    mc_sootyield(i - 1, nx - 1) = mc_variable(X - 1)
                    iTaken(X) = 1
                    i = i + 1
                End If
            End While

            'co2 yield
            For i = 1 To N
                tempdatain(i - 1) = (1 / N) * Ru.Next + (i - 1) / N
            Next i

            For Each odistribution In oitemdistributions
                If odistribution.id = thisitem Then
                    If odistribution.varname = "co2 yield" Then
                        If odistribution.distribution = "Normal" Then
                            'define a normal distribution
                            Dim Distribution3 As New CenterSpace.NMath.Stats.NormalDistribution(CDbl(odistribution.mean), CDbl(odistribution.variance))
                            'truncated normal, check to see if any tempdatain values are out of bounds
                            cdf_upper = Distribution3.CDF(odistribution.ubound)
                            cdf_lower = Distribution3.CDF(odistribution.lbound)
                            For i = 0 To N - 1
                                If tempdatain(i) < cdf_lower Then tempdatain(i) = cdf_lower 'set lower bound
                                If tempdatain(i) > cdf_upper Then tempdatain(i) = cdf_upper 'set upper bound
                            Next
                            For counter = 0 To N - 1
                                CO2Yield(nx) = Distribution3.InverseCDF(tempdatain(counter))
                                mc_variable(counter) = CO2Yield(nx)
                            Next
                        ElseIf odistribution.distribution = "Uniform" Then
                            Dim Distribution3 As New CenterSpace.NMath.Stats.UniformDistribution(CDbl(odistribution.lbound), CDbl(odistribution.ubound))

                            For counter = 0 To N - 1
                                CO2Yield(nx) = Distribution3.InverseCDF(tempdatain(counter))
                                mc_variable(counter) = CO2Yield(nx)
                            Next
                        ElseIf odistribution.distribution = "Triangular" Then
                            Dim Distribution3 As New CenterSpace.NMath.Stats.TriangularDistribution(CDbl(odistribution.lbound), CDbl(odistribution.ubound), CDbl(odistribution.mode))

                            For counter = 0 To N - 1
                                CO2Yield(nx) = Distribution3.InverseCDF(tempdatain(counter))
                                mc_variable(counter) = CO2Yield(nx)
                            Next
                        ElseIf odistribution.distribution = "Log Normal" Then
                            Dim Distribution3 As New CenterSpace.NMath.Stats.LognormalDistribution(CDbl(odistribution.mean), Math.Sqrt(CDbl(odistribution.variance)))

                            For counter = 0 To N - 1
                                CO2Yield(nx) = Distribution3.InverseCDF(tempdatain(counter))
                                mc_variable(counter) = CO2Yield(nx)
                            Next
                        Else
                            CO2Yield(nx) = CSng(odistribution.varvalue)
                            For counter = 0 To N - 1
                                mc_variable(counter) = CO2Yield(nx)
                            Next
                        End If
                    End If
                End If
            Next

            ReDim iTaken(0 To N)
            i = 1
            While i <= N
                If N = 1 Then
                    X = 1
                Else
                    X = Rand1N.Next 'randomly select an iteration number between 1 and N
                End If
                If iTaken(X) = 0 And X <= N Then
                    mc_co2yield(i - 1, nx - 1) = mc_variable(X - 1)
                    iTaken(X) = 1
                    i = i + 1
                End If
            End While


            ''FTP limit pilot
            'For i = 1 To N
            '    tempdatain(i - 1) = (1 / N) * Ru.Next + (i - 1) / N
            'Next i

            'If oitems(thisitem - 1).ftplimitpilotdistribution = "Normal" Then
            '    'define a normal distribution
            '    Dim Distribution4 As New CenterSpace.NMath.Stats.NormalDistribution(CDbl(oitems(thisitem - 1).ftplimitpilotmean), CDbl(oitems(thisitem - 1).ftplimitpilotvariance))
            '    'truncated normal, check to see if any tempdatain values are out of bounds
            '    cdf_upper = Distribution4.CDF(oitems(thisitem - 1).ftplimitpilotubound)
            '    cdf_lower = Distribution4.CDF(oitems(thisitem - 1).ftplimitpilotlbound)
            '    For i = 0 To N - 1
            '        If tempdatain(i) < cdf_lower Then tempdatain(i) = cdf_lower 'set lower bound
            '        If tempdatain(i) > cdf_upper Then tempdatain(i) = cdf_upper 'set upper bound
            '    Next
            '    For counter = 0 To N - 1
            '        ObjFTPlimitpilot(nx) = Distribution4.InverseCDF(tempdatain(counter))
            '        mc_variable(counter) = ObjFTPlimitpilot(nx)
            '    Next
            'ElseIf oitems(thisitem - 1).FTPlimitpilotdistribution = "Uniform" Then
            '    Dim Distribution4 As New CenterSpace.NMath.Stats.UniformDistribution(CDbl(oitems(thisitem - 1).FTPlimitpilotlbound), CDbl(oitems(thisitem - 1).FTPlimitpilotubound))

            '    For counter = 0 To N - 1
            '        ObjFTPlimitpilot(nx) = Distribution4.InverseCDF(tempdatain(counter))
            '        mc_variable(counter) = ObjFTPlimitpilot(nx)
            '    Next
            'ElseIf oitems(thisitem - 1).FTPlimitpilotdistribution = "Triangular" Then
            '    Dim Distribution4 As New CenterSpace.NMath.Stats.TriangularDistribution(CDbl(oitems(thisitem - 1).FTPlimitpilotlbound), CDbl(oitems(thisitem - 1).FTPlimitpilotubound), CDbl(oitems(thisitem - 1).FTPlimitpilotmean))

            '    For counter = 0 To N - 1
            '        ObjFTPlimitpilot(nx) = Distribution4.InverseCDF(tempdatain(counter))
            '        mc_variable(counter) = ObjFTPlimitpilot(nx)
            '    Next
            'ElseIf oitems(thisitem - 1).FTPlimitpilotdistribution = "Log Normal" Then
            '    Dim Distribution4 As New CenterSpace.NMath.Stats.LognormalDistribution(CDbl(oitems(thisitem - 1).FTPlimitpilotmean), Math.Sqrt(CDbl(oitems(thisitem - 1).FTPlimitpilotvariance)))

            '    For counter = 0 To N - 1
            '        ObjFTPlimitpilot(nx) = Distribution4.InverseCDF(tempdatain(counter))
            '        mc_variable(counter) = ObjFTPlimitpilot(nx)
            '    Next
            'Else
            '    ObjFTPlimitpilot(nx) = CSng(oitems(thisitem - 1).FTPlimitpilot)
            '    For counter = 0 To N - 1
            '        mc_variable(counter) = ObjFTPlimitpilot(nx)
            '    Next
            'End If

            'ReDim iTaken(0 To N)
            'i = 1
            'While i <= N
            '    If N = 1 Then
            '        X = 1
            '    Else
            '        X = Rand1N.Next 'randomly select an iteration number between 1 and N
            '    End If
            '    If iTaken(X) = 0 And X <= N Then
            '        mc_FTPlimitpilot(i - 1, nx - 1) = mc_variable(X - 1)
            '        iTaken(X) = 1
            '        i = i + 1
            '    End If
            'End While

            ''FTP index pilot
            'For i = 1 To N
            '    tempdatain(i - 1) = (1 / N) * Ru.Next + (i - 1) / N
            'Next i

            'If oitems(thisitem - 1).ftpindexpilotdistribution = "Normal" Then
            '    'define a normal distribution
            '    Dim Distribution5 As New CenterSpace.NMath.Stats.NormalDistribution(CDbl(oitems(thisitem - 1).ftpindexpilotmean), CDbl(oitems(thisitem - 1).ftpindexpilotvariance))
            '    'truncated normal, check to see if any tempdatain values are out of bounds
            '    cdf_upper = Distribution5.CDF(oitems(thisitem - 1).ftpindexpilotubound)
            '    cdf_lower = Distribution5.CDF(oitems(thisitem - 1).ftpindexpilotlbound)
            '    For i = 0 To N - 1
            '        If tempdatain(i) < cdf_lower Then tempdatain(i) = cdf_lower 'set lower bound
            '        If tempdatain(i) > cdf_upper Then tempdatain(i) = cdf_upper 'set upper bound
            '    Next
            '    For counter = 0 To N - 1
            '        ObjFTPindexpilot(nx) = Distribution5.InverseCDF(tempdatain(counter))
            '        mc_variable(counter) = ObjFTPindexpilot(nx)
            '    Next
            'ElseIf oitems(thisitem - 1).ftpindexpilotdistribution = "Uniform" Then
            '    Dim Distribution5 As New CenterSpace.NMath.Stats.UniformDistribution(CDbl(oitems(thisitem - 1).ftpindexpilotlbound), CDbl(oitems(thisitem - 1).ftpindexpilotubound))

            '    For counter = 0 To N - 1
            '        ObjFTPindexpilot(nx) = Distribution5.InverseCDF(tempdatain(counter))
            '        mc_variable(counter) = ObjFTPindexpilot(nx)
            '    Next
            'ElseIf oitems(thisitem - 1).ftpindexpilotdistribution = "Triangular" Then
            '    Dim Distribution5 As New CenterSpace.NMath.Stats.TriangularDistribution(CDbl(oitems(thisitem - 1).ftpindexpilotlbound), CDbl(oitems(thisitem - 1).ftpindexpilotubound), CDbl(oitems(thisitem - 1).ftpindexpilotmean))

            '    For counter = 0 To N - 1
            '        ObjFTPindexpilot(nx) = Distribution5.InverseCDF(tempdatain(counter))
            '        mc_variable(counter) = ObjFTPindexpilot(nx)
            '    Next
            'ElseIf oitems(thisitem - 1).ftpindexpilotdistribution = "Log Normal" Then
            '    Dim Distribution5 As New CenterSpace.NMath.Stats.LognormalDistribution(CDbl(oitems(thisitem - 1).ftpindexpilotmean), Math.Sqrt(CDbl(oitems(thisitem - 1).ftpindexpilotvariance)))

            '    For counter = 0 To N - 1
            '        ObjFTPindexpilot(nx) = Distribution5.InverseCDF(tempdatain(counter))
            '        mc_variable(counter) = ObjFTPindexpilot(nx)
            '    Next
            'Else
            '    ObjFTPindexpilot(nx) = CSng(oitems(thisitem - 1).ftpindexpilot)
            '    For counter = 0 To N - 1
            '        mc_variable(counter) = ObjFTPindexpilot(nx)
            '    Next
            'End If

            'ReDim iTaken(0 To N)
            'i = 1
            'While i <= N
            '    If N = 1 Then
            '        X = 1
            '    Else
            '        X = Rand1N.Next 'randomly select an iteration number between 1 and N
            '    End If
            '    If iTaken(X) = 0 And X <= N Then
            '        mc_FTPindexpilot(i - 1, nx - 1) = mc_variable(X - 1)
            '        iTaken(X) = 1
            '        i = i + 1
            '    End If
            'End While

            ''critical flux pilot
            'For i = 1 To N
            '    tempdatain(i - 1) = (1 / N) * Ru.Next + (i - 1) / N
            'Next i

            'If oitems(thisitem - 1).critfluxdistribution = "Normal" Then
            '    'define a normal distribution
            '    Dim Distribution6 As New CenterSpace.NMath.Stats.NormalDistribution(CDbl(oitems(thisitem - 1).critfluxmean), CDbl(oitems(thisitem - 1).critfluxvariance))
            '    'truncated normal, check to see if any tempdatain values are out of bounds
            '    cdf_upper = Distribution6.CDF(oitems(thisitem - 1).critfluxubound)
            '    cdf_lower = Distribution6.CDF(oitems(thisitem - 1).critfluxlbound)
            '    For i = 0 To N - 1
            '        If tempdatain(i) < cdf_lower Then tempdatain(i) = cdf_lower 'set lower bound
            '        If tempdatain(i) > cdf_upper Then tempdatain(i) = cdf_upper 'set upper bound
            '    Next
            '    For counter = 0 To N - 1
            '        ObjCRF(nx) = Distribution6.InverseCDF(tempdatain(counter))
            '        mc_variable(counter) = ObjCRF(nx)
            '    Next
            'ElseIf oitems(thisitem - 1).critfluxdistribution = "Uniform" Then
            '    Dim Distribution6 As New CenterSpace.NMath.Stats.UniformDistribution(CDbl(oitems(thisitem - 1).critfluxlbound), CDbl(oitems(thisitem - 1).critfluxubound))

            '    For counter = 0 To N - 1
            '        ObjCRF(nx) = Distribution6.InverseCDF(tempdatain(counter))
            '        mc_variable(counter) = ObjCRF(nx)
            '    Next
            'ElseIf oitems(thisitem - 1).critfluxdistribution = "Triangular" Then
            '    Dim Distribution6 As New CenterSpace.NMath.Stats.TriangularDistribution(CDbl(oitems(thisitem - 1).critfluxlbound), CDbl(oitems(thisitem - 1).critfluxubound), CDbl(oitems(thisitem - 1).critfluxmean))

            '    For counter = 0 To N - 1
            '        ObjCRF(nx) = Distribution6.InverseCDF(tempdatain(counter))
            '        mc_variable(counter) = ObjCRF(nx)
            '    Next
            'ElseIf oitems(thisitem - 1).critfluxdistribution = "Log Normal" Then
            '    Dim Distribution6 As New CenterSpace.NMath.Stats.LognormalDistribution(CDbl(oitems(thisitem - 1).critfluxmean), Math.Sqrt(CDbl(oitems(thisitem - 1).critfluxvariance)))

            '    For counter = 0 To N - 1
            '        ObjCRF(nx) = Distribution6.InverseCDF(tempdatain(counter))
            '        mc_variable(counter) = ObjCRF(nx)
            '    Next
            'Else
            '    ObjCRF(nx) = CSng(oitems(thisitem - 1).critflux)
            '    For counter = 0 To N - 1
            '        mc_variable(counter) = ObjCRF(nx)
            '    Next
            'End If

            'ReDim iTaken(0 To N)
            'i = 1
            'While i <= N
            '    If N = 1 Then
            '        X = 1
            '    Else
            '        X = Rand1N.Next 'randomly select an iteration number between 1 and N
            '    End If
            '    If iTaken(X) = 0 And X <= N Then
            '        mc_critflux(i - 1, nx - 1) = mc_variable(X - 1)
            '        iTaken(X) = 1
            '        i = i + 1
            '    End If
            'End While

            Exit Sub

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in DFG.vb SampleFireData2_LHS")
        End Try
    End Sub
    Public Sub sample_fireload()
        '
        'fled - extracts a value for FLED from the distribution assigned
        '
        Try
            Dim distname As String = ""
            Dim mean As Single
            Dim variance As Single
            Dim value As Single
            Dim lbound As Single
            Dim ubound As Single
            Dim alpha As Single
            Dim beta As Single
            Dim mode As Single

            Dim odistributions As List(Of oDistribution)
            odistributions = DistributionClass.GetDistributions()

            For Each oDistribution In odistributions
                If oDistribution.varname = "Fire Load Energy Density" Then
                    distname = oDistribution.distribution
                    mean = oDistribution.mean
                    variance = oDistribution.variance
                    value = oDistribution.varvalue
                    lbound = oDistribution.lbound
                    ubound = oDistribution.ubound
                    alpha = oDistribution.alpha
                    beta = oDistribution.beta
                    mode = oDistribution.mode
                End If
            Next

            If distname = "Normal" Then 'fled
                Dim Distribution4 As New RandGenNormal(CDbl(mean), CDbl(variance))
                FLED = Distribution4.Next
            ElseIf distname = "Uniform" Then
                Dim Distribution4 As New RandGenUniform(CDbl(lbound), CDbl(ubound))
                FLED = Distribution4.Next
            ElseIf distname = "Triangular" Then
                Dim Distribution4 As New RandGenTriangular(CDbl(lbound), CDbl(ubound), CDbl(mode))
                FLED = Distribution4.Next
            ElseIf distname = "Log Normal" Then
                Dim Distribution4 As New RandGenLogNormal(CDbl(mean), Math.Sqrt(CDbl(variance)))
                FLED = Distribution4.Next
            Else
                FLED = CSng(value)
            End If
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in DFG.vb sample_fireload")
        End Try
    End Sub
End Module

