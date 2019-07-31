Imports System.Drawing.Graphics
Imports System.Math
Imports CenterSpace.NMath.Core
Imports CenterSpace.NMath.Stats
Imports System.Drawing.Drawing2D
Imports System.Collections.Generic
Imports VB = Microsoft.VisualBasic


Public Class frmPopulate
    Dim oSprinklers As List(Of oSprinkler)
    Dim oSmokeDets As List(Of oSmokeDet)


    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub frmPopulate_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Centre_Form(Me)

        If autopopulate = True Then
            RadioButton1.Checked = True
        Else
            RadioButton2.Checked = True
        End If
        If calc_sprdist = True Then
            frmSprinklerList.chkCalcSprinkRadialDist.Checked = True
        Else
            frmSprinklerList.chkCalcSprinkRadialDist.Checked = False
        End If
        chk_fix_item.Checked = fixitem1

        'run the subroutine "button1_click" repeatedly at the interval given by timer1.interval property
        AddHandler Timer1.Tick, AddressOf Button1_Click


    End Sub
    Public Sub paintme()
        'code to run each time the form is painted, displaying the location of items in the room

        Dim drawroom As Graphics
        Dim drawscale As Graphics
        Dim blackpen As New Pen(Color.Black)
        Dim rwidth As Single
        Dim rlength As Single
        Dim gridpixels As Single
        Dim convert As Single
        Dim a, b, c, d As Single
        Dim str, str1 As String

        Try
            If IsNothing(vent_space) Then ReDim vent_space(0 To NumVents, 0 To 3)

            Dim numFRvents As Integer = 0
            For i = 1 To NumberRooms + 1
                For j = 1 To NumberVents(i, fireroom)
                    numFRvents = numFRvents + 1
                Next
            Next

            'figure out suitable scaling to draw the room on screen
            If RoomLength(fireroom) > RoomWidth(fireroom) Then
                Panel1.Width = 500 'room length
                convert = Panel1.Width / RoomLength(fireroom) 'pixels/metre
                Panel1.Height = convert * RoomWidth(fireroom)  'room width
            Else
                Panel1.Height = 500
                convert = Panel1.Height / RoomWidth(fireroom) 'pixels/metre
                Panel1.Width = convert * RoomLength(fireroom)
            End If

            gridpixels = CSng(gridsize * convert)
            If gridpixels = 0 Then Exit Sub

            'room dimensions to use in pixels
            rlength = CSng(Panel1.Width)
            rwidth = CSng(Panel1.Height)

            drawroom = Panel1.CreateGraphics

            drawroom.Clear(Color.White)

            'draw room outline
            drawroom.DrawRectangle(Pens.Black, 0, 0, rlength, rwidth)

            'draw vent keep clear areas
            Dim p As New Pen(Color.CornflowerBlue)
            Dim myBrush As New HatchBrush(HatchStyle.LightUpwardDiagonal, Color.Red)
            p.DashStyle = DashStyle.Dot

            If CheckBox2.Checked Then 'show grid
                For i = 1 To CInt(rwidth / gridpixels)
                    drawroom.DrawLine(p, 0, CInt(i * gridpixels), rlength, CInt(i * gridpixels))
                Next i
                For i = 1 To CInt(rlength / gridpixels)
                    drawroom.DrawLine(p, CInt(i * gridpixels), 0, CInt(i * gridpixels), rwidth)
                Next i
            End If

            p.Color = Color.Red

            For n = 1 To numFRvents
                If vent_space.GetUpperBound(0) < n Then Exit For
                a = CSng(vent_space(n, 0) * convert)
                b = rwidth - CSng(vent_space(n, 3) * convert)
                c = CSng((vent_space(n, 1) - vent_space(n, 0)) * convert)
                d = CSng((vent_space(n, 3) - vent_space(n, 2)) * convert)

                drawroom.DrawRectangle(p, a, b, c, d)
                drawroom.DrawString("keep clear", New Font("Arial", 8), New SolidBrush(Color.Red), a + 5, b)
            Next

            'draw sprinklers
            If chkShowSprink.Checked Then
                oSprinklers = SprinklerDB.GetSprinklers2
                For Each oSprinkler In oSprinklers
                    a = CSng(oSprinkler.sprx * convert)
                    b = rwidth - CSng(oSprinkler.spry * convert)
                    drawroom.DrawRectangle(Pens.Purple, a, b, 5, 5)
                Next
            End If

            'draw smoke detectors
            If chkShowSD.Checked Then
                oSmokeDets = SmokeDetDB.GetSmokDets
                For Each oSmokeDet In oSmokeDets
                    a = CSng(oSmokeDet.sdx * convert)
                    b = rwidth - CSng(oSmokeDet.sdy * convert)
                    'drawroom.DrawRectangle(Pens.Green, a, b, 5, 5)
                    drawroom.FillRectangle(Brushes.Chartreuse, a, b, 5, 5)
                Next
            End If

            'draw items
            If n_max > NumberObjects Then Exit Sub
            For n = 1 To n_max
                If vectorlength.Length <= n Then Exit For
                'centre to centre distances
                vectorlength(n, 1) = Sqrt((item_location(1, n) - item_location(1, 1)) ^ 2 + (item_location(2, n) - item_location(2, 1)) ^ 2)

                a = CSng(item_location(3, n) * convert)
                c = CSng((item_location(4, n) - item_location(3, n)) * convert)
                d = CSng((item_location(6, n) - item_location(5, n)) * convert)
                b = rwidth - CSng((item_location(6, n) * convert))


                'draw item
                'label item
                If n < itime.Count Then
                    If itime(n) < SimTime Then
                        If n > 1 And itime(n) = 0 Then
                            If optidnum.Checked = True Then
                                str1 = ObjectItemID(n)
                            Else
                                str1 = ObjLabel(n).ToString
                            End If
                        Else
                            If optidnum.Checked = True Then
                                str1 = ObjectItemID(n) & Chr(10) & "" & ignmode(n) & "I " & itime(n) & "s"
                            Else
                                str1 = ObjLabel(n).ToString & Chr(10) & "" & ignmode(n) & "I " & itime(n) & "s"
                            End If
                        End If
                    Else
                        If optidnum.Checked = True Then
                            str1 = n.ToString
                        Else
                            str1 = ObjLabel(n).ToString
                        End If
                    End If
                Else
                    If optidnum.Checked = True Then
                        str1 = n.ToString
                    Else
                        str1 = ObjLabel(n).ToString
                    End If
                End If

                If n = 1 Then
                    'designate item 1 as the item first ignited
                    drawroom.DrawRectangle(Pens.Black, a, b, c, d)
                    drawroom.DrawString(str1, New Font("Arial", 8), New SolidBrush(Color.Black), a, b)
                Else
                    'other items are secondary targets

                    drawroom.DrawRectangle(Pens.Blue, a, b, c, d)
                    drawroom.DrawString(str1, New Font("Arial", 8), New SolidBrush(Color.Blue), a, b)

                    'draw line vectors from first item ignited to secondary items
                    If CheckBox1.Checked = True Then

                        'str = CStr(Format(vectorlength(n, 1), "0.00")) & " m"
                        str = CStr(Format(targetdistance(1, n), "0.000")) & " m"

                        drawroom.DrawString(str, New Font("Arial", 8), New SolidBrush(Color.Black), a + c / 4, b + d / 2)
                        'drawroom.DrawLine(Pens.Black, CSng(item_location(1, 1) * convert), rwidth - CSng(item_location(2, 1) * convert), CSng(item_location(1, n) * convert), rwidth - CSng(item_location(2, n) * convert))

                        drawroom.DrawLine(Pens.Black, CSng(item_location(1, 1) * convert), rwidth - CSng(item_location(2, 1) * convert), CSng(targetxpoint(1, n) * convert), rwidth - CSng(targetypoint(1, n) * convert))

                    End If

                End If
            Next

            drawscale = Panel2.CreateGraphics
            Dim t As New Pen(Color.Black, 1)
            t.EndCap = LineCap.Custom
            t.StartCap = LineCap.Custom
            t.CustomEndCap = New AdjustableArrowCap(5, 5, True)
            t.CustomStartCap = New AdjustableArrowCap(5, 5, True)
            'draw line with arrowhead each end, 1 m length
            drawscale.DrawLine(t, 0, 5, 1 * convert, 5)
            drawscale.DrawString("1 m", Font, Brushes.Black, convert / 3, 10, StringFormat.GenericTypographic)
            drawscale.DrawString("Length", Font, Brushes.Black, convert / 3, 25, StringFormat.GenericTypographic)


        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in frmPopulate Events Paint")
        End Try
    End Sub


    Public Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPopulate.Click

        If usepowerlawdesignfire = True Then
            n_max = 1
            'Exit Sub
        End If

        ReDim Item1X(0 To frmInputs.txtNumberIterations.Text - 1)
        ReDim Item1Y(0 To frmInputs.txtNumberIterations.Text - 1)
        ReDim itime(0 To n_max)
        ReDim ignmode(0 To n_max)

        '-----------------
        'new 2019
        fixitem1 = chk_fix_item.Checked
        Call populate_items_manual(n_max, 1, 0, 0) 'this will manually locate the first item regardless
        If RadioButton1.Checked Then
            Call populate_items(n_max, 1, 0, 0) 'use this auto-populate all items except the first item (if specifed)
        End If
        '-----------------

        'If RadioButton1.Checked Then
        '    Call populate_items(n_max, 1, 0, 0)
        'Else
        '    Call populate_items_manual(n_max, 1, 0, 0)
        'End If

        Call Target_distance()

        Call paintme()


        Me.Show()

    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        If RadioButton1.Checked = True Then
            autopopulate = True
        Else
            autopopulate = False
        End If
        If chk_fix_item.Checked = True Then
            fixitem1 = True
        End If
        Me.Close()
    End Sub

    Private Sub cmdStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStart.Click

        Timer1.Interval = 1500
        Timer1.Start()

    End Sub

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click

        Timer1.Stop()

    End Sub

    Private Sub AltToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        frmItemList.Show()
        frmItemList.BringToFront()
    End Sub

    Private Sub cmdLayout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLayout.Click
        Try

            Dim iname As String, oname As String
            iname = RiskDataDirectory & "input" & updown_itcounter.Value & ".xml"
            oname = RiskDataDirectory & "output" & updown_itcounter.Value & ".xml"

            frmInputs.Read_InputFile_xml(iname)

            Call Target_distance()

            ReDim Preserve itime(0 To n_max)
            ReDim Preserve ignmode(0 To n_max)

            'retrieve item ignition times
            frmInputs.Retrieve_ign_times(oname, itime, ignmode)
            If Not IsNothing(frmInputs.mc_FLED_actual) Then
                txtFLED.Text = frmInputs.mc_FLED_actual(updown_itcounter.Value - 1)
            End If

            txtNumber_Items.Text = NumberObjects.ToString

            Call paintme()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in frmPopulate cmdLayout")
        End Try

    End Sub

    Private Sub updown_itcounter_Scroll(ByVal sender As Object, ByVal e As System.EventArgs) Handles updown_itcounter.Click
        'cmdLayout.PerformClick()
        'Me.Show()
        Try

            Dim iname As String, oname As String
            iname = RiskDataDirectory & "input" & updown_itcounter.Value & ".xml"
            oname = RiskDataDirectory & "output" & updown_itcounter.Value & ".xml"

            frmInputs.Read_InputFile_xml(iname)

            Call Target_distance()

            ReDim Preserve itime(0 To n_max)
            ReDim Preserve ignmode(0 To n_max)

            'retrieve item ignition times
            frmInputs.Retrieve_ign_times(oname, itime, ignmode)
            If Not IsNothing(frmInputs.mc_FLED_actual) Then
                txtFLED.Text = frmInputs.mc_FLED_actual(updown_itcounter.Value - 1)
            End If

            txtNumber_Items.Text = NumberObjects.ToString

            Call paintme()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in frmPopulate updown_itcounter")
        End Try
    End Sub

    Private Sub updown_itcounter_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles updown_itcounter.ValueChanged

        'Call cmdLayout.PerformClick()
        'Me.Show()

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim file As String = "yourfile.jpg"

        Dim blackpen As New Pen(Color.Black)
        Dim rwidth As Single
        Dim rlength As Single
        Dim gridpixels As Single
        Dim convert As Single
        Dim a, b, c, d As Single
        Dim str, str1 As String

        Try

            'figure out suitable scaling to draw the room on screen
            If RoomLength(fireroom) > RoomWidth(fireroom) Then
                Panel1.Width = 500 'room length
                convert = Panel1.Width / RoomLength(fireroom) 'pixels/metre
                Panel1.Height = convert * RoomWidth(fireroom)  'room width
            Else
                Panel1.Height = 500
                convert = Panel1.Height / RoomWidth(fireroom) 'pixels/metre
                Panel1.Width = convert * RoomLength(fireroom)
            End If

            Dim rect As Rectangle = New Rectangle(0, 0, Panel1.Width, Panel1.Height)
            Dim displayGraphics As Graphics = Me.CreateGraphics()
            Dim img As Image = New Bitmap(rect.Width, rect.Height, displayGraphics)
            Dim drawroom As Graphics = Graphics.FromImage(img)
            Dim drawscale As Graphics = Graphics.FromImage(img)

            gridpixels = CSng(gridsize * convert)
            If gridpixels = 0 Then Exit Sub

            'room dimensions to use in pixels
            rlength = CSng(Panel1.Width)
            rwidth = CSng(Panel1.Height)

            drawroom.Clear(Color.White)

            'draw room outline
            drawroom.DrawRectangle(Pens.Black, 0, 0, rlength, rwidth)

            'draw vent keep clear areas
            Dim p As New Pen(Color.CornflowerBlue)
            Dim myBrush As New HatchBrush(HatchStyle.LightUpwardDiagonal, Color.Red)
            p.DashStyle = DashStyle.Dot

            If CheckBox2.Checked Then 'show grid
                For i = 1 To CInt(rwidth / gridpixels)
                    drawroom.DrawLine(p, 0, CInt(i * gridpixels), rlength, CInt(i * gridpixels))
                Next i
                For i = 1 To CInt(rlength / gridpixels)
                    drawroom.DrawLine(p, CInt(i * gridpixels), 0, CInt(i * gridpixels), rwidth)
                Next i
            End If

            p.Color = Color.Red
            For n = 1 To number_vents
                a = CSng(vent_space(n, 0) * convert)
                b = rwidth - CSng(vent_space(n, 3) * convert)
                c = CSng((vent_space(n, 1) - vent_space(n, 0)) * convert)
                d = CSng((vent_space(n, 3) - vent_space(n, 2)) * convert)

                drawroom.DrawRectangle(p, a, b, c, d)
                drawroom.DrawString("keep clear", New Font("Arial", 9), New SolidBrush(Color.Red), a + 5, b)
            Next

            'draw sprinklers
            If chkShowSprink.Checked Then
                oSprinklers = SprinklerDB.GetSprinklers2
                For Each oSprinkler In oSprinklers
                    a = CSng(oSprinkler.sprx * convert)
                    b = rwidth - CSng(oSprinkler.spry * convert)
                    drawroom.DrawRectangle(Pens.Purple, a, b, 5, 5)
                Next
            End If

            'draw items
            For n = 1 To n_max
                a = CSng(item_location(3, n) * convert)
                c = CSng((item_location(4, n) - item_location(3, n)) * convert)
                d = CSng((item_location(6, n) - item_location(5, n)) * convert)
                b = rwidth - CSng((item_location(6, n) * convert))

                'draw item
                'label item
                If n < itime.Count Then
                    If itime(n) < SimTime Then
                        If n > 1 And itime(n) = 0 Then
                            str1 = n.ToString
                        Else
                            str1 = n.ToString & Chr(10) & " (ign " & itime(n) & "s)" & ignmode(n)
                        End If

                    Else
                        str1 = n.ToString
                    End If
                End If
                If n = 1 Then
                    'designate item 1 as the item first ignited
                    drawroom.DrawRectangle(Pens.Black, a, b, c, d)
                    drawroom.DrawString(str1, New Font("Arial", 10), New SolidBrush(Color.Black), a, b)
                Else
                    'other items are secondary targets
                    drawroom.DrawRectangle(Pens.Blue, a, b, c, d)
                    drawroom.DrawString(str1, New Font("Arial", 10), New SolidBrush(Color.Blue), a, b)

                    'draw line vectors from first item ignited to secondary items
                    If CheckBox1.Checked = True Then

                        'str = CStr(Format(vectorlength(n - 2, 0), "0.00")) & " m"
                        str = CStr(Format(vectorlength(n, 1), "0.00")) & " m"
                        drawroom.DrawString(str, New Font("Arial", 8), New SolidBrush(Color.Black), a + c / 2, b + d / 2)
                        drawroom.DrawLine(Pens.Black, CSng(item_location(1, 1) * convert), rwidth - CSng(item_location(2, 1) * convert), CSng(item_location(1, n) * convert), rwidth - CSng(item_location(2, n) * convert))

                    End If

                End If
            Next

            img.Save(file)

            MessageBox.Show("Saved to file: " & file)


        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in frmPopulate Button1_Click_1")

        End Try
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click

        'copy layout to clipboard

        Dim file As String = "layout.jpg"
        Dim blackpen As New Pen(Color.Black, 2)
        Dim rwidth As Single
        Dim rlength As Single
        Dim gridpixels As Single
        Dim convert As Single
        Dim a, b, c, d As Single
        Dim str, str1 As String


        Try
            Dim numFRvents As Integer = 0
            For i = 1 To NumberRooms + 1
                For j = 1 To NumberVents(i, fireroom)
                    numFRvents = numFRvents + 1
                Next
            Next

            'figure out suitable scaling to draw the room on screen
            If RoomLength(fireroom) > RoomWidth(fireroom) Then
                Panel1.Width = 500 'room length
                convert = Panel1.Width / RoomLength(fireroom) 'pixels/metre
                Panel1.Height = convert * RoomWidth(fireroom)  'room width
            Else
                Panel1.Height = 500
                convert = Panel1.Height / RoomWidth(fireroom) 'pixels/metre
                Panel1.Width = convert * RoomLength(fireroom)
            End If

            Dim rect As Rectangle = New Rectangle(0, 0, Panel1.Width, Panel1.Height)
            Dim displayGraphics As Graphics = Me.CreateGraphics()
            Dim img As Image = New Bitmap(rect.Width, rect.Height, displayGraphics)
            Dim drawroom As Graphics = Graphics.FromImage(img)
            Dim drawscale As Graphics = Graphics.FromImage(img)

            gridpixels = CSng(gridsize * convert)
            If gridpixels = 0 Then Exit Sub

            'room dimensions to use in pixels
            rlength = CSng(Panel1.Width)
            rwidth = CSng(Panel1.Height)

            'drawroom = Panel1.CreateGraphics
            drawroom.Clear(Color.White)

            'draw room outline
            'drawroom.DrawRectangle(Pens.Black, 0, 0, rlength, rwidth)
            drawroom.DrawRectangle(blackpen, 0, 0, rlength, rwidth)

            'draw vent keep clear areas
            Dim p As New Pen(Color.CornflowerBlue)
            Dim myBrush As New HatchBrush(HatchStyle.LightUpwardDiagonal, Color.Red)
            p.DashStyle = DashStyle.Dot

            If CheckBox2.Checked Then 'show grid
                For i = 1 To CInt(rwidth / gridpixels)
                    drawroom.DrawLine(p, 0, CInt(i * gridpixels), rlength, CInt(i * gridpixels))
                Next i
                For i = 1 To CInt(rlength / gridpixels)
                    drawroom.DrawLine(p, CInt(i * gridpixels), 0, CInt(i * gridpixels), rwidth)
                Next i
            End If

            p.Color = Color.Red
            For n = 1 To number_vents
                a = CSng(vent_space(n, 0) * convert)
                b = rwidth - CSng(vent_space(n, 3) * convert)
                c = CSng((vent_space(n, 1) - vent_space(n, 0)) * convert)
                d = CSng((vent_space(n, 3) - vent_space(n, 2)) * convert)

                drawroom.DrawRectangle(p, a, b, c, d)
                drawroom.DrawString("keep clear", New Font("Arial", 9), New SolidBrush(Color.Red), a + 5, b)
            Next

            'draw sprinklers
            If chkShowSprink.Checked Then
                oSprinklers = SprinklerDB.GetSprinklers2
                For Each oSprinkler In oSprinklers
                    a = CSng(oSprinkler.sprx * convert)
                    b = rwidth - CSng(oSprinkler.spry * convert)
                    drawroom.DrawRectangle(Pens.Purple, a, b, 5, 5)
                Next
            End If

            'draw items
            If n_max > NumberObjects Then Exit Sub

            For n = 1 To n_max
                If vectorlength.Length <= n Then Exit For
                'centre to centre distances
                vectorlength(n, 1) = Sqrt((item_location(1, n) - item_location(1, 1)) ^ 2 + (item_location(2, n) - item_location(2, 1)) ^ 2)

                a = CSng(item_location(3, n) * convert)
                c = CSng((item_location(4, n) - item_location(3, n)) * convert)
                d = CSng((item_location(6, n) - item_location(5, n)) * convert)
                b = rwidth - CSng((item_location(6, n) * convert))

                'draw item
                'draw item

                'label item
                'If n < itime.Count Then
                '    If itime(n) < SimTime Then
                '        If n > 1 And itime(n) = 0 Then
                '            If optidnum.Checked = True Then
                '                str1 = n.ToString
                '            Else
                '                str1 = ObjLabel(n).ToString
                '            End If
                '        Else
                '            If optidnum.Checked = True Then
                '                str1 = n.ToString & Chr(10) & "" & ignmode(n) & "I " & itime(n) & "s"
                '            Else
                '                str1 = ObjLabel(n).ToString & Chr(10) & "" & ignmode(n) & "I " & itime(n) & "s"
                '            End If
                '        End If
                '    Else
                '        If optidnum.Checked = True Then
                '            str1 = n.ToString
                '        Else
                '            str1 = ObjLabel(n).ToString
                '        End If
                '    End If
                'Else
                '    If optidnum.Checked = True Then
                '        str1 = n.ToString
                '    Else
                '        str1 = ObjLabel(n).ToString
                '    End If
                'End If
                'label item
                If n < itime.Count Then
                    If itime(n) < SimTime Then
                        If n > 1 And itime(n) = 0 Then
                            If optidnum.Checked = True Then
                                str1 = ObjectItemID(n)
                            Else
                                str1 = ObjLabel(n).ToString
                            End If
                        Else
                            If optidnum.Checked = True Then
                                str1 = ObjectItemID(n) & Chr(10) & "" & ignmode(n) & "I " & itime(n) & "s"
                            Else
                                str1 = ObjLabel(n).ToString & Chr(10) & "" & ignmode(n) & "I " & itime(n) & "s"
                            End If
                        End If
                    Else
                        If optidnum.Checked = True Then
                            str1 = n.ToString
                        Else
                            str1 = ObjLabel(n).ToString
                        End If
                    End If
                Else
                    If optidnum.Checked = True Then
                        str1 = n.ToString
                    Else
                        str1 = ObjLabel(n).ToString
                    End If
                End If

                If n = 1 Then
                    'designate item 1 as the item first ignited
                    drawroom.DrawRectangle(Pens.Black, a, b, c, d)
                    drawroom.DrawString(str1, New Font("Arial", 10), New SolidBrush(Color.Black), a, b)
                Else
                    'other items are secondary targets

                    drawroom.DrawRectangle(Pens.Blue, a, b, c, d)
                    drawroom.DrawString(str1, New Font("Arial", 10), New SolidBrush(Color.Blue), a, b)

                    'draw line vectors from first item ignited to secondary items
                    If CheckBox1.Checked = True Then

                        'str = CStr(Format(vectorlength(n, 1), "0.00")) & " m"
                        str = CStr(Format(targetdistance(1, n), "0.000")) & " m"

                        drawroom.DrawString(str, New Font("Arial", 8), New SolidBrush(Color.Black), a + c / 4, b + d / 2)
                        'drawroom.DrawLine(Pens.Black, CSng(item_location(1, 1) * convert), rwidth - CSng(item_location(2, 1) * convert), CSng(item_location(1, n) * convert), rwidth - CSng(item_location(2, n) * convert))

                        drawroom.DrawLine(Pens.Black, CSng(item_location(1, 1) * convert), rwidth - CSng(item_location(2, 1) * convert), CSng(targetxpoint(1, n) * convert), rwidth - CSng(targetypoint(1, n) * convert))

                    End If

                End If
            Next

            Clipboard.SetImage(img)
            img.Dispose()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in ToolStripMenuItem1")

        End Try

    End Sub

    Private Sub CopyLayoutToClipboardToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyLayoutToClipboardToolStripMenuItem.Click
        ToolStripMenuItem1.PerformClick()
    End Sub


    Private Sub txtVentClearance_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVentClearance.Validated
        ErrorProvider1.Clear()
        ventclearance = CSng(txtVentClearance.Text)
    End Sub

    Private Sub txtVentClearance_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtVentClearance.Validating
        If IsNumeric(txtVentClearance.Text) Then
            If (CSng(txtVentClearance.Text) <= Min(RoomWidth(fireroom), RoomLength(fireroom)) And CSng(txtVentClearance.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtVentClearance.Select(0, txtVentClearance.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtVentClearance, "Invalid Entry.")

    End Sub

    Private Sub txtGridSize_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGridSize.Validated
        ErrorProvider1.Clear()
        gridsize = CSng(txtGridSize.Text)
    End Sub

    Private Sub txtGridSize_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtGridSize.Validating
        If IsNumeric(txtGridSize.Text) Then
            If (CSng(txtGridSize.Text) <= 10.0 And CSng(txtGridSize.Text) >= 0.1) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtGridSize.Select(0, txtGridSize.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtGridSize, "Invalid Entry. Grid size must be in the range 0.1 to 10m.")
    End Sub

    Private Sub chkShowSprink_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowSprink.CheckedChanged

    End Sub

    Private Sub txtGridSize_TextChanged(sender As Object, e As EventArgs) Handles txtGridSize.TextChanged

    End Sub
End Class