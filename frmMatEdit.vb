Imports System.Data.OleDb
Public Class frmMatEdit
    Dim surfaceindex(0 To 1) As String
    Dim substrateindex(0 To 1) As String

    Private Sub frmMatEdit_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.CenterToParent()

        'TODO: This line of code loads data into the 'ThermalDataSet1.Table1' table. You can move, or remove it, as needed.
        Me.Table1TableAdapter.Fill(Me.ThermalDataSet1.Table1)

        Dim combo As ComboBox

        Dim label1 As Label
        Dim text1 As TextBox
        Dim text2 As TextBox
        Dim bindingsource1 As BindingSource
        Dim bindingsource2 As BindingSource
        ReDim surfaceindex(0 To NumberRooms)
        ReDim substrateindex(0 To NumberRooms)

        TableLayoutPanel2.RowCount = NumberRooms + 1
        cmd_save.Tag = Me.Tag

        If Me.Tag = "ceiling" Then

            For room = 1 To NumberRooms
                bindingsource1 = New BindingSource
                bindingsource2 = New BindingSource
                bindingsource1.DataSource = ThermalDataSet1
                bindingsource1.DataMember = "Table1"
                bindingsource2.DataSource = ThermalDataSet1
                bindingsource2.DataMember = "Table1"

                label1 = New Label
                label1.Text = room.ToString
                label1.Width = 10

                combo = New ComboBox

                text1 = New TextBox
                text2 = New TextBox
                text1.Text = CeilingThickness(room).ToString
                text2.Text = CeilingSubThickness(room).ToString

                TableLayoutPanel2.RowStyles.Add(New RowStyle(Height = 30))
                TableLayoutPanel2.Controls.Add(label1, 0, room + 1)
                TableLayoutPanel2.Controls.Add(combo, 1, room + 1)
                TableLayoutPanel2.Controls.Add(text1, 2, room + 1)

                combo.DropDownStyle = ComboBoxStyle.DropDownList
                combo.Width = 200
                combo.DataSource = bindingsource1
                combo.ValueMember = "Material Description"
                combo.DisplayMember = "Material Description"
                combo.SelectedValue = CeilingSurface(room)
                surfaceindex(room) = combo.SelectedValue

                combo = New ComboBox

                TableLayoutPanel2.Controls.Add(combo, 4, room + 1)
                TableLayoutPanel2.Controls.Add(text2, 5, room + 1)

                combo.DropDownStyle = ComboBoxStyle.DropDownList
                combo.Width = 200
                combo.DataSource = bindingsource2
                combo.ValueMember = "Material Description"
                combo.DisplayMember = "Material Description"
                combo.SelectedValue = CeilingSubstrate(room)
                substrateindex(room) = combo.SelectedValue
            Next
        ElseIf Me.Tag = "wall" Then
            For room = 1 To NumberRooms
                bindingsource1 = New BindingSource
                bindingsource2 = New BindingSource
                bindingsource1.DataSource = ThermalDataSet1
                bindingsource1.DataMember = "Table1"
                bindingsource2.DataSource = ThermalDataSet1
                bindingsource2.DataMember = "Table1"

                label1 = New Label
                label1.Text = room.ToString
                label1.Width = 10

                combo = New ComboBox

                text1 = New TextBox
                text2 = New TextBox
                text1.Text = WallThickness(room).ToString
                text2.Text = WallSubThickness(room).ToString

                TableLayoutPanel2.RowStyles.Add(New RowStyle(Height = 30))
                TableLayoutPanel2.Controls.Add(label1, 0, room + 1)
                TableLayoutPanel2.Controls.Add(combo, 1, room + 1)
                TableLayoutPanel2.Controls.Add(text1, 2, room + 1)

                combo.DropDownStyle = ComboBoxStyle.DropDownList
                combo.Width = 200
                combo.DataSource = bindingsource1
                combo.ValueMember = "Material Description"
                combo.DisplayMember = "Material Description"
                combo.SelectedValue = WallSurface(room)
                surfaceindex(room) = combo.SelectedValue

                combo = New ComboBox

                TableLayoutPanel2.Controls.Add(combo, 4, room + 1)
                TableLayoutPanel2.Controls.Add(text2, 5, room + 1)

                combo.DropDownStyle = ComboBoxStyle.DropDownList
                combo.Width = 200
                combo.DataSource = bindingsource2
                combo.ValueMember = "Material Description"
                combo.DisplayMember = "Material Description"
                combo.SelectedValue = WallSubstrate(room)
                substrateindex(room) = combo.SelectedValue
            Next
        ElseIf Me.Tag = "floor" Then
            For room = 1 To NumberRooms
                bindingsource1 = New BindingSource
                bindingsource2 = New BindingSource
                bindingsource1.DataSource = ThermalDataSet1
                bindingsource1.DataMember = "Table1"
                bindingsource2.DataSource = ThermalDataSet1
                bindingsource2.DataMember = "Table1"

                label1 = New Label
                label1.Text = room.ToString
                label1.Width = 10

                combo = New ComboBox

                text1 = New TextBox
                text2 = New TextBox
                text1.Text = FloorThickness(room).ToString
                text2.Text = FloorSubThickness(room).ToString

                TableLayoutPanel2.RowStyles.Add(New RowStyle(Height = 30))
                TableLayoutPanel2.Controls.Add(label1, 0, room + 1)
                TableLayoutPanel2.Controls.Add(combo, 1, room + 1)
                TableLayoutPanel2.Controls.Add(text1, 2, room + 1)

                combo.DropDownStyle = ComboBoxStyle.DropDownList
                combo.Width = 200
                combo.DataSource = bindingsource1
                combo.ValueMember = "Material Description"
                combo.DisplayMember = "Material Description"
                combo.SelectedValue = FloorSurface(room)
                surfaceindex(room) = combo.SelectedValue

                combo = New ComboBox

                TableLayoutPanel2.Controls.Add(combo, 4, room + 1)
                TableLayoutPanel2.Controls.Add(text2, 5, room + 1)

                combo.DropDownStyle = ComboBoxStyle.DropDownList
                combo.Width = 200
                combo.DataSource = bindingsource2
                combo.ValueMember = "Material Description"
                combo.DisplayMember = "Material Description"
                combo.SelectedValue = FloorSubstrate(room)
                substrateindex(room) = combo.SelectedValue
            Next
        End If

        Me.AutoSizeMode = AutoSizeMode.GrowAndShrink
        Me.AutoSize = True
        Dim x As Integer = TableLayoutPanel2.Location.X + CInt(TableLayoutPanel2.Width) - cmd_save.Width
        Dim y As Integer = TableLayoutPanel2.Location.Y + CInt(TableLayoutPanel2.Height) + 20
        cmd_save.Location = New Point(x, y)

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles cmd_save.Click

        Dim ct As New Control
        Dim cbo As New ComboBox
        Dim tx As New TextBox
        Dim room As Integer

        'String variable to hold Connection String.
        Dim conString As String = "Provider=Microsoft.Jet.OLEDB.4.0;" &
                                               "Data Source = C:\Program Files\Microsoft Visual Studio\VB98\NWIND.mdb"
        'Make sure you change the Data Source to the appropriate PATH.
        conString = My.Settings.thermalConnectionString

        'Create an oleDbConnection object,
        'and then pass in the ConnectionString to the constructor.
        Dim conn As OleDbConnection = New OleDbConnection(conString)

        Dim selectString As String       'Variable to hold the SQL statement.
        Dim cmd As OleDbCommand         'Create an OleDbCommand object.
        Dim reader As OleDbDataReader   'Create an OleDbDataReader object.

        Try

            For Each cbo In TableLayoutPanel2.Controls.OfType(Of ComboBox)
                Dim cellpos As TableLayoutPanelCellPosition = TableLayoutPanel2.GetCellPosition(cbo)
                If cellpos.Column = 1 Then
                    surfaceindex(cellpos.Row - 1) = cbo.SelectedValue
                    If IsNothing(cbo.SelectedValue) Then surfaceindex(cellpos.Row - 1) = "None"
                End If
                    If cellpos.Column = 4 Then
                    substrateindex(cellpos.Row - 1) = cbo.SelectedValue
                    If IsNothing(cbo.SelectedValue) Then substrateindex(cellpos.Row - 1) = "None"
                End If
            Next
            If Me.Tag = "ceiling" Then

                For Each tx In TableLayoutPanel2.Controls.OfType(Of TextBox)
                    Dim cellpos As TableLayoutPanelCellPosition = TableLayoutPanel2.GetCellPosition(tx)
                    If cellpos.Column = 2 Then
                        If IsNumeric(tx.Text) Then
                            CeilingThickness(cellpos.Row - 1) = tx.Text
                        Else
                            MsgBox("Thickness is not a number.", MsgBoxStyle.Exclamation)
                            CeilingThickness(cellpos.Row - 1) = 0
                        End If

                    End If
                    If cellpos.Column = 5 Then
                        If IsNumeric(tx.Text) Then
                            CeilingSubThickness(cellpos.Row - 1) = tx.Text
                            If tx.Text > 0 Then HaveCeilingSubstrate(cellpos.Row - 1) = True
                        Else
                            MsgBox("Thickness is not a number.", MsgBoxStyle.Exclamation)
                            CeilingSubThickness(cellpos.Row - 1) = 0
                        End If
                    End If
                Next

                'Open the connection.
                conn.Open()

                For room = 1 To NumberRooms
                    'Initialize SQL string.
                    selectString = "SELECT * FROM Table1"

                    'Initialize OleDbCommand object.
                    cmd = New OleDbCommand(selectString, conn)

                    'Send the CommandText to the connection, and then build an OleDbDataReader.
                    reader = cmd.ExecuteReader()

                    'Loop through the records and print the values.
                    While (reader.Read())
                        If reader(1) = surfaceindex(room) Then
                            CeilingSurface(room) = reader(1).ToString
                            CeilingConductivity(room) = reader(2).ToString
                            CeilingSpecificHeat(room) = reader(3).ToString
                            CeilingDensity(room) = reader(4).ToString
                            Surface_Emissivity(1, room) = reader(5).ToString
                            'min temp for spread 6
                            'flamespread parameter 7
                            CeilingConeDataFile(room) = reader(8).ToString
                            CeilingEffectiveHeatofCombustion(room) = reader(9).ToString
                            'comments 10
                            CeilingSootYield(room) = reader(11).ToString
                            CeilingCO2Yield(room) = reader(12).ToString
                            CeilingH2OYield(room) = reader(13).ToString
                            'calibration 14
                            'comments 15
                            'Exit While
                        End If
                        If reader(1) = substrateindex(room) Then
                            CeilingSubstrate(room) = reader(1).ToString
                            CeilingSubConductivity(room) = reader(2).ToString
                            CeilingSubSpecificHeat(room) = reader(3).ToString
                            CeilingSubDensity(room) = reader(4).ToString
                            'Exit While
                        End If
                    End While
                Next
            ElseIf Me.Tag = "wall" Then

                For Each tx In TableLayoutPanel2.Controls.OfType(Of TextBox)
                    Dim cellpos As TableLayoutPanelCellPosition = TableLayoutPanel2.GetCellPosition(tx)
                    If cellpos.Column = 2 Then
                        If IsNumeric(tx.Text) Then
                            WallThickness(cellpos.Row - 1) = tx.Text
                        Else
                            MsgBox("Thickness is not a number.", MsgBoxStyle.Exclamation)
                            WallThickness(cellpos.Row - 1) = 0
                        End If

                    End If
                    If cellpos.Column = 5 Then
                        If IsNumeric(tx.Text) Then
                            WallSubThickness(cellpos.Row - 1) = tx.Text
                            If tx.Text > 0 Then HaveWallSubstrate(cellpos.Row - 1) = True
                        Else
                            MsgBox("Thickness is not a number.", MsgBoxStyle.Exclamation)
                            WallSubThickness(cellpos.Row - 1) = 0
                        End If
                    End If
                Next

                'Open the connection.
                conn.Open()

                For room = 1 To NumberRooms
                    'Initialize SQL string.
                    selectString = "SELECT * FROM Table1"

                    'Initialize OleDbCommand object.
                    cmd = New OleDbCommand(selectString, conn)

                    'Send the CommandText to the connection, and then build an OleDbDataReader.
                    reader = cmd.ExecuteReader()

                    'Loop through the records and print the values.
                    While (reader.Read())
                        If reader(1) = surfaceindex(room) Then
                            WallSurface(room) = reader(1).ToString
                            WallConductivity(room) = reader(2).ToString
                            WallSpecificHeat(room) = reader(3).ToString
                            WallDensity(room) = reader(4).ToString
                            Surface_Emissivity(2, room) = reader(5).ToString
                            Surface_Emissivity(3, room) = reader(5).ToString
                            'min temp for spread 6
                            WallTSMin(room) = reader(6).ToString
                            'flamespread parameter 7
                            WallFlameSpreadParameter(room) = reader(7).ToString
                            WallConeDataFile(room) = reader(8).ToString
                            WallEffectiveHeatofCombustion(room) = reader(9).ToString
                            'comments 10
                            WallSootYield(room) = reader(11).ToString
                            WallCO2Yield(room) = reader(12).ToString
                            WallH2OYield(room) = reader(13).ToString
                            'calibration 14
                            'comments 15

                        End If
                        If reader(1) = substrateindex(room) Then
                            WallSubstrate(room) = reader(1).ToString
                            WallSubConductivity(room) = reader(2).ToString
                            WallSubSpecificHeat(room) = reader(3).ToString
                            WallSubDensity(room) = reader(4).ToString

                        End If
                    End While
                Next
            ElseIf Me.Tag = "floor" Then
                For Each tx In TableLayoutPanel2.Controls.OfType(Of TextBox)
                    Dim cellpos As TableLayoutPanelCellPosition = TableLayoutPanel2.GetCellPosition(tx)
                    If cellpos.Column = 2 Then
                        If IsNumeric(tx.Text) Then
                            FloorThickness(cellpos.Row - 1) = tx.Text
                        Else
                            MsgBox("Thickness is not a number.", MsgBoxStyle.Exclamation)
                            FloorThickness(cellpos.Row - 1) = 0
                        End If

                    End If
                    If cellpos.Column = 5 Then
                        If IsNumeric(tx.Text) Then
                            FloorSubThickness(cellpos.Row - 1) = tx.Text
                            If tx.Text > 0 Then HaveFloorSubstrate(cellpos.Row - 1) = True
                        Else
                            MsgBox("Thickness is not a number.", MsgBoxStyle.Exclamation)
                            FloorSubThickness(cellpos.Row - 1) = 0
                        End If
                    End If
                Next

                'Open the connection.
                conn.Open()

                For room = 1 To NumberRooms
                    'Initialize SQL string.
                    selectString = "SELECT * FROM Table1"

                    'Initialize OleDbCommand object.
                    cmd = New OleDbCommand(selectString, conn)

                    'Send the CommandText to the connection, and then build an OleDbDataReader.
                    reader = cmd.ExecuteReader()

                    'Loop through the records and print the values.
                    While (reader.Read())
                        If reader(1) = surfaceindex(room) Then
                            FloorSurface(room) = reader(1).ToString
                            FloorConductivity(room) = reader(2).ToString
                            FloorSpecificHeat(room) = reader(3).ToString
                            FloorDensity(room) = reader(4).ToString
                            Surface_Emissivity(4, room) = reader(5).ToString

                            'min temp for spread 6
                            FloorTSMin(room) = reader(6).ToString
                            'flamespread parameter 7
                            FloorFlameSpreadParameter(room) = reader(7).ToString
                            FloorConeDataFile(room) = reader(8).ToString
                            FloorEffectiveHeatofCombustion(room) = reader(9).ToString
                            'comments 10
                            FloorSootYield(room) = reader(11).ToString
                            FloorCO2Yield(room) = reader(12).ToString
                            FloorH2OYield(room) = reader(13).ToString
                            'calibration 14
                            'comments 15

                        End If
                        If reader(1) = substrateindex(room) Then
                            FloorSubstrate(room) = reader(1).ToString
                            FloorSubConductivity(room) = reader(2).ToString
                            FloorSubSpecificHeat(room) = reader(3).ToString
                            FloorSubDensity(room) = reader(4).ToString

                        End If
                    End While
                Next
            End If

            'Close the reader and the related connection.
            reader.Close()
            conn.Close()
            Close()

        Catch Excep As System.Exception
            MessageBox.Show(Excep.Message, "Access Database", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class