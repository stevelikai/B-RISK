Imports System.Xml
Imports System.Collections.Generic

Public Class RoomDB
    Public Shared Function RemoveRooms(ByVal oRooms As List(Of oRoom), ByVal oroomdistributions As List(Of oDistribution))
        'get room details from rooms.xml file
        oRooms = RoomDB.GetRooms()

        Dim icount As Integer = oRooms.Count

        Dim j As Integer = 0
        If icount > 0 Then

            oRooms.Clear()
            oroomdistributions = RoomDB.GetRoomDistributions()
            oroomdistributions.Clear()
            RoomDB.SaveRooms(oRooms, oroomdistributions)

        End If

        Return oRooms
    End Function
    Public Shared Function GetRooms() As List(Of oRoom)

        Dim oRooms As New List(Of oRoom)
        Dim oroomdistributions As New List(Of oDistribution)
        ProjectDirectory = RiskDataDirectory

        Dim Path As String = ProjectDirectory & "rooms.xml"
        Dim xmlIn As New XmlTextReader(Path)

        If My.Computer.FileSystem.FileExists(Path) = False Then
            MsgBox("The file rooms.xml is missing. The file will be created for you.", MsgBoxStyle.Information, "Creating Room File")

            If My.Computer.FileSystem.FileExists(ApplicationPath & "\data\" & "rooms.xml") = True Then
                My.Computer.FileSystem.CopyFile(ApplicationPath & "\data\" & "rooms.xml", RiskDataDirectory & "rooms.xml", True)
            End If

            Dim oroom As oRoom

            For i = 1 To NumberRooms
                oroom = New oRoom(i, RoomDescription(i), CeilingSlope(i), RoomAbsX(i), RoomAbsY(i), FloorElevation(i), RoomHeight(i), MinStudHeight(i), RoomLength(i), RoomWidth(i))
                oRooms.Add(oroom)

                Dim oDistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
                oDistribution.id = i
                oDistribution.varname = "length"
                oDistribution.varvalue = RoomLength(i)
                oroomdistributions.Add(oDistribution)

                oDistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
                oDistribution.id = i
                oDistribution.varname = "width"
                oDistribution.varvalue = RoomWidth(i)
                oroomdistributions.Add(oDistribution)

            Next

            RoomDB.SaveRooms(oRooms, oroomdistributions)
            xmlIn.Close()
            Return oRooms
            Exit Function

        End If

        If My.Computer.FileSystem.FileExists(Path) Then
            xmlIn.WhitespaceHandling = WhitespaceHandling.None
            Dim oldfile As Boolean = True
            Dim roomver As Single = 0

            Do While xmlIn.Name <> "room"
                xmlIn.Read()
                If xmlIn.EOF = True Then 'no room
                    xmlIn.Close()
                    Return oRooms
                    Exit Function
                End If
                If xmlIn.Name = "version" Then
                    oldfile = False
                    roomver = CSng(xmlIn.ReadElementString("version"))
                End If
            Loop

            If oldfile = True Then
                MsgBox("Your rooms.xml file is obsolete. Please delete it and try again.", MsgBoxStyle.Exclamation)
                xmlIn.Close()
                Return oRooms
                Exit Function
            End If

            Do While xmlIn.Name = "room"

                Dim oroom As New oRoom
                xmlIn.ReadStartElement("room")
                oroom.num = CInt(xmlIn.ReadElementString("room_id"))
                oroom.length = CDbl(xmlIn.ReadElementString("room_length"))
                oroom.width = CDbl(xmlIn.ReadElementString("room_width"))
                oroom.minheight = CDbl(xmlIn.ReadElementString("room_minheight"))
                oroom.maxheight = CDbl(xmlIn.ReadElementString("room_maxheight"))
                oroom.elevation = CDbl(xmlIn.ReadElementString("room_elevation"))
                oroom.absx = CSng(xmlIn.ReadElementString("room_absx"))
                oroom.absy = CSng(xmlIn.ReadElementString("room_absy"))
                oroom.description = CStr(xmlIn.ReadElementString("room_description"))

                If oroom.minheight < oroom.maxheight Then
                    oroom.cslope = True
                Else
                    oroom.cslope = False
                End If

                Dim bx As Integer = 0
                Do While xmlIn.Name = "distribution"
                    Dim odistribution As New oDistribution
                    bx = bx + 1
                    xmlIn.ReadStartElement("distribution")
                    odistribution.id = oroom.num 'saves id
                    odistribution.varname = CStr(xmlIn.ReadElementString("varname"))
                    odistribution.varvalue = xmlIn.ReadElementString("value")

                    'only include variables that have a distribution
                    If odistribution.varname = "width" Then oroom.width = odistribution.varvalue
                    If odistribution.varname = "length" Then oroom.length = odistribution.varvalue

                    odistribution.distribution = CStr(xmlIn.ReadElementString("distribution"))
                    odistribution.mean = xmlIn.ReadElementString("mean")
                    odistribution.variance = xmlIn.ReadElementString("variance")
                    odistribution.lbound = xmlIn.ReadElementString("lbound")
                    odistribution.ubound = xmlIn.ReadElementString("ubound")
                    odistribution.mode = xmlIn.ReadElementString("mode")
                    odistribution.alpha = xmlIn.ReadElementString("alpha")
                    odistribution.beta = xmlIn.ReadElementString("beta")
                    xmlIn.ReadEndElement()

                    oroomdistributions.Add(odistribution)

                Loop

                xmlIn.ReadEndElement()

                oRooms.Add(oroom)

            Loop

        End If

        xmlIn.Close()
        Return oRooms
        Exit Function


    End Function
    Public Shared Sub SaveRooms(ByVal oRooms As List(Of oRoom), ByVal oroomdistributions As List(Of oDistribution))

        Dim rname As String = ProjectDirectory & "rooms.xml"
        Dim Path As String = rname
        Dim xmlOut As New XmlTextWriter(Path, System.Text.Encoding.UTF8)

        Try
            xmlOut.Formatting = Formatting.Indented

            xmlOut.WriteStartDocument()
            xmlOut.WriteStartElement("rooms")
            xmlOut.WriteElementString("version", Version.ToString)

            Dim oRoom As oRoom
            For i As Integer = 0 To oRooms.Count - 1
                oRoom = CType(oRooms(i), oRoom)
                xmlOut.WriteStartElement("room")
                xmlOut.WriteElementString("room_id", CStr(i + 1))
                xmlOut.WriteElementString("room_length", oRoom.length.ToString)
                xmlOut.WriteElementString("room_width", oRoom.width.ToString)
                xmlOut.WriteElementString("room_minheight", oRoom.minheight.ToString)
                xmlOut.WriteElementString("room_maxheight", oRoom.maxheight.ToString)
                xmlOut.WriteElementString("room_elevation", oRoom.elevation.ToString)
                xmlOut.WriteElementString("room_absx", oRoom.absx.ToString)
                xmlOut.WriteElementString("room_absy", oRoom.absy.ToString)
                xmlOut.WriteElementString("room_description", oRoom.description.ToString)

                If Not IsNothing(oroomdistributions) Then
                    For Each odistribution In oroomdistributions
                        If odistribution.id = i + 1 Then
                            xmlOut.WriteStartElement("distribution")
                            xmlOut.WriteElementString("varname", odistribution.varname.ToString)
                            xmlOut.WriteElementString("value", odistribution.varvalue.ToString)
                            xmlOut.WriteElementString("distribution", odistribution.distribution.ToString)
                            xmlOut.WriteElementString("mean", odistribution.mean.ToString)
                            xmlOut.WriteElementString("variance", odistribution.variance.ToString)
                            xmlOut.WriteElementString("lbound", odistribution.lbound.ToString)
                            xmlOut.WriteElementString("ubound", odistribution.ubound.ToString)
                            xmlOut.WriteElementString("mode", odistribution.mode.ToString)
                            xmlOut.WriteElementString("alpha", odistribution.alpha.ToString)
                            xmlOut.WriteElementString("beta", odistribution.beta.ToString)
                            xmlOut.WriteEndElement()
                        End If
                    Next
                Else
                    xmlOut.WriteStartElement("distribution")
                    xmlOut.WriteElementString("varname", "length")
                    xmlOut.WriteElementString("value", "3.6")
                    xmlOut.WriteElementString("distribution", "None")
                    xmlOut.WriteElementString("mean", "0")
                    xmlOut.WriteElementString("variance", "0")
                    xmlOut.WriteElementString("lbound", "0")
                    xmlOut.WriteElementString("ubound", "0")
                    xmlOut.WriteElementString("mode", "0")
                    xmlOut.WriteElementString("alpha", "0")
                    xmlOut.WriteElementString("beta", "0")
                    xmlOut.WriteEndElement()

                    xmlOut.WriteStartElement("distribution")
                    xmlOut.WriteElementString("varname", "width")
                    xmlOut.WriteElementString("value", "2.4")
                    xmlOut.WriteElementString("distribution", "None")
                    xmlOut.WriteElementString("mean", "0")
                    xmlOut.WriteElementString("variance", "0")
                    xmlOut.WriteElementString("lbound", "0")
                    xmlOut.WriteElementString("ubound", "0")
                    xmlOut.WriteElementString("mode", "0")
                    xmlOut.WriteElementString("alpha", "0")
                    xmlOut.WriteElementString("beta", "0")
                    xmlOut.WriteEndElement()

                End If

                xmlOut.WriteEndElement()
            Next

            xmlOut.WriteEndElement()
            xmlOut.WriteEndDocument()
            xmlOut.Close()
            NumberRooms = oRooms.Count
            Exit Sub

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in RoomDB.vb SaveRooms")
            xmlOut.WriteEndDocument()
            xmlOut.Close()
        End Try
    End Sub
    Public Shared Function GetRoomDistributions() As List(Of oDistribution)

        Dim oRooms As New List(Of oRoom)
        Dim oroomdistributions As New List(Of oDistribution)
        Dim oroomdistributions2 As New List(Of oDistribution)
        Dim Path As String = RiskDataDirectory & "rooms.xml"
        Dim ver As Single
        Dim xmlIn As New XmlTextReader(Path)
        Dim vflag As Boolean = False
        Try

            If My.Computer.FileSystem.DirectoryExists(RiskDataDirectory) = True Then
                If My.Computer.FileSystem.FileExists(RiskDataDirectory & "rooms.xml") Then

                    xmlIn.WhitespaceHandling = WhitespaceHandling.None

                    Do While xmlIn.Name <> "room"
                        xmlIn.Read()
                        If xmlIn.Name = "version" Then
                            ver = CSng(xmlIn.ReadElementString("version"))
                        End If
                        If xmlIn.EOF = True Then 'no rooms
                            xmlIn.Close()
                            Return oroomdistributions
                            Exit Function
                        End If
                    Loop

                    Dim bx = 0
                    Do While xmlIn.Name = "room"
                        Dim oRoom As New oRoom

                        xmlIn.ReadStartElement("room")
                        bx = bx + 1
                        oRoom.num = bx


                        Do While xmlIn.Name <> "distribution"
                            xmlIn.Read()
                            If xmlIn.EOF Then
                                MsgBox("distributions not found in rooms.xml file")
                                xmlIn.Close()
                                Return oroomdistributions
                                Exit Function
                            End If
                        Loop

                        Do While xmlIn.Name = "distribution"
                            Dim odistribution As New oDistribution

                            xmlIn.ReadStartElement("distribution")
                            odistribution.id = oRoom.num 'saves id
                            odistribution.varname = CStr(xmlIn.ReadElementString("varname"))
                            odistribution.varvalue = CDbl(xmlIn.ReadElementString("value"))
                            odistribution.distribution = CStr(xmlIn.ReadElementString("distribution"))
                            If odistribution.distribution = "none" Then odistribution.distribution = "None"
                            odistribution.mean = CDbl(xmlIn.ReadElementString("mean"))
                            odistribution.variance = CDbl(xmlIn.ReadElementString("variance"))
                            odistribution.lbound = CDbl(xmlIn.ReadElementString("lbound"))
                            odistribution.ubound = CDbl(xmlIn.ReadElementString("ubound"))
                            odistribution.mode = CDbl(xmlIn.ReadElementString("mode"))
                            odistribution.alpha = CDbl(xmlIn.ReadElementString("alpha"))
                            odistribution.beta = CDbl(xmlIn.ReadElementString("beta"))
                            xmlIn.ReadEndElement()
                            oroomdistributions.Add(odistribution)
                        Loop


                        xmlIn.ReadEndElement()
                        oRooms.Add(oRoom)

                    Loop

                End If
            End If

            xmlIn.Close()
            Return oroomdistributions
            Exit Function

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in RoomDB.vb GetRoomDistributions")
            xmlIn.Close()
            Return oroomdistributions
            Exit Function
        End Try

    End Function
End Class
