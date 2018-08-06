Imports System.Xml
Imports System.Collections.Generic

Public Class SprinklerDB
    ' Use file in project directory

    Public Shared Function RemoveSprinklers(ByVal osprinklers As List(Of oSprinkler), ByVal osprdistributions As List(Of oDistribution))
        'get sprinkler details from sprinklers.xml file

        osprinklers = SprinklerDB.GetSprinklers2()

        Dim icount As Integer = osprinklers.Count

        Dim j As Integer = 0
        If icount > 0 Then

            osprinklers.Clear()

            osprdistributions = SprinklerDB.GetSprDistributions()
            osprdistributions.Clear()

            SprinklerDB.SaveSprinklers2(osprinklers, osprdistributions)

        End If

        Return osprinklers
    End Function
    Public Shared Function GetSprinklers2() As List(Of oSprinkler)

        'get sprinkler details from sprinklers.xml file
        Dim osprinklers As New List(Of oSprinkler)
        Dim osprdistributions As New List(Of oDistribution)

        Dim Path As String = ProjectDirectory & "sprinklers.xml"
        Dim xmlIn As New XmlTextReader(Path)

        If My.Computer.FileSystem.FileExists(Path) = False Then
            MsgBox("The file sprinklers.xml is missing", MsgBoxStyle.Critical, "Missing File")
            xmlIn.Close()
            Return osprinklers
            DetectorType(fireroom) = 0
            Exit Function
        End If

        If My.Computer.FileSystem.FileExists(Path) Then

            xmlIn.WhitespaceHandling = WhitespaceHandling.None
            Dim oldfile As Boolean = True
            Do While xmlIn.Name <> "Sprinkler"
                xmlIn.Read()
                If xmlIn.EOF = True Then 'no sprinklers
                    xmlIn.Close()
                    Return osprinklers
                    DetectorType(fireroom) = 0
                    Exit Function
                End If
                If xmlIn.Name = "version" Then
                    oldfile = False
                End If
            Loop
            If oldfile = True Then
                MsgBox("Your sprinklers.xml file is obsolete. Please delete it and try again.", MsgBoxStyle.Exclamation)
                xmlIn.Close()
                Return osprinklers
                DetectorType(fireroom) = 0
                Exit Function
            End If
            DetectorType(fireroom) = 2

            Do While xmlIn.Name = "Sprinkler"
                Dim osprinkler As New oSprinkler
                xmlIn.ReadStartElement("Sprinkler")
                osprinkler.sprid = CSng(xmlIn.ReadElementString("sprid"))
                osprinkler.room = CSng(xmlIn.ReadElementString("room"))
                osprinkler.sprx = CSng(xmlIn.ReadElementString("sprx"))
                osprinkler.spry = CSng(xmlIn.ReadElementString("spry"))
                osprinkler.responsetime = CSng(xmlIn.ReadElementString("responsetime"))

                Dim bx As Integer = 0
                Do While xmlIn.Name = "sdistribution"
                    Dim odistribution As New oDistribution
                    bx = bx + 1
                    xmlIn.ReadStartElement("sdistribution")
                    odistribution.id = osprinkler.sprid 'saves id
                    odistribution.varname = CStr(xmlIn.ReadElementString("varname"))
                    odistribution.varvalue = CStr(xmlIn.ReadElementString("value"))
                    odistribution.distribution = CStr(xmlIn.ReadElementString("distribution"))
                    odistribution.mean = CStr(xmlIn.ReadElementString("mean"))
                    odistribution.variance = CSng(xmlIn.ReadElementString("variance"))
                    odistribution.lbound = CSng(xmlIn.ReadElementString("lbound"))
                    odistribution.ubound = CSng(xmlIn.ReadElementString("ubound"))
                    odistribution.mode = CSng(xmlIn.ReadElementString("mode"))
                    odistribution.alpha = CSng(xmlIn.ReadElementString("alpha"))
                    odistribution.beta = CSng(xmlIn.ReadElementString("beta"))
                    xmlIn.ReadEndElement()
                    osprdistributions.Add(odistribution)
                Loop


                xmlIn.ReadEndElement()

                osprinklers.Add(osprinkler)

            Loop
        End If

        xmlIn.Close()
        Return osprinklers
    End Function

    Public Shared Sub SaveSprinklers2(ByVal osprinklers As List(Of oSprinkler), ByVal osprdistributions As List(Of oDistribution))

        Dim rname As String = ProjectDirectory & "sprinklers.xml"
        Dim Path As String = rname
        Dim xmlOut As New XmlTextWriter(Path, System.Text.Encoding.UTF8)

        Try

            xmlOut.Formatting = Formatting.Indented

            xmlOut.WriteStartDocument()
            xmlOut.WriteStartElement("Sprinklers")
            xmlOut.WriteElementString("version", Version.ToString)

            Dim osprinkler As oSprinkler
            For i As Integer = 0 To osprinklers.Count - 1
                osprinkler = CType(osprinklers(i), oSprinkler)
                xmlOut.WriteStartElement("Sprinkler")
                xmlOut.WriteElementString("sprid", CStr(i + 1))
                xmlOut.WriteElementString("room", osprinkler.room.ToString)

                xmlOut.WriteElementString("sprx", osprinkler.sprx.ToString)
                xmlOut.WriteElementString("spry", osprinkler.spry.ToString)
                xmlOut.WriteElementString("responsetime", osprinkler.responsetime.ToString)

                If Not IsNothing(osprdistributions) Then
                    For Each odistribution In osprdistributions
                        If odistribution.id = i + 1 Then
                            xmlOut.WriteStartElement("sdistribution")
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
                    xmlOut.WriteStartElement("sdistribution")
                    xmlOut.WriteElementString("varname", "rti")
                    xmlOut.WriteElementString("value", "50")
                    xmlOut.WriteElementString("distribution", "None")
                    xmlOut.WriteElementString("mean", "0")
                    xmlOut.WriteElementString("variance", "0")
                    xmlOut.WriteElementString("lbound", "0")
                    xmlOut.WriteElementString("ubound", "0")
                    xmlOut.WriteElementString("mode", "0")
                    xmlOut.WriteElementString("alpha", "0")
                    xmlOut.WriteElementString("beta", "0")
                    xmlOut.WriteEndElement()

                    xmlOut.WriteStartElement("sdistribution")
                    xmlOut.WriteElementString("varname", "cfactor")
                    xmlOut.WriteElementString("value", "0.4")
                    xmlOut.WriteElementString("distribution", "None")
                    xmlOut.WriteElementString("mean", "0")
                    xmlOut.WriteElementString("variance", "0")
                    xmlOut.WriteElementString("lbound", "0")
                    xmlOut.WriteElementString("ubound", "0")
                    xmlOut.WriteElementString("mode", "0")
                    xmlOut.WriteElementString("alpha", "0")
                    xmlOut.WriteElementString("beta", "0")
                    xmlOut.WriteEndElement()

                    xmlOut.WriteStartElement("sdistribution")
                    xmlOut.WriteElementString("varname", "sprdensity")
                    xmlOut.WriteElementString("value", "4.2")
                    xmlOut.WriteElementString("distribution", "None")
                    xmlOut.WriteElementString("mean", "0")
                    xmlOut.WriteElementString("variance", "0")
                    xmlOut.WriteElementString("lbound", "0")
                    xmlOut.WriteElementString("ubound", "0")
                    xmlOut.WriteElementString("mode", "0")
                    xmlOut.WriteElementString("alpha", "0")
                    xmlOut.WriteElementString("beta", "0")
                    xmlOut.WriteEndElement()

                    xmlOut.WriteStartElement("sdistribution")
                    xmlOut.WriteElementString("varname", "sprr")
                    xmlOut.WriteElementString("value", "0")
                    xmlOut.WriteElementString("distribution", "None")
                    xmlOut.WriteElementString("mean", "0")
                    xmlOut.WriteElementString("variance", "0")
                    xmlOut.WriteElementString("lbound", "0")
                    xmlOut.WriteElementString("ubound", "0")
                    xmlOut.WriteElementString("mode", "0")
                    xmlOut.WriteElementString("alpha", "0")
                    xmlOut.WriteElementString("beta", "0")
                    xmlOut.WriteEndElement()

                    xmlOut.WriteStartElement("sdistribution")
                    xmlOut.WriteElementString("varname", "sprz")
                    xmlOut.WriteElementString("value", "0.025")
                    xmlOut.WriteElementString("distribution", "None")
                    xmlOut.WriteElementString("mean", "0")
                    xmlOut.WriteElementString("variance", "0")
                    xmlOut.WriteElementString("lbound", "0")
                    xmlOut.WriteElementString("ubound", "0")
                    xmlOut.WriteElementString("mode", "0")
                    xmlOut.WriteElementString("alpha", "0")
                    xmlOut.WriteElementString("beta", "0")
                    xmlOut.WriteEndElement()

                    xmlOut.WriteStartElement("sdistribution")
                    xmlOut.WriteElementString("varname", "acttemp")
                    xmlOut.WriteElementString("value", "74")
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

            NumSprinklers = osprinklers.Count

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in SprinklerDB.vb SaveSprinklers2")
            xmlOut.WriteEndDocument()
            xmlOut.Close()
        End Try
    End Sub
    Public Shared Function GetSprDistributions() As List(Of oDistribution)

        Dim osprinklers As New List(Of oSprinkler)
        Dim osprdistributions As New List(Of oDistribution)
        Dim Path As String = RiskDataDirectory & "sprinklers.xml"
        Dim ver As Single
        Dim xmlIn As New XmlTextReader(Path)

        Try

            If My.Computer.FileSystem.DirectoryExists(RiskDataDirectory) = True Then
                If My.Computer.FileSystem.FileExists(RiskDataDirectory & "sprinklers.xml") Then

                    xmlIn.WhitespaceHandling = WhitespaceHandling.None

                    Do While xmlIn.Name <> "Sprinkler"
                        xmlIn.Read()
                        If xmlIn.Name = "version" Then
                            ver = CSng(xmlIn.ReadElementString("version"))
                        End If
                        If xmlIn.EOF = True Then 'no items
                            xmlIn.Close()
                            Return osprdistributions
                            Exit Function
                        End If
                    Loop

                    Dim bx = 0
                    Do While xmlIn.Name = "Sprinkler"
                        Dim osprinkler As New oSprinkler

                        xmlIn.ReadStartElement("Sprinkler")
                        bx = bx + 1
                        osprinkler.sprid = bx


                        Do While xmlIn.Name <> "sdistribution"
                            xmlIn.Read()
                            If xmlIn.EOF Then
                                MsgBox("distributions not found in sprinklers.xml file")
                                xmlIn.Close()
                                Return osprdistributions
                                Exit Function
                            End If
                        Loop

                        Do While xmlIn.Name = "sdistribution"
                            Dim odistribution As New oDistribution

                            xmlIn.ReadStartElement("sdistribution")
                            odistribution.id = osprinkler.sprid 'saves id
                            'odistribution.id = CStr(xmlIn.ReadElementString("id"))
                            odistribution.varname = CStr(xmlIn.ReadElementString("varname"))
                            If ver < 2012.5 And odistribution.varname = "sprdensity" Then
                                odistribution.varvalue = 60 * CSng(xmlIn.ReadElementString("value"))
                            Else
                                odistribution.varvalue = CStr(xmlIn.ReadElementString("value"))
                            End If
                            'odistribution.varvalue = CStr(xmlIn.ReadElementString("value"))
                            odistribution.distribution = CStr(xmlIn.ReadElementString("distribution"))
                            If odistribution.distribution = "none" Then odistribution.distribution = "None"
                            odistribution.mean = CStr(xmlIn.ReadElementString("mean"))
                            odistribution.variance = CSng(xmlIn.ReadElementString("variance"))
                            odistribution.lbound = CSng(xmlIn.ReadElementString("lbound"))
                            odistribution.ubound = CSng(xmlIn.ReadElementString("ubound"))
                            odistribution.mode = CSng(xmlIn.ReadElementString("mode"))
                            odistribution.alpha = CSng(xmlIn.ReadElementString("alpha"))
                            odistribution.beta = CSng(xmlIn.ReadElementString("beta"))
                            xmlIn.ReadEndElement()
                            osprdistributions.Add(odistribution)
                        Loop

                        xmlIn.ReadEndElement()
                        osprinklers.Add(osprinkler)

                    Loop
                End If
            End If

            xmlIn.Close()
            Return osprdistributions

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in SprinklerDB.vb GetSprDistributions")
            xmlIn.Close()
            Return osprdistributions
        End Try

    End Function
End Class
