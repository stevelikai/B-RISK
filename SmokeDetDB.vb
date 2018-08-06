Imports System.Xml
Imports System.Collections.Generic

Public Class SmokeDetDB
    ' Use file in project directory

    Public Shared Function RemoveSmokeDets(ByVal oSmokeDets As List(Of oSmokeDet), ByVal osddistributions As List(Of oDistribution))

        oSmokeDets = SmokeDetDB.GetSmokDets()

        Dim icount As Integer = oSmokeDets.Count

        Dim j As Integer = 0
        If icount > 0 Then

            oSmokeDets.Clear()

            osddistributions = SmokeDetDB.GetSDDistributions()
            osddistributions.Clear()

            SmokeDetDB.SaveSmokeDets(oSmokeDets, osddistributions)

        End If

        Return oSmokeDets
    End Function

    Public Shared Function GetSmokDets() As List(Of oSmokeDet)

        'get sd details from smokedets.xml file
        Dim oSmokedets As New List(Of oSmokeDet)
        Dim osddistributions As New List(Of oDistribution)

        Dim Path As String = ProjectDirectory & "smokedets.xml"
        Dim xmlIn As New XmlTextReader(Path)

        If My.Computer.FileSystem.FileExists(Path) = False Then
            MsgBox("The file smokedets.xml is missing. A new file will be a created, please check your smoke detector inputs.", MsgBoxStyle.Critical, "Missing File")

            'My.Computer.FileSystem.CopyFile(UserAppDataFolder & gcs_folder_ext & "\" & "riskdata\basemodel_default\" & "smokedets.xml", RiskDataDirectory & "fans.xml", True)
            My.Computer.FileSystem.CopyFile(DataFolder & "smokedets.xml", RiskDataDirectory & "fans.xml", True)

            'xmlIn.Close()
            'Return oSmokedets
            'Exit Function
        End If

        If My.Computer.FileSystem.FileExists(Path) Then

            xmlIn.WhitespaceHandling = WhitespaceHandling.None
            Dim oldfile As Boolean = True
            Dim sdver As Single

            Do While xmlIn.Name <> "Smoke_Detector"
                xmlIn.Read()
                If xmlIn.EOF = True Then 'no sd
                    xmlIn.Close()
                    Return oSmokedets
                    Exit Function
                End If
                If xmlIn.Name = "version" Then
                    oldfile = False
                    sdver = CSng(xmlIn.ReadElementString("version"))
                End If
            Loop
            If oldfile = True Then
                MsgBox("Your smokedets.xml file is obsolete. Please delete it and try again.", MsgBoxStyle.Exclamation)
                xmlIn.Close()
                Return oSmokedets
                Exit Function
            End If

            Do While xmlIn.Name = "Smoke_Detector"
                Dim oSmokeDet As New oSmokeDet
                xmlIn.ReadStartElement("Smoke_Detector")
                oSmokeDet.sdid = CSng(xmlIn.ReadElementString("sdid"))
                oSmokeDet.room = CInt(xmlIn.ReadElementString("room"))

                HaveSD(oSmokeDet.room) = True

                oSmokeDet.sdx = CSng(xmlIn.ReadElementString("sdx"))
                oSmokeDet.sdy = CSng(xmlIn.ReadElementString("sdy"))
                oSmokeDet.responsetime = CSng(xmlIn.ReadElementString("responsetime"))
                oSmokeDet.sdinside = CBool(xmlIn.ReadElementString("sdinside"))

                If sdver < CSng("2012.40") Then oSmokeDet.charlength = CSng(xmlIn.ReadElementString("charlength"))
                If sdver > CSng("2013.05") Then oSmokeDet.transit = CSng(xmlIn.ReadElementString("transit"))

                If sdver > CSng("2015.02") Then
                    oSmokeDet.sdbeam = CBool(xmlIn.ReadElementString("sdbeam"))
                    oSmokeDet.sdbeamalarmtrans = CDbl(xmlIn.ReadElementString("sdbeamalarmtrans"))
                    oSmokeDet.sdbeampathlength = CDbl(xmlIn.ReadElementString("sdbeampathlength"))
                End If

                Dim bx As Integer = 0
                Do While xmlIn.Name = "sdistribution"
                    Dim odistribution As New oDistribution
                    bx = bx + 1
                    xmlIn.ReadStartElement("sdistribution")
                    odistribution.id = oSmokeDet.sdid 'saves id
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
                    osddistributions.Add(odistribution)
                Loop

                xmlIn.ReadEndElement()

                oSmokedets.Add(oSmokeDet)

            Loop
        End If

        xmlIn.Close()
        Return oSmokedets

    End Function

    Public Shared Sub SaveSmokeDets(ByVal oSmokeDets As List(Of oSmokeDet), ByVal osddistributions As List(Of oDistribution))

        Dim rname As String = ProjectDirectory & "smokedets.xml"
        Dim Path As String = rname
        Dim xmlOut As New XmlTextWriter(Path, System.Text.Encoding.UTF8)

        Try

            xmlOut.Formatting = Formatting.Indented

            xmlOut.WriteStartDocument()
            xmlOut.WriteStartElement("Smoke_Detectors")
            xmlOut.WriteElementString("version", Version.ToString)

            Dim oSmokeDet As oSmokeDet
            For i As Integer = 0 To oSmokeDets.Count - 1
                oSmokeDet = CType(oSmokeDets(i), oSmokeDet)
                xmlOut.WriteStartElement("Smoke_Detector")
                xmlOut.WriteElementString("sdid", CStr(i + 1))
                xmlOut.WriteElementString("room", oSmokeDet.room.ToString)
                xmlOut.WriteElementString("sdx", oSmokeDet.sdx.ToString)
                xmlOut.WriteElementString("sdy", oSmokeDet.sdy.ToString)
                xmlOut.WriteElementString("responsetime", oSmokeDet.responsetime.ToString)
                xmlOut.WriteElementString("sdinside", oSmokeDet.sdinside.ToString)
                xmlOut.WriteElementString("transit", oSmokeDet.transit.ToString)
                xmlOut.WriteElementString("sdbeam", oSmokeDet.sdbeam.ToString)
                xmlOut.WriteElementString("sdbeamalarmtrans", oSmokeDet.sdbeamalarmtrans.ToString)
                xmlOut.WriteElementString("sdbeampathlength", oSmokeDet.sdbeampathlength.ToString)

                If Not IsNothing(osddistributions) Then
                    For Each odistribution In osddistributions
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
                    xmlOut.WriteElementString("varname", "OD")
                    xmlOut.WriteElementString("value", "0.097")
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
                    xmlOut.WriteElementString("varname", "sdr")
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
                    xmlOut.WriteElementString("varname", "sdz")
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
                    xmlOut.WriteElementString("varname", "charlength")
                    xmlOut.WriteElementString("value", "15")
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

            NumSmokeDetectors = oSmokeDets.Count

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in SmokeDetDB.vb SaveSmokeDets")
            xmlOut.WriteEndDocument()
            xmlOut.Close()
        End Try
    End Sub
    Public Shared Function GetSDDistributions() As List(Of oDistribution)

        Dim oSmokeDets As New List(Of oSmokeDet)
        Dim osddistributions As New List(Of oDistribution)
        Dim osddistributions2 As New List(Of oDistribution)
        Dim Path As String = RiskDataDirectory & "smokedets.xml"
        Dim ver As Single
        Dim xmlIn As New XmlTextReader(Path)
        Dim vflag As Boolean = False
        Try

            If My.Computer.FileSystem.DirectoryExists(RiskDataDirectory) = True Then
                If My.Computer.FileSystem.FileExists(RiskDataDirectory & "smokedets.xml") Then

                    xmlIn.WhitespaceHandling = WhitespaceHandling.None

                    Do While xmlIn.Name <> "Smoke_Detector"
                        xmlIn.Read()
                        If xmlIn.Name = "version" Then
                            ver = CSng(xmlIn.ReadElementString("version"))
                        End If
                        If xmlIn.EOF = True Then 'no items
                            xmlIn.Close()
                            Return osddistributions
                            Exit Function
                        End If
                    Loop

                    Dim bx = 0
                    Do While xmlIn.Name = "Smoke_Detector"
                        Dim oSmokeDet As New oSmokeDet

                        xmlIn.ReadStartElement("Smoke_Detector")
                        bx = bx + 1
                        oSmokeDet.sdid = bx


                        Do While xmlIn.Name <> "sdistribution"
                            xmlIn.Read()
                            If xmlIn.EOF Then
                                MsgBox("distributions not found in smokedets.xml file")
                                xmlIn.Close()
                                Return osddistributions
                                Exit Function
                            End If
                        Loop

                        Do While xmlIn.Name = "sdistribution"
                            Dim odistribution As New oDistribution

                            xmlIn.ReadStartElement("sdistribution")
                            odistribution.id = oSmokeDet.sdid 'saves id
                            odistribution.varname = CStr(xmlIn.ReadElementString("varname"))
                            odistribution.varvalue = CStr(xmlIn.ReadElementString("value"))
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
                            osddistributions.Add(odistribution)
                        Loop

                        If osddistributions.Count Mod 3 = 0 And osddistributions.Last.varname = "OD" Then
                            vflag = True
                            For Each oDistribution In osddistributions
                                If oDistribution.id = oSmokeDet.sdid Then
                                    osddistributions2.Add(oDistribution)

                                    If oDistribution.varname = "OD" Then
                                        Dim x As New oDistribution
                                        x.id = oSmokeDet.sdid 'saves item id

                                        x.varname = CStr("charlength")
                                        x.varvalue = 1
                                        x.distribution = CStr("None")
                                        x.mean = 0
                                        x.variance = 0
                                        x.lbound = 0
                                        x.ubound = 0
                                        x.mode = 0
                                        x.alpha = 0
                                        x.beta = 0

                                        osddistributions2.Add(x)
                                    End If
                                End If
                            Next

                        End If

                        xmlIn.ReadEndElement()
                        oSmokeDets.Add(oSmokeDet)

                    Loop
                    If vflag = True Then osddistributions = osddistributions2
                End If
            End If

            xmlIn.Close()
            Return osddistributions

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in SmokeDetDB.vb GetSDDistributions")
            xmlIn.Close()
            Return osddistributions
        End Try

    End Function
End Class
