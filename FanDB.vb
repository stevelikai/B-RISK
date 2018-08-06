Imports System.Xml
Imports System.Collections.Generic

Public Class FanDB

    Public Shared Function RemoveFans(ByVal oFans As List(Of oFan), ByVal ofandistributions As List(Of oDistribution))

        oFans = FanDB.GetFans()

        Dim icount As Integer = oFans.Count

        Dim j As Integer = 0
        If icount > 0 Then

            oFans.Clear()

            ofandistributions = FanDB.GetFanDistributions()
            ofandistributions.Clear()

            FanDB.SaveFans(oFans, ofandistributions)

        End If

        Return oFans

    End Function

    Public Shared Function GetFans() As List(Of oFan)

        'get fan details from fans.xml file
        Dim oFans As New List(Of oFan)
        Dim oFandistributions As New List(Of oDistribution)

        If IsNothing(ProjectDirectory) Then ProjectDirectory = RiskDataDirectory

        Dim Path As String = ProjectDirectory & "fans.xml"
        Dim xmlIn As New XmlTextReader(Path)

        If My.Computer.FileSystem.FileExists(Path) = False Then
            MsgBox("The file fans.xml is missing. A new file will be a created, please check your fan inputs.", MsgBoxStyle.Critical, "Missing File")

            'My.Computer.FileSystem.CopyFile(UserAppDataFolder & gcs_folder_ext & "\" & "riskdata\basemodel_default\" & "fans.xml", RiskDataDirectory & "fans.xml", True)
            My.Computer.FileSystem.CopyFile(DataFolder & "fans.xml", RiskDataDirectory & "fans.xml", True)

        End If

        If My.Computer.FileSystem.FileExists(Path) Then

            xmlIn.WhitespaceHandling = WhitespaceHandling.None
            Dim oldfile As Boolean = True
            Dim fanver As Single

            Do While xmlIn.Name <> "Fan"
                xmlIn.Read()
                If xmlIn.EOF = True Then 'no fan
                    xmlIn.Close()
                    Return oFans
                    Exit Function
                End If
                If xmlIn.Name = "version" Then
                    oldfile = False
                    fanver = CSng(xmlIn.ReadElementString("version"))
                End If
            Loop
            If oldfile = True Then
                MsgBox("Your fans.xml file is obsolete. Please delete it and try again.", MsgBoxStyle.Exclamation)
                xmlIn.Close()
                Return oFans
                Exit Function
            End If

            Do While xmlIn.Name = "Fan"
                Dim oFan As New oFan
                xmlIn.ReadStartElement("Fan")
                oFan.fanid = CInt(xmlIn.ReadElementString("fanid"))
                oFan.fanroom = CInt(xmlIn.ReadElementString("fanroom"))
                oFan.fanelevation = CSng(xmlIn.ReadElementString("fanelevation"))
                oFan.fanextract = CBool(xmlIn.ReadElementString("fanextract"))

                Dim dummy As String = xmlIn.ReadElementString("fanmanual")
                If dummy = "True" Then
                    oFan.fanmanual = 0
                ElseIf dummy = "False" Then
                    oFan.fanmanual = 1
                Else
                    oFan.fanmanual = CInt(dummy)
                End If

                oFan.fancurve = CBool(xmlIn.ReadElementString("fancurve"))


                Dim bx As Integer = 0
                Do While xmlIn.Name = "fdistribution"
                    Dim odistribution As New oDistribution
                    bx = bx + 1
                    xmlIn.ReadStartElement("fdistribution")
                    odistribution.id = oFan.fanid 'saves id
                    odistribution.varname = CStr(xmlIn.ReadElementString("varname"))
                    odistribution.varvalue = CStr(xmlIn.ReadElementString("value"))

                    If odistribution.varname = "fanflowrate" Then oFan.fanflowrate = odistribution.varvalue
                    If odistribution.varname = "fanstarttime" Then oFan.fanstarttime = odistribution.varvalue
                    If odistribution.varname = "fanpressurelimit" Then oFan.fanpressurelimit = odistribution.varvalue
                    If odistribution.varname = "fanreliability" Then oFan.fanreliability = odistribution.varvalue

                    odistribution.distribution = CStr(xmlIn.ReadElementString("distribution"))
                    odistribution.mean = CStr(xmlIn.ReadElementString("mean"))
                    odistribution.variance = CSng(xmlIn.ReadElementString("variance"))
                    odistribution.lbound = CSng(xmlIn.ReadElementString("lbound"))
                    odistribution.ubound = CSng(xmlIn.ReadElementString("ubound"))
                    odistribution.mode = CSng(xmlIn.ReadElementString("mode"))
                    odistribution.alpha = CSng(xmlIn.ReadElementString("alpha"))
                    odistribution.beta = CSng(xmlIn.ReadElementString("beta"))
                    xmlIn.ReadEndElement()
                    oFandistributions.Add(odistribution)
                Loop

                xmlIn.ReadEndElement()

                oFans.Add(oFan)

            Loop
        End If

        xmlIn.Close()
        Return oFans

    End Function

    Public Shared Sub SaveFans(ByVal oFans As List(Of oFan), ByVal ofandistributions As List(Of oDistribution))

        Dim rname As String = ProjectDirectory & "fans.xml"
        Dim Path As String = rname
        Dim xmlOut As New XmlTextWriter(Path, System.Text.Encoding.UTF8)

        Try

            xmlOut.Formatting = Formatting.Indented

            xmlOut.WriteStartDocument()
            xmlOut.WriteStartElement("Fans")
            xmlOut.WriteElementString("version", Version.ToString)

            Dim oFan As oFan
            For i As Integer = 0 To oFans.Count - 1
                oFan = CType(oFans(i), oFan)
                xmlOut.WriteStartElement("Fan")
                xmlOut.WriteElementString("fanid", CStr(i + 1))
                xmlOut.WriteElementString("fanroom", oFan.fanroom.ToString)
                xmlOut.WriteElementString("fanelevation", oFan.fanelevation.ToString)
                xmlOut.WriteElementString("fanextract", oFan.fanextract.ToString)
                xmlOut.WriteElementString("fanmanual", oFan.fanmanual.ToString)
                xmlOut.WriteElementString("fancurve", oFan.fancurve.ToString)

                If Not IsNothing(ofandistributions) Then
                    For Each odistribution In ofandistributions
                        If odistribution.id = i + 1 Then
                            xmlOut.WriteStartElement("fdistribution")
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
                    xmlOut.WriteStartElement("fdistribution")
                    xmlOut.WriteElementString("varname", "fanflowrate")
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

                    xmlOut.WriteStartElement("fdistribution")
                    xmlOut.WriteElementString("varname", "fanstarttime")
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

                    xmlOut.WriteStartElement("fdistribution")
                    xmlOut.WriteElementString("varname", "fanpressurelimit")
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

                    xmlOut.WriteStartElement("fdistribution")
                    xmlOut.WriteElementString("varname", "fanreliability")
                    xmlOut.WriteElementString("value", "1")
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

            NumFans = oFans.Count

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in FanDB.vb SaveFans")
            xmlOut.WriteEndDocument()
            xmlOut.Close()
        End Try
    End Sub
    Public Shared Function GetFanDistributions() As List(Of oDistribution)

        Dim ofans As New List(Of oFan)
        Dim ofandistributions As New List(Of oDistribution)
        Dim ofandistributions2 As New List(Of oDistribution)
        Dim Path As String = RiskDataDirectory & "fans.xml"
        Dim ver As Single
        Dim xmlIn As New XmlTextReader(Path)
        Dim vflag As Boolean = False
        Try

            If My.Computer.FileSystem.DirectoryExists(RiskDataDirectory) = True Then
                If My.Computer.FileSystem.FileExists(RiskDataDirectory & "fans.xml") Then

                    xmlIn.WhitespaceHandling = WhitespaceHandling.None

                    Do While xmlIn.Name <> "Fan"
                        xmlIn.Read()
                        If xmlIn.Name = "version" Then
                            ver = CSng(xmlIn.ReadElementString("version"))
                        End If
                        If xmlIn.EOF = True Then 'no fans
                            xmlIn.Close()
                            Return ofandistributions
                            Exit Function
                        End If
                    Loop

                    Dim bx = 0
                    Do While xmlIn.Name = "Fan"
                        Dim oFan As New oFan

                        xmlIn.ReadStartElement("Fan")
                        bx = bx + 1
                        oFan.fanid = bx


                        Do While xmlIn.Name <> "fdistribution"
                            xmlIn.Read()
                            If xmlIn.EOF Then
                                MsgBox("distributions not found in fans.xml file")
                                xmlIn.Close()
                                Return ofandistributions
                                Exit Function
                            End If
                        Loop

                        Do While xmlIn.Name = "fdistribution"
                            Dim odistribution As New oDistribution

                            xmlIn.ReadStartElement("fdistribution")
                            odistribution.id = oFan.fanid 'saves id
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
                            ofandistributions.Add(odistribution)
                        Loop

                        'If ofandistributions.Count Mod 3 = 0 And ofandistributions.Last.varname = "fanflowrate" Then
                        '    vflag = True
                        '    For Each oDistribution In ofandistributions
                        '        If oDistribution.id = oFan.fanid Then
                        '            ofandistributions2.Add(oDistribution)

                        '            If oDistribution.varname = "fanflowrate" Then
                        '                Dim x As New oDistribution
                        '                x.id = oFan.fanid 'saves item id

                        '                x.varname = CStr("charlength")
                        '                x.varvalue = 1
                        '                x.distribution = CStr("None")
                        '                x.mean = 0
                        '                x.variance = 0
                        '                x.lbound = 0
                        '                x.ubound = 0
                        '                x.mode = 0
                        '                x.alpha = 0
                        '                x.beta = 0

                        '                osddistributions2.Add(x)
                        '            End If
                        '        End If
                        '    Next

                        'End If

                        xmlIn.ReadEndElement()
                        ofans.Add(oFan)

                    Loop
                    'If vflag = True Then osddistributions = osddistributions2
                End If
            End If

            xmlIn.Close()
            Return ofandistributions

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in FanDB.vb GetFanDistributions")
            xmlIn.Close()
            Return ofandistributions
        End Try

    End Function
End Class
