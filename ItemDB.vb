Imports System.Xml
Imports System.Xml.Linq
Imports System.Xml.Linq.XElement
Imports System.Collections.Generic

Public Class ItemDB
    Public Shared Function GetItemDistributions() As List(Of oDistribution)

        Dim oitems As New List(Of oItem)
        Dim oitemdistributions As New List(Of oDistribution)
        Dim Path As String = RiskDataDirectory & "items.xml"
        Dim ver As Single
        Dim xmlIn As New XmlTextReader(Path)

        Try

            If My.Computer.FileSystem.DirectoryExists(RiskDataDirectory) = True Then
                If My.Computer.FileSystem.FileExists(RiskDataDirectory & "items.xml") Then

                    xmlIn.WhitespaceHandling = WhitespaceHandling.None

                    Do While xmlIn.Name <> "Item"
                        xmlIn.Read()
                        If xmlIn.Name = "version" Then
                            ver = CSng(xmlIn.ReadElementString("version"))
                        End If
                        If xmlIn.EOF = True Then 'no items
                            xmlIn.Close()
                            Return oitemdistributions
                            Exit Function
                        End If
                    Loop


                    Do While xmlIn.Name = "Item"
                        Dim oitem As New oItem

                        xmlIn.ReadStartElement("Item")

                        oitem.id = CInt(xmlIn.ReadElementString("id"))


                        Do While xmlIn.Name <> "idistribution"
                            xmlIn.Read()
                            If xmlIn.EOF Then
                                MsgBox("distributions not found in items.xml file")

                                xmlIn.Close()

                                'if reading older file, add any missing distributions

                                'add heat of combustion
                                Dim odistribution As New oDistribution
                                odistribution.id = oitem.id
                                odistribution.varname = "heat of combustion"
                                odistribution.distribution = "None"
                                odistribution.varvalue = 20
                                oitemdistributions.Add(odistribution)

                                'add co2 yield
                                odistribution.id = oitem.id
                                odistribution.varname = "co2 yield"
                                odistribution.distribution = "None"
                                odistribution.varvalue = 1.27
                                oitemdistributions.Add(odistribution)

                                'add latent heat of gasification
                                odistribution.id = oitem.id
                                odistribution.varname = "Latent Heat of Gasification"
                                odistribution.distribution = "None"
                                odistribution.varvalue = 3
                                oitemdistributions.Add(odistribution)

                                'add radiant loss fraction
                                odistribution.id = oitem.id
                                odistribution.varname = "Radiant Loss Fraction"
                                odistribution.distribution = "None"
                                odistribution.varvalue = 0.35
                                oitemdistributions.Add(odistribution)

                                'hrrua
                                odistribution.id = oitem.id
                                odistribution.varname = "Heat Release Rate per Unit Area"
                                odistribution.distribution = "None"
                                odistribution.varvalue = 250
                                oitemdistributions.Add(odistribution)

                                xmlIn.Close()
                                Return oitemdistributions
                                Exit Function

                            End If
                        Loop

                        Dim name As String

                        Do While xmlIn.Name = "idistribution"
                            Dim odistribution As New oDistribution

                            xmlIn.ReadStartElement("idistribution")
                            odistribution.id = oitem.id 'saves item id
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
                            name = odistribution.varname
                            oitemdistributions.Add(odistribution)
                        Loop
                        xmlIn.ReadEndElement()

                        If name <> "HRRUA" Then
                            'add  hrrua
                            Dim odistribution As New oDistribution

                            odistribution.id = oitem.id 'saves item id
                            odistribution.varname = "HRRUA"
                            odistribution.distribution = "None"
                            odistribution.varvalue = ObjectMLUA(2, oitem.id - 1)
                            oitemdistributions.Add(odistribution)
                        End If

                        oitems.Add(oitem)

                    Loop
                End If
            End If

            xmlIn.Close()
            Return oitemdistributions

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in ItemDB.vb GetItemDistributions")
            xmlIn.Close()
            Return oitemdistributions
        End Try

    End Function

    Public Shared Function SaveWebItems(ByVal oitems As List(Of oItem), ByVal oitemdistributions As List(Of oDistribution))
        'these are distinct items listed and stored in the items.xml file.
        'for a given simulation these items can appear multiple times in a room
        'for a given simulation these items can appear multiple times in a room

        Dim Path As String = RiskDataDirectory & "webitems.xml"
        Dim xmlOut As New XmlTextWriter(Path, System.Text.Encoding.UTF8)

        Try
        
            xmlOut.Formatting = Formatting.Indented

            xmlOut.WriteStartDocument()

            xmlOut.WriteStartElement("Items")
            xmlOut.WriteElementString("version", Version.ToString)

            Dim oItem As oItem
            For i As Integer = 0 To oitems.Count - 1
                oItem = CType(oitems(i), oItem)
                xmlOut.WriteStartElement("Item")
                xmlOut.WriteElementString("id", CStr(i + 1))
                xmlOut.WriteElementString("description", oItem.description.ToString)
                xmlOut.WriteElementString("detaileddescription", oItem.detaileddescription.ToString)
                xmlOut.WriteElementString("userlabel", oItem.userlabel.ToString)
                xmlOut.WriteElementString("type", oItem.type.ToString)
                xmlOut.WriteElementString("length", oItem.length.ToString)
                xmlOut.WriteElementString("width", oItem.width.ToString)
                xmlOut.WriteElementString("height", oItem.height.ToString)
                xmlOut.WriteElementString("elevation", oItem.elevation.ToString)
                xmlOut.WriteElementString("mass", oItem.mass.ToString)

                'If oItem.critfluxdistribution = Nothing Then oItem.critfluxdistribution = "None"
                'If oItem.ftplimitpilotdistribution = Nothing Then oItem.ftplimitpilotdistribution = "None"
                'If oItem.ftpindexpilotdistribution = Nothing Then oItem.ftpindexpilotdistribution = "None"

                xmlOut.WriteElementString("critical_flux_pilot", oItem.critflux.ToString)
                xmlOut.WriteElementString("critical_flux_auto", oItem.critfluxauto.ToString)
                xmlOut.WriteElementString("FTP_limit_pilot", oItem.ftplimitpilot.ToString)
                xmlOut.WriteElementString("FTP_limit_auto", oItem.ftplimitauto.ToString)
                xmlOut.WriteElementString("FTP_index_pilot", oItem.ftpindexpilot.ToString)
                xmlOut.WriteElementString("FTP_index_auto", oItem.ftpindexauto.ToString)
                xmlOut.WriteElementString("probability", oItem.prob.ToString)
                xmlOut.WriteElementString("hrr", oItem.hrr.ToString)
                'xmlOut.WriteElementString("co2", oItem.co2.ToString)
                'xmlOut.WriteElementString("co", oItem.co.ToString)
                'xmlOut.WriteElementString("soot", oItem.soot.ToString)
                xmlOut.WriteElementString("ignition_time", oItem.ignitiontime.ToString)
                xmlOut.WriteElementString("max_num", oItem.maxnumitem.ToString)
                xmlOut.WriteElementString("xleft", oItem.xleft.ToString)
                xmlOut.WriteElementString("ybottom", oItem.ybottom.ToString)
                xmlOut.WriteElementString("radiantlossfraction", oItem.radlossfrac.ToString)
                xmlOut.WriteElementString("constantA", oItem.constantA.ToString)
                xmlOut.WriteElementString("constantB", oItem.constantB.ToString)
                xmlOut.WriteElementString("LHog", oItem.LHoG.ToString)

                For Each odistribution In oitemdistributions
                    If odistribution.id = i + 1 Then
                        xmlOut.WriteStartElement("idistribution")
                        'xmlOut.WriteElementString("id", CStr(i + 1))
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
                xmlOut.WriteEndElement()
            Next

            xmlOut.WriteEndElement()
            xmlOut.WriteEndDocument()
            xmlOut.Close()

            NumItems = oitems.Count

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in ItemDB.vb SaveWebItems")
            xmlOut.WriteEndDocument()
            xmlOut.Close()
        End Try
    End Function
    Public Shared Function GetItemsv2() As List(Of oItem)

        Dim oitems As New List(Of oItem)
        Dim oitemdistributions As New List(Of oDistribution)
        Dim Path As String = RiskDataDirectory & "items.xml"
        Dim ver As Single
        Dim xmlIn As New XmlTextReader(Path)
        Dim dummy As String

        Try

            If My.Computer.FileSystem.DirectoryExists(RiskDataDirectory) = True Then
                If My.Computer.FileSystem.FileExists(RiskDataDirectory & "items.xml") Then

                    xmlIn.WhitespaceHandling = WhitespaceHandling.None

                    Do While xmlIn.Name <> "Item"
                        xmlIn.Read()
                        If xmlIn.Name = "version" Then
                            ver = CSng(xmlIn.ReadElementString("version"))
                        End If
                        If xmlIn.EOF = True Then 'no items
                            xmlIn.Close()
                            Return oitems
                            Exit Function
                        End If
                    Loop

                    Do While xmlIn.Name = "Item"
                        Dim oitem As New oItem

                        xmlIn.ReadStartElement("Item")

                        oitem.id = CInt(xmlIn.ReadElementString("id"))
                        oitem.description = CStr(xmlIn.ReadElementString("description"))
                        If ver > CSng(2012.03) Then
                            oitem.detaileddescription = CStr(xmlIn.ReadElementString("detaileddescription"))
                            oitem.userlabel = CStr(xmlIn.ReadElementString("userlabel"))
                        End If
                        oitem.type = CStr(xmlIn.ReadElementString("type"))
                        oitem.length = CSng(xmlIn.ReadElementString("length"))
                        oitem.width = CSng(xmlIn.ReadElementString("width"))
                        oitem.height = CSng(xmlIn.ReadElementString("height"))
                        oitem.elevation = CSng(xmlIn.ReadElementString("elevation"))
                        oitem.mass = CSng(xmlIn.ReadElementString("mass"))

                        If ver < CSng(2011.17) Then
                            oitem.critflux = CSng(xmlIn.ReadElementString("critical_flux_pilot"))
                        ElseIf ver < CSng(2011.24) Then
                            xmlIn.ReadStartElement("critical_flux_pilot")
                            oitem.critflux = CSng(xmlIn.ReadElementString("value"))
                            dummy = CStr(xmlIn.ReadElementString("distribution"))
                            dummy = CSng(xmlIn.ReadElementString("meanmode"))
                            dummy = CSng(xmlIn.ReadElementString("variance"))
                            dummy = CSng(xmlIn.ReadElementString("lbound"))
                            dummy = CSng(xmlIn.ReadElementString("ubound"))
                            xmlIn.ReadEndElement()
                        Else
                            oitem.critflux = CSng(xmlIn.ReadElementString("critical_flux_pilot"))
                        End If

                        oitem.critfluxauto = CSng(xmlIn.ReadElementString("critical_flux_auto"))

                        If ver < CSng(2011.17) Then
                            oitem.ftplimitpilot = CSng(xmlIn.ReadElementString("FTP_limit_pilot"))
                        ElseIf ver < CSng(2011.24) Then
                            xmlIn.ReadStartElement("FTP_limit_pilot")
                            oitem.ftplimitpilot = CSng(xmlIn.ReadElementString("value"))
                            dummy = CStr(xmlIn.ReadElementString("distribution"))
                            dummy = CSng(xmlIn.ReadElementString("meanmode"))
                            dummy = CSng(xmlIn.ReadElementString("variance"))
                            dummy = CSng(xmlIn.ReadElementString("lbound"))
                            dummy = CSng(xmlIn.ReadElementString("ubound"))
                            xmlIn.ReadEndElement()
                        Else
                            oitem.ftplimitpilot = CSng(xmlIn.ReadElementString("FTP_limit_pilot"))
                        End If

                        oitem.ftplimitauto = CSng(xmlIn.ReadElementString("FTP_limit_auto"))

                        If ver < CSng(2011.17) Then
                            oitem.ftpindexpilot = CSng(xmlIn.ReadElementString("FTP_index_pilot"))
                        ElseIf ver < CSng(2011.24) Then

                            xmlIn.ReadStartElement("FTP_index_pilot")
                            oitem.ftpindexpilot = CSng(xmlIn.ReadElementString("value"))
                            dummy = CStr(xmlIn.ReadElementString("distribution"))
                            dummy = CSng(xmlIn.ReadElementString("meanmode"))
                            dummy = CSng(xmlIn.ReadElementString("variance"))
                            dummy = CSng(xmlIn.ReadElementString("lbound"))
                            dummy = CSng(xmlIn.ReadElementString("ubound"))
                            xmlIn.ReadEndElement()
                        Else
                            oitem.ftpindexpilot = CSng(xmlIn.ReadElementString("FTP_index_pilot"))
                        End If

                        oitem.ftpindexauto = CSng(xmlIn.ReadElementString("FTP_index_auto"))

                        oitem.prob = CSng(xmlIn.ReadElementString("probability"))
                        oitem.hrr = CStr(xmlIn.ReadElementString("hrr"))
                        If xmlIn.Name = "mlr" Then
                            oitem.mlrfreeburn = CStr(xmlIn.ReadElementString("mlr"))

                            If FuelResponseEffects = True Then
                                Dim linetext As String = CStr(oitem.mlrfreeburn)
                                linetext = CStr(linetext.Trim)
                                linetext = linetext.Replace(Chr(9), ",")
                                linetext = linetext.Replace(Chr(10), ",")
                                linetext = linetext.Replace(Chr(13), "")
                                linetext = linetext.Replace(" ", "")
                                linetext = linetext.Replace(",,", ",")
                                'remove the comma off the end
                                If linetext.Length > 0 Then
                                    If linetext.Chars(linetext.Length - 1) = "," Then linetext = linetext.Remove(linetext.Length - 1, 1)
                                End If
                                Dim j As Integer
                                Dim NumberPoints As Integer = 0
                                Dim str_data As String() = linetext.Split(CChar(","))
                                Dim i As Integer = str_data.Count - 1
                                If i > 0 Then
                                    j = 0
                                    Dim id As Integer = oitem.id
                                    NumberDataPoints(id) = 0
                                    Do While (NumberPoints < i And NumberPoints < 1000)
                                        NumberPoints = NumberPoints + 2
                                        If IsNumeric(str_data(NumberPoints - 2)) And IsNumeric(str_data(NumberPoints - 1)) Then
                                            MLRData(1, NumberDataPoints(id) + 1, id) = CDbl(str_data(NumberPoints - 2).Normalize) 'time
                                            MLRData(2, NumberDataPoints(id) + 1, id) = CDbl(str_data(NumberPoints - 1).Normalize) 'mlr
                                        End If
                                        j = j + 1
                                        NumberDataPoints(id) = CShort(NumberDataPoints(id) + 1)
                                    Loop
                                End If

                                'Else
                                'oitem.mlrfreeburn = "0,0"
                            End If
                        End If

                        oitem.ignitiontime = CSng(xmlIn.ReadElementString("ignition_time")) '??? not unique to item
                        oitem.maxnumitem = CInt(xmlIn.ReadElementString("max_num"))
                oItem.xleft = CSng(xmlIn.ReadElementString("xleft"))
                oItem.ybottom = CSng(xmlIn.ReadElementString("ybottom"))

                If xmlIn.Name = "radiantlossfraction" Then oItem.radlossfrac = CSng(xmlIn.ReadElementString("radiantlossfraction"))
                If xmlIn.Name = "constantA" Then oItem.constantA = CSng(xmlIn.ReadElementString("constantA"))
                If xmlIn.Name = "constantB" Then oItem.constantB = CSng(xmlIn.ReadElementString("constantB"))
                If xmlIn.Name = "LHoG" Then oItem.LHoG = CSng(xmlIn.ReadElementString("LHoG"))
                If xmlIn.Name = "HRRUA" Then oItem.HRRUA = CSng(xmlIn.ReadElementString("HRRUA"))
                        If xmlIn.Name = "windeffect" Then oitem.windeffect = CSng(xmlIn.ReadElementString("windeffect"))

                        If xmlIn.Name = "pyrolysisoption" Then oitem.pyrolysisoption = CInt(xmlIn.ReadElementString("pyrolysisoption"))

                        If ver > CSng(2016.02) Then

                            If xmlIn.Name = "pooldensity" Then oitem.pooldensity = CDbl(xmlIn.ReadElementString("pooldensity"))
                            If xmlIn.Name = "pooldiameter" Then oitem.pooldiameter = CDbl(xmlIn.ReadElementString("pooldiameter"))
                            If xmlIn.Name = "poolFBMLR" Then oitem.poolFBMLR = CDbl(xmlIn.ReadElementString("poolFBMLR"))
                            If xmlIn.Name = "poolramp" Then oitem.poolramp = CDbl(xmlIn.ReadElementString("poolramp"))
                            If xmlIn.Name = "poolvolume" Then oitem.poolvolume = CDbl(xmlIn.ReadElementString("poolvolume"))
                            If xmlIn.Name = "poolvaptemp" Then oitem.poolvaptemp = CDbl(xmlIn.ReadElementString("poolvaptemp"))
                        End If

                        If ver > CSng(2011.23) Then
                    Dim bx As Integer = 0
                    Do While xmlIn.Name = "idistribution"
                        Dim odistribution As New oDistribution
                        bx = bx + 1
                        xmlIn.ReadStartElement("idistribution")
                        odistribution.id = oItem.id 'saves item id
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
                        oitemdistributions.Add(odistribution)
                    Loop

                    'if reading older file, add any missing distributions
                    If bx = 3 Then
                        'add lhog
                        Dim oDistribution As New oDistribution
                        oDistribution.id = oItem.id
                        oDistribution.varname = "Latent Heat of Gasification"
                        oDistribution.distribution = "None"
                        oDistribution.units = "kJ/g"
                        oDistribution.varvalue = oItem.LHoG
                        oitemdistributions.Add(oDistribution)
                        bx = bx + 1
                    End If

                    If bx = 4 Then
                        'add rlf
                        Dim oDistribution As New oDistribution
                        oDistribution.id = oItem.id
                        oDistribution.varname = "Radiant Loss Fraction"
                        oDistribution.distribution = "None"
                        oDistribution.units = "-"
                        oDistribution.varvalue = oItem.radlossfrac
                        oitemdistributions.Add(oDistribution)
                        bx = bx + 1
                    End If

                    If bx = 5 Then
                        'add hrrua
                        Dim oDistribution As New oDistribution
                        oDistribution.id = oItem.id
                        oDistribution.varname = "HRRUA"
                        oDistribution.distribution = "None"
                        oDistribution.units = "kW/m2"
                        oDistribution.varvalue = oItem.HRRUA
                        oitemdistributions.Add(oDistribution)
                        bx = bx + 1
                    End If

                Else
                    Do While xmlIn.Name = "hoc"
                        Dim odistribution As New oDistribution

                        xmlIn.ReadStartElement("hoc")
                        odistribution.id = oItem.id 'saves item id
                        odistribution.varname = "heat of combustion"
                        odistribution.varvalue = CStr(xmlIn.ReadElementString("value"))
                        odistribution.distribution = CStr(xmlIn.ReadElementString("distribution"))
                        odistribution.mean = CStr(xmlIn.ReadElementString("meanmode"))
                        odistribution.variance = CSng(xmlIn.ReadElementString("variance"))
                        odistribution.lbound = CSng(xmlIn.ReadElementString("lbound"))
                        odistribution.ubound = CSng(xmlIn.ReadElementString("ubound"))
                        odistribution.mode = odistribution.mean
                        odistribution.alpha = 0
                        odistribution.beta = 0
                        xmlIn.ReadEndElement()
                        oitemdistributions.Add(odistribution)
                    Loop
                    Do While xmlIn.Name = "soot"
                        Dim odistribution As New oDistribution
                        xmlIn.ReadStartElement("soot")
                        odistribution.id = oItem.id 'saves item id
                        odistribution.varname = "soot yield"
                        odistribution.varvalue = CStr(xmlIn.ReadElementString("value"))
                        odistribution.distribution = CStr(xmlIn.ReadElementString("distribution"))
                        odistribution.mean = CStr(xmlIn.ReadElementString("meanmode"))
                        odistribution.variance = CSng(xmlIn.ReadElementString("variance"))
                        odistribution.lbound = CSng(xmlIn.ReadElementString("lbound"))
                        odistribution.ubound = CSng(xmlIn.ReadElementString("ubound"))
                        odistribution.mode = odistribution.mean
                        odistribution.alpha = 0
                        odistribution.beta = 0
                        xmlIn.ReadEndElement()
                        oitemdistributions.Add(odistribution)
                    Loop
                    Do While xmlIn.Name = "co2"
                        Dim odistribution As New oDistribution
                        xmlIn.ReadStartElement("co2")
                        odistribution.id = oItem.id 'saves item id
                        odistribution.varname = "co2 yield"
                        odistribution.varvalue = CStr(xmlIn.ReadElementString("value"))
                        odistribution.distribution = CStr(xmlIn.ReadElementString("distribution"))
                        odistribution.mean = CStr(xmlIn.ReadElementString("meanmode"))
                        odistribution.variance = CSng(xmlIn.ReadElementString("variance"))
                        odistribution.lbound = CSng(xmlIn.ReadElementString("lbound"))
                        odistribution.ubound = CSng(xmlIn.ReadElementString("ubound"))
                        odistribution.mode = odistribution.mean
                        odistribution.alpha = 0
                        odistribution.beta = 0
                        xmlIn.ReadEndElement()
                        oitemdistributions.Add(odistribution)
                    Loop
                End If


                xmlIn.ReadEndElement()
                oitems.Add(oItem)

                    Loop
                End If
            End If

            xmlIn.Close()

            If ver > CSng(2011.24) Then ItemDB.SaveItemsv2(oitems, oitemdistributions)
            If oitems.Count > 0 Then
                frmInputs.StartToolStripLabel1.Visible = True
            Else
                frmInputs.StartToolStripLabel1.Visible = False
            End If

            Return oitems

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in ItemDB.vb GetItemsv2")
            xmlIn.Close()
            Return oitems
        End Try

    End Function

    Public Shared Sub SaveItemsv2(ByVal oitems As List(Of oItem), ByVal oitemdistributions As List(Of oDistribution))
        'these are distinct items listed and stored in the items.xml file.
        'for a given simulation these items can appear multiple times in a room


        Dim Path As String = RiskDataDirectory & "items.xml"
        Dim xmlOut As New XmlTextWriter(Path, System.Text.Encoding.UTF8)

        Try
            xmlOut.Formatting = Formatting.Indented

            xmlOut.WriteStartDocument()

            xmlOut.WriteStartElement("Items")
            xmlOut.WriteElementString("version", Version.ToString)

            Dim oItem As oItem
            For i As Integer = 0 To oitems.Count - 1
                oItem = CType(oitems(i), oItem)
                xmlOut.WriteStartElement("Item")
                xmlOut.WriteElementString("id", CStr(i + 1))
                xmlOut.WriteElementString("description", oItem.description.ToString)
                If oItem.detaileddescription IsNot Nothing Then
                    xmlOut.WriteElementString("detaileddescription", oItem.detaileddescription.ToString)
                Else
                    xmlOut.WriteElementString("detaileddescription", "")
                End If
                If oItem.userlabel IsNot Nothing Then
                    xmlOut.WriteElementString("userlabel", oItem.userlabel.ToString)
                Else
                    xmlOut.WriteElementString("userlabel", "")
                End If

                xmlOut.WriteElementString("type", oItem.type.ToString)
                xmlOut.WriteElementString("length", oItem.length.ToString)
                xmlOut.WriteElementString("width", oItem.width.ToString)
                xmlOut.WriteElementString("height", oItem.height.ToString)
                xmlOut.WriteElementString("elevation", oItem.elevation.ToString)
                xmlOut.WriteElementString("mass", oItem.mass.ToString)

                xmlOut.WriteElementString("critical_flux_pilot", oItem.critflux.ToString)
                xmlOut.WriteElementString("critical_flux_auto", oItem.critfluxauto.ToString)
                xmlOut.WriteElementString("FTP_limit_pilot", oItem.ftplimitpilot.ToString)
                xmlOut.WriteElementString("FTP_limit_auto", oItem.ftplimitauto.ToString)
                xmlOut.WriteElementString("FTP_index_pilot", oItem.ftpindexpilot.ToString)
                xmlOut.WriteElementString("FTP_index_auto", oItem.ftpindexauto.ToString)
                xmlOut.WriteElementString("probability", oItem.prob.ToString)
                xmlOut.WriteElementString("hrr", oItem.hrr.ToString)
                If IsNothing(oItem.mlrfreeburn) = False Then
                    xmlOut.WriteElementString("mlr", oItem.mlrfreeburn.ToString)
                End If
                xmlOut.WriteElementString("ignition_time", oItem.ignitiontime.ToString)
                xmlOut.WriteElementString("max_num", oItem.maxnumitem.ToString)
                xmlOut.WriteElementString("xleft", oItem.xleft.ToString)
                xmlOut.WriteElementString("ybottom", oItem.ybottom.ToString)
                xmlOut.WriteElementString("constantA", oItem.constantA.ToString)
                xmlOut.WriteElementString("constantB", oItem.constantB.ToString)
                xmlOut.WriteElementString("windeffect", oItem.windeffect.ToString)
                xmlOut.WriteElementString("pyrolysisoption", oItem.pyrolysisoption.ToString)
                xmlOut.WriteElementString("pooldensity", oItem.pooldensity.ToString)
                xmlOut.WriteElementString("pooldiameter", oItem.pooldiameter.ToString)
                xmlOut.WriteElementString("poolFBMLR", oItem.poolFBMLR.ToString)
                xmlOut.WriteElementString("poolramp", oItem.poolramp.ToString)
                xmlOut.WriteElementString("poolvolume", oItem.poolvolume.ToString)
                xmlOut.WriteElementString("poolvaptemp", oItem.poolvaptemp.ToString)

                For Each odistribution In oitemdistributions
                    If odistribution.id = i + 1 Then
                        xmlOut.WriteStartElement("idistribution")
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

                xmlOut.WriteEndElement()
            Next

            xmlOut.WriteEndElement()
            xmlOut.WriteEndDocument()
            xmlOut.Close()

            NumItems = oitems.Count

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in ItemDB.vb SaveItemsv2")
            xmlOut.WriteEndDocument()
            xmlOut.Close()
        End Try
    End Sub

    Public Sub New()

    End Sub
End Class
