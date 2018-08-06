Imports System.Xml
Imports System.Collections.Generic
'Imports System.IO.FileStream

Public Class VentDB
    Public Shared Sub SaveVents(ByVal ovents As List(Of oVent), ByVal oventdistributions As List(Of oDistribution))
        'these are distinct items listed and stored in the vents.xml file.

        Dim ch As Boolean

        If IsNothing(ovents) Then Exit Sub

        Dim Path As String = RiskDataDirectory & "vents.xml"
        If Not My.Computer.FileSystem.DirectoryExists(RiskDataDirectory) Then
            My.Computer.FileSystem.CreateDirectory(RiskDataDirectory)
        End If

        ch = frmInputs.myFileInUse(RiskDataDirectory & "vents.xml")
        Do While ch = True
            ch = frmInputs.myFileInUse(RiskDataDirectory & "vents.xml")
            frmInputs.ToolStripStatusLabel3.Text = "Waiting ... while vents.xml is in use."
        Loop
        frmInputs.ToolStripStatusLabel3.Text = ""

        Dim xmlOut As New XmlTextWriter(Path, System.Text.Encoding.UTF8)
        Try
            xmlOut.Formatting = Formatting.Indented
            xmlOut.WriteStartDocument()

            xmlOut.WriteStartElement("Vents")
            xmlOut.WriteElementString("version", Version.ToString)
            xmlOut.WriteElementString("ENZ_simtime", enzsimtime.ToString)
            xmlOut.WriteElementString("BRISK_simtime", brisksimtime.ToString)

            Dim oVent As oVent
            For i As Integer = 0 To ovents.Count - 1
                oVent = CType(ovents(i), oVent)
                'If oVent.fromroom > NumberRooms Then
                '    Continue For
                'End If

                xmlOut.WriteStartElement("Vent")
                xmlOut.WriteElementString("id", CStr(i + 1))
                xmlOut.WriteElementString("description", oVent.description)
                xmlOut.WriteElementString("fromroom", oVent.fromroom.ToString)
                xmlOut.WriteElementString("toroom", oVent.toroom.ToString)
                xmlOut.WriteElementString("offset", oVent.offset.ToString)
                xmlOut.WriteElementString("face", oVent.face.ToString)
                xmlOut.WriteElementString("sillheight", oVent.sillheight.ToString)
                xmlOut.WriteElementString("opentime", oVent.opentime.ToString)
                xmlOut.WriteElementString("closetime", oVent.closetime.ToString)
                xmlOut.WriteElementString("walllength1", oVent.walllength1.ToString)
                xmlOut.WriteElementString("walllength2", oVent.walllength2.ToString)

                xmlOut.WriteElementString("sdtriggerroom", oVent.sdtriggerroom.ToString)
                xmlOut.WriteElementString("HRRthreshold", oVent.HRRthreshold.ToString)
                xmlOut.WriteElementString("HRRventopendelay", oVent.HRRventopendelay.ToString)
                xmlOut.WriteElementString("HRRventopenduration", oVent.HRRventopenduration.ToString)
                xmlOut.WriteElementString("triggerventopendelay", oVent.triggerventopendelay.ToString)
                xmlOut.WriteElementString("triggerventopenduration", oVent.triggerventopenduration.ToString)
                xmlOut.WriteElementString("autoopenvent", oVent.autoopenvent.ToString)

                xmlOut.WriteElementString("triggerHD", oVent.triggerHD.ToString)
                xmlOut.WriteElementString("triggerHRR", oVent.triggerHRR.ToString)
                xmlOut.WriteElementString("triggerSD", oVent.triggerSD.ToString)
                xmlOut.WriteElementString("triggerFO", oVent.triggerFO.ToString)
                xmlOut.WriteElementString("triggerVL", oVent.triggerVL.ToString)
                xmlOut.WriteElementString("triggerHO", oVent.triggerHO.ToString)
                xmlOut.WriteElementString("triggerFR", oVent.triggerFR.ToString)

                xmlOut.WriteElementString("autobreakglass", oVent.autobreakglass.ToString)
                xmlOut.WriteElementString("glassconductivity", oVent.glassconductivity.ToString)
                xmlOut.WriteElementString("glassemissivity", oVent.glassemissivity.ToString)
                xmlOut.WriteElementString("glassexpansion", oVent.glassexpansion.ToString)
                xmlOut.WriteElementString("glassthickness", oVent.glassthickness.ToString)
                xmlOut.WriteElementString("glassshading", oVent.glassshading.ToString)
                xmlOut.WriteElementString("glassbreakingstress", oVent.glassbreakingstress.ToString)
                xmlOut.WriteElementString("glassalpha", oVent.glassalpha.ToString)
                xmlOut.WriteElementString("glassYoungsModulus", oVent.glassYoungsModulus.ToString)
                xmlOut.WriteElementString("glassflameflux", oVent.glassflameflux.ToString)
                xmlOut.WriteElementString("glassdistance", oVent.glassdistance.ToString)
                xmlOut.WriteElementString("glassfalloutime", oVent.glassfalloutime.ToString)

                'xmlOut.WriteElementString("spillplume", oVent.spillplume.ToString)
                xmlOut.WriteElementString("spillplumemodel", oVent.spillplumemodel.ToString)
                'xmlOut.WriteElementString("spillplumebalc", oVent.spillplumebalc.ToString)
                xmlOut.WriteElementString("downstand", oVent.downstand.ToString)
                xmlOut.WriteElementString("spillbalconyprojection", oVent.spillbalconyprojection.ToString)

                xmlOut.WriteElementString("cd", oVent.cd.ToString)
                xmlOut.WriteElementString("probventclosed", oVent.probventclosed.ToString)
                xmlOut.WriteElementString("FRcriteria", oVent.FRcriteria.ToString)
                xmlOut.WriteElementString("ENZ_open", oVent.enzopen.ToString)

                Dim bx As Integer = 0
                For Each odistribution In oventdistributions
                    bx = bx + 1

                    If odistribution.id = i + 1 Then
                        xmlOut.WriteStartElement("vdistribution")
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

            NumVents = ovents.Count

        Catch ex As Exception
            If TalkToEVACNZ = False Then MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in VentDB.vb SaveVents")
        Finally
            xmlOut.WriteEndDocument()
            xmlOut.Close()
        End Try

    End Sub
    Public Shared Function GetISO9705Vents() As List(Of oVent)
        Dim ovents As New List(Of oVent)
        Dim oventdistributions As New List(Of oDistribution)
        Dim ovent As oVent

        Try

            ovents = VentDB.GetVents
            oventdistributions = VentDB.GetVentDistributions()
here:
            For Each x In oventdistributions
                oventdistributions.Remove(x)
                GoTo here
            Next
            For Each ovent In ovents
                ovents.Remove(ovent)
                GoTo here
            Next

            ovent = New oVent(0, 200, False, 10, 100, 0, 1, 1, False, False, False, 0, gcs_DischargeCoeff, False, False, False, False, 0, 0, 0, 0, 500, 1, False, 0, False, 0, 0, 0, False, 72000, 0.00000042, 47, 20, 6, 0.0000083, 1, 0.937, "Vent 1", 1, 2, 1, 2, 0.8, 0, 0, 0, 0, 0, False, False)
            ovent.id = 1

            ovents.Add(ovent)
            number_vents = 1
            NumberVents(1, 2) = number_vents
            NumberVents(2, 1) = number_vents
            Resize_Vents()

            Dim odistribution As New oDistribution("", "", "None", 1, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "height"
            oventdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 1, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "width"
            oventdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "prob"
            oventdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 1, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "HOreliability"
            oventdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "integrity"
            oventdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 100, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "maxopening"
            oventdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 10, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "maxopeningtime"
            oventdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 200, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "gastemp"
            oventdistributions.Add(odistribution)

            VentDB.SaveVents(ovents, oventdistributions)

            Return ovents

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in VentDB.vb GetISO9705Vents")
            Return ovents
        End Try
    End Function
    Public Shared Function GetVents() As List(Of oVent)

        Dim ovents As New List(Of oVent)
        Dim oventdistributions As New List(Of oDistribution)
        Dim Path As String = RiskDataDirectory & "vents.xml"
        Dim ver As Single
        Dim xmlIn As New XmlTextReader(Path)
        Dim ch As Boolean

        Try

            If My.Computer.FileSystem.DirectoryExists(RiskDataDirectory) = True Then
                If My.Computer.FileSystem.FileExists(RiskDataDirectory & "vents.xml") Then

                    ch = frmInputs.myFileInUse(RiskDataDirectory & "vents.xml")
                    Do While ch = True
                        ch = frmInputs.myFileInUse(RiskDataDirectory & "vents.xml")
                        frmInputs.ToolStripStatusLabel3.Text = "Waiting ... while vents.xml is in use."
                    Loop
                    frmInputs.ToolStripStatusLabel3.Text = ""

                    xmlIn.WhitespaceHandling = WhitespaceHandling.None

                    Do While xmlIn.Name <> "Vent"
                        xmlIn.Read()
                        If xmlIn.Name = "version" Then
                            ver = CSng(xmlIn.ReadElementString("version"))
                        End If
                        'If xmlIn.Name = "ENZ_simtime" Then
                        '    enzsimtime = CDbl(xmlIn.ReadElementString("ENZ_simtime"))
                        'End If

                        If xmlIn.EOF = True Then 'no vents
                            xmlIn.Close()
                            Return ovents
                            Exit Function
                        End If
                    Loop

                    Do While xmlIn.Name = "Vent"
                        Dim ovent As New oVent

                        xmlIn.ReadStartElement("Vent")

                        ovent.id = CInt(xmlIn.ReadElementString("id"))
                        ovent.description = CStr(xmlIn.ReadElementString("description"))
                        ovent.fromroom = CStr(xmlIn.ReadElementString("fromroom"))
                        ovent.toroom = CStr(xmlIn.ReadElementString("toroom"))
                        If ver > CSng(2012.42) Then
                            ovent.offset = CDbl(xmlIn.ReadElementString("offset"))
                        End If
                        ovent.face = CInt(xmlIn.ReadElementString("face"))
                        ovent.sillheight = CStr(xmlIn.ReadElementString("sillheight"))
                        ovent.opentime = CStr(xmlIn.ReadElementString("opentime"))
                        ovent.closetime = CStr(xmlIn.ReadElementString("closetime"))

                        If ver > CSng(2012.15) Then
                            ovent.walllength1 = CStr(xmlIn.ReadElementString("walllength1"))
                            ovent.walllength2 = CStr(xmlIn.ReadElementString("walllength2"))

                            ovent.sdtriggerroom = CStr(xmlIn.ReadElementString("sdtriggerroom"))
                            ovent.HRRthreshold = CStr(xmlIn.ReadElementString("HRRthreshold"))
                            ovent.HRRventopendelay = CStr(xmlIn.ReadElementString("HRRventopendelay"))
                            ovent.HRRventopenduration = CStr(xmlIn.ReadElementString("HRRventopenduration"))
                            ovent.triggerventopendelay = CStr(xmlIn.ReadElementString("triggerventopendelay"))
                            ovent.triggerventopenduration = CStr(xmlIn.ReadElementString("triggerventopenduration"))
                            ovent.autoopenvent = CStr(xmlIn.ReadElementString("autoopenvent"))

                            If ver > CSng(2012.19) Then
                                ovent.triggerHD = CStr(xmlIn.ReadElementString("triggerHD"))
                                ovent.triggerHRR = CStr(xmlIn.ReadElementString("triggerHRR"))
                                ovent.triggerSD = CStr(xmlIn.ReadElementString("triggerSD"))
                                ovent.triggerFO = CStr(xmlIn.ReadElementString("triggerFO"))
                            End If

                            If ver > CSng(2012.2) Then
                                ovent.triggerVL = CStr(xmlIn.ReadElementString("triggerVL"))
                            End If

                            If ver > CSng(2012.38) Then
                                ovent.triggerHO = CStr(xmlIn.ReadElementString("triggerHO"))
                            End If
                            If ver > CSng(2013.08) Then
                                ovent.triggerFR = CStr(xmlIn.ReadElementString("triggerFR"))
                            End If
                            ovent.autobreakglass = CStr(xmlIn.ReadElementString("autobreakglass"))
                            ovent.glassconductivity = CStr(xmlIn.ReadElementString("glassconductivity"))
                            ovent.glassemissivity = CStr(xmlIn.ReadElementString("glassemissivity"))
                            ovent.glassexpansion = CStr(xmlIn.ReadElementString("glassexpansion"))
                            ovent.glassthickness = CStr(xmlIn.ReadElementString("glassthickness"))
                            ovent.glassshading = CStr(xmlIn.ReadElementString("glassshading"))
                            ovent.glassbreakingstress = CStr(xmlIn.ReadElementString("glassbreakingstress"))
                            ovent.glassalpha = CStr(xmlIn.ReadElementString("glassalpha"))
                            ovent.glassYoungsModulus = CStr(xmlIn.ReadElementString("glassYoungsModulus"))
                            ovent.glassflameflux = CStr(xmlIn.ReadElementString("glassflameflux"))
                            ovent.glassdistance = CStr(xmlIn.ReadElementString("glassdistance"))
                            ovent.glassfalloutime = CStr(xmlIn.ReadElementString("glassfalloutime"))
                            If ver <= CSng(2012.23) Then
                                ovent.spillplume = CStr(xmlIn.ReadElementString("spillplume"))
                            End If
                            ovent.spillplumemodel = CStr(xmlIn.ReadElementString("spillplumemodel"))
                            If ver <= CSng(2012.23) Then
                                ovent.spillplumebalc = CStr(xmlIn.ReadElementString("spillplumebalc"))
                            End If
                            ovent.downstand = CStr(xmlIn.ReadElementString("downstand"))
                            If ver > CSng(2012.23) Then
                                ovent.spillbalconyprojection = CStr(xmlIn.ReadElementString("spillbalconyprojection"))
                            Else
                                Dim dummy As Boolean = xmlIn.ReadElementString("spillplumesingle")
                            End If

                        End If
                        If ver > CSng(2012.16) Then
                            ovent.cd = CStr(xmlIn.ReadElementString("cd"))
                            ovent.probventclosed = CStr(xmlIn.ReadElementString("probventclosed"))
                        End If

                        Try
                            If ver > CSng(2014.15) Then
                                ovent.FRcriteria = CStr(xmlIn.ReadElementString("FRcriteria"))
                            End If
                        Catch ex As Exception
                            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in VentDB.vb GetVents. FR criteria not found.")
                        End Try
                        
                        Try
                            If ver > CSng(2015.03) Then
                                'ovent.enzopen = CStr(xmlIn.ReadElementString("ENZ_open"))
                                Dim dummy As Boolean = CStr(xmlIn.ReadElementString("ENZ_open"))
                            End If
                        Catch ex As Exception
                            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in VentDB.vb GetVents. EvacuatioNZ data not found.")
                        End Try

                        Dim bx As Integer = 0

                        Do While xmlIn.Name = "vdistribution"
                            Dim odistribution As New oDistribution
                            bx = bx + 1
                            xmlIn.ReadStartElement("vdistribution")
                            odistribution.id = ovent.id 'saves item id
                            odistribution.varname = CStr(xmlIn.ReadElementString("varname"))
                            odistribution.varvalue = xmlIn.ReadElementString("value")
                            odistribution.distribution = xmlIn.ReadElementString("distribution")
                            odistribution.mean = xmlIn.ReadElementString("mean")
                            odistribution.variance = xmlIn.ReadElementString("variance")
                            odistribution.lbound = xmlIn.ReadElementString("lbound")
                            odistribution.ubound = xmlIn.ReadElementString("ubound")
                            odistribution.mode = xmlIn.ReadElementString("mode")
                            odistribution.alpha = xmlIn.ReadElementString("alpha")
                            odistribution.beta = xmlIn.ReadElementString("beta")
                            xmlIn.ReadEndElement()

                            If odistribution.varname = "height" Then ovent.height = odistribution.varvalue
                            If odistribution.varname = "width" Then ovent.width = odistribution.varvalue
                            If odistribution.varname = "prob" Then ovent.probventclosed = odistribution.varvalue
                            If odistribution.varname = "HOreliability" Then ovent.horeliability = odistribution.varvalue
                            If odistribution.varname = "integrity" Then ovent.integrity = odistribution.varvalue
                            If odistribution.varname = "maxopening" Then ovent.maxopening = odistribution.varvalue
                            If odistribution.varname = "maxopeningtime" Then ovent.maxopeningtime = odistribution.varvalue
                            If odistribution.varname = "gastemp" Then ovent.gastemp = odistribution.varvalue

                            oventdistributions.Add(odistribution)
                        Loop

                        ovents.Add(ovent)
                        xmlIn.ReadEndElement()

                    Loop
                End If
            End If

        Catch ex As Exception
            If TalkToEVACNZ = False Then MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in VentDB.vb GetVents")
        Finally
            xmlIn.Close()
        End Try

        Try
            VentDB.SaveVents(ovents, oventdistributions)
        Catch ex As Exception
            If TalkToEVACNZ = False Then MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in VentDB.vb GetVents")
        End Try
       
        Return ovents

    End Function
    Public Shared Function GetVentDistributions() As List(Of oDistribution)

        Dim ovents As New List(Of oVent)
        Dim oventdistributions As New List(Of oDistribution)
        Dim oventdistributions2 As New List(Of oDistribution)
        Dim Path As String = RiskDataDirectory & "vents.xml"
        Dim ver As Single
        Dim xmlIn As New XmlTextReader(Path)
        Dim vflag As Boolean = False
        Dim ch As Boolean

        Try
            'Throw New Exception("test")

            If My.Computer.FileSystem.DirectoryExists(RiskDataDirectory) = True Then
                If My.Computer.FileSystem.FileExists(RiskDataDirectory & "vents.xml") Then

                    Do While ch = True
                        ch = frmInputs.myFileInUse(RiskDataDirectory & "vents.xml")
                        frmInputs.ToolStripStatusLabel3.Text = "Waiting ... while vents.xml is in use."
                    Loop

                    xmlIn.WhitespaceHandling = WhitespaceHandling.None

                    Do While xmlIn.Name <> "Vent"
                        xmlIn.Read()
                        If xmlIn.Name = "version" Then
                            ver = CSng(xmlIn.ReadElementString("version"))
                        End If
                        If xmlIn.EOF = True Then 'no items
                            xmlIn.Close()
                            Return oventdistributions
                            Exit Function
                        End If
                    Loop


                    Do While xmlIn.Name = "Vent"
                        Dim ovent As New oVent

                        xmlIn.ReadStartElement("Vent")

                        ovent.id = CInt(xmlIn.ReadElementString("id"))

                        Do While xmlIn.Name <> "vdistribution"
                            xmlIn.Read()
                            If xmlIn.EOF Then
                                MsgBox("distributions not found in vents.xml file")

                                xmlIn.Close()

                                Dim odistribution As New oDistribution

                                odistribution.id = ovent.id
                                odistribution.varname = "height"
                                odistribution.distribution = "None"
                                odistribution.varvalue = 1
                                oventdistributions.Add(odistribution)

                                odistribution.id = ovent.id
                                odistribution.varname = "width"
                                odistribution.distribution = "None"
                                odistribution.varvalue = 1
                                oventdistributions.Add(odistribution)

                                odistribution.id = ovent.id
                                odistribution.varname = "prob"
                                odistribution.distribution = "None"
                                odistribution.varvalue = 1
                                oventdistributions.Add(odistribution)

                                odistribution.id = ovent.id
                                odistribution.varname = "HOreliability"
                                odistribution.distribution = "None"
                                odistribution.varvalue = 1
                                oventdistributions.Add(odistribution)

                                odistribution.id = ovent.id
                                odistribution.varname = "integrity"
                                odistribution.distribution = "None"
                                odistribution.varvalue = 0
                                oventdistributions.Add(odistribution)

                                odistribution.id = ovent.id
                                odistribution.varname = "maxopening"
                                odistribution.distribution = "None"
                                odistribution.varvalue = 100
                                oventdistributions.Add(odistribution)

                                odistribution.id = ovent.id
                                odistribution.varname = "maxopeningtime"
                                odistribution.distribution = "None"
                                odistribution.varvalue = 10
                                oventdistributions.Add(odistribution)

                                odistribution.id = ovent.id
                                odistribution.varname = "gastemp"
                                odistribution.distribution = "None"
                                odistribution.varvalue = 200
                                oventdistributions.Add(odistribution)

                                Return oventdistributions
                                xmlIn.Close()
                                Exit Function

                            End If
                        Loop

                        Do While xmlIn.Name = "vdistribution"

                            Dim odistribution As New oDistribution

                            xmlIn.ReadStartElement("vdistribution")
                            odistribution.id = ovent.id 'saves item id

                            odistribution.varname = CStr(xmlIn.ReadElementString("varname"))
                            odistribution.varvalue = CStr(xmlIn.ReadElementString("value"))
                            odistribution.distribution = CStr(xmlIn.ReadElementString("distribution"))
                            If odistribution.distribution = "none" Then odistribution.distribution = "None"
                            odistribution.mean = xmlIn.ReadElementString("mean")
                            odistribution.variance = xmlIn.ReadElementString("variance")
                            odistribution.lbound = xmlIn.ReadElementString("lbound")
                            odistribution.ubound = xmlIn.ReadElementString("ubound")
                            odistribution.mode = xmlIn.ReadElementString("mode")
                            odistribution.alpha = xmlIn.ReadElementString("alpha")
                            odistribution.beta = xmlIn.ReadElementString("beta")
                            xmlIn.ReadEndElement()

                            oventdistributions.Add(odistribution)


                        Loop

                        If oventdistributions.Count Mod 3 = 0 And oventdistributions.Last.varname = "prob" Then
                            vflag = True
                            For Each oDistribution In oventdistributions
                                If oDistribution.id = ovent.id Then
                                    oventdistributions2.Add(oDistribution)

                                    If oDistribution.varname = "prob" Then
                                        Dim x As New oDistribution
                                        x.id = ovent.id 'saves item id

                                        x.varname = CStr("HOreliability")
                                        x.varvalue = 1
                                        x.distribution = CStr("None")
                                        x.mean = 0
                                        x.variance = 0
                                        x.lbound = 0
                                        x.ubound = 0
                                        x.mode = 0
                                        x.alpha = 0
                                        x.beta = 0

                                        oventdistributions2.Add(x)
                                    End If
                                End If
                            Next

                        End If

                        If oventdistributions.Count Mod 4 = 0 And oventdistributions.Last.varname = "HOreliability" Then
                            vflag = True
                            For Each oDistribution In oventdistributions
                                If oDistribution.id = ovent.id Then
                                    oventdistributions2.Add(oDistribution)

                                    If oDistribution.varname = "HOreliability" Then
                                        Dim x As New oDistribution
                                        x.id = ovent.id 'saves item id

                                        x.varname = CStr("integrity")
                                        x.varvalue = 0
                                        x.distribution = CStr("None")
                                        x.mean = 0
                                        x.variance = 0
                                        x.lbound = 0
                                        x.ubound = 0
                                        x.mode = 0
                                        x.alpha = 0
                                        x.beta = 0

                                        oventdistributions2.Add(x)

                                        Dim y As New oDistribution
                                        y.id = ovent.id 'saves item id

                                        y.varname = CStr("maxopening")
                                        y.varvalue = 100
                                        y.distribution = CStr("None")
                                        y.mean = 0
                                        y.variance = 0
                                        y.lbound = 0
                                        y.ubound = 0
                                        y.mode = 0
                                        y.alpha = 0
                                        y.beta = 0

                                        oventdistributions2.Add(y)

                                        Dim z As New oDistribution
                                        z.id = ovent.id 'saves item id

                                        z.varname = CStr("maxopeningtime")
                                        z.varvalue = 10
                                        z.distribution = CStr("None")
                                        z.mean = 0
                                        z.variance = 0
                                        z.lbound = 0
                                        z.ubound = 0
                                        z.mode = 0
                                        z.alpha = 0
                                        z.beta = 0

                                        oventdistributions2.Add(z)

                                    End If

                                End If
                            Next

                        End If

                        If oventdistributions.Count Mod 7 = 0 And oventdistributions.Last.varname = "maxopeningtime" Then
                            vflag = True
                            For Each oDistribution In oventdistributions
                                If oDistribution.id = ovent.id Then
                                    oventdistributions2.Add(oDistribution)

                                    If oDistribution.varname = "maxopeningtime" Then
                                        Dim x As New oDistribution
                                        x.id = ovent.id 'saves item id

                                        x.varname = CStr("gastemp")
                                        x.varvalue = 200
                                        x.distribution = CStr("None")
                                        x.mean = 0
                                        x.variance = 0
                                        x.lbound = 0
                                        x.ubound = 0
                                        x.mode = 0
                                        x.alpha = 0
                                        x.beta = 0

                                        oventdistributions2.Add(x)
                                    End If
                                End If
                            Next
                        End If

                        xmlIn.ReadEndElement()
                        ovents.Add(ovent)

                    Loop
                    If vflag = True Then oventdistributions = oventdistributions2
                End If
            End If

            Return oventdistributions

        Catch ex As Exception
            If TalkToEVACNZ = False Then MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in VentDB.vb GetVentDistributions")
            Return oventdistributions
        Finally
            xmlIn.Close()
        End Try
    End Function

    Public Sub New()

    End Sub

    Public Shared Function GetCVents() As List(Of oCVent)

        Dim ocvents As New List(Of oCVent)
        Dim ocventdistributions As New List(Of oDistribution)
        Dim Path As String = RiskDataDirectory & "cvents.xml"
        Dim ver As Single
        Dim xmlIn As New XmlTextReader(Path)

        Try

            If My.Computer.FileSystem.DirectoryExists(RiskDataDirectory) = True Then
                If My.Computer.FileSystem.FileExists(RiskDataDirectory & "cvents.xml") Then

                    ReDim Preserve NumberCVents(0 To MaxNumberRooms + 1, 0 To MaxNumberRooms + 1)

                    xmlIn.WhitespaceHandling = WhitespaceHandling.None

                    Do While xmlIn.Name <> "CVent"
                        xmlIn.Read()
                        If xmlIn.Name = "version" Then
                            ver = CSng(xmlIn.ReadElementString("version"))
                        End If
                        If xmlIn.EOF = True Then 'no vents
                            xmlIn.Close()
                            Return ocvents
                            Exit Function
                        End If
                    Loop



                    Do While xmlIn.Name = "CVent"
                        Dim ocvent As New oCVent

                        xmlIn.ReadStartElement("CVent")

                        ocvent.id = CInt(xmlIn.ReadElementString("id"))
                        ocvent.description = CStr(xmlIn.ReadElementString("description"))
                        ocvent.upperroom = CInt(xmlIn.ReadElementString("upperroom"))
                        ocvent.lowerroom = CInt(xmlIn.ReadElementString("lowerroom"))
                        'ocvent.area = CDbl(xmlIn.ReadElementString("area"))
                        ocvent.opentime = CDbl(xmlIn.ReadElementString("opentime"))
                        ocvent.closetime = CDbl(xmlIn.ReadElementString("closetime"))
                        ocvent.autoopenvent = CBool(xmlIn.ReadElementString("autoopen"))
                        'ocvent.integrity = CDbl(xmlIn.ReadElementString("integrity"))
                        'ocvent.maxopening = CDbl(xmlIn.ReadElementString("maxopening"))
                        'ocvent.maxopeningtime = CDbl(xmlIn.ReadElementString("maxopeningtime"))
                        'ocvent.gastemp = CDbl(xmlIn.ReadElementString("gastemp"))
                        ocvent.FRcriteria = CInt(xmlIn.ReadElementString("criteria"))
                        ocvent.triggerFR = CBool(xmlIn.ReadElementString("triggerFR"))

                        'added 25/11/14
                        If ver > CSng(2014.17) Then
                            ocvent.sdtriggerroom = CStr(xmlIn.ReadElementString("sdtriggerroom"))
                            ocvent.HRRthreshold = CStr(xmlIn.ReadElementString("HRRthreshold"))
                            ocvent.HRRventopendelay = CStr(xmlIn.ReadElementString("HRRventopendelay"))
                            ocvent.HRRventopenduration = CStr(xmlIn.ReadElementString("HRRventopenduration"))
                            ocvent.triggerventopendelay = CStr(xmlIn.ReadElementString("triggerventopendelay"))
                            ocvent.triggerventopenduration = CStr(xmlIn.ReadElementString("triggerventopenduration"))
                            ocvent.autoopenvent = CStr(xmlIn.ReadElementString("autoopenvent"))
                            ocvent.triggerHD = CStr(xmlIn.ReadElementString("triggerHD"))
                            ocvent.triggerHRR = CStr(xmlIn.ReadElementString("triggerHRR"))
                            ocvent.triggerSD = CStr(xmlIn.ReadElementString("triggerSD"))
                            ocvent.triggerFO = CStr(xmlIn.ReadElementString("triggerFO"))
                            ocvent.triggerVL = CStr(xmlIn.ReadElementString("triggerVL"))
                            ocvent.triggerHO = CStr(xmlIn.ReadElementString("triggerHO"))
                            ocvent.triggerFR = CStr(xmlIn.ReadElementString("triggerFR"))
                        End If

                        'added 1/7/2016
                        If ver > CSng(2016.01) Then
                            ocvent.dischargecoeff = CStr(xmlIn.ReadElementString("dischargecoeff"))
                        End If

                        Dim upperroom As Integer = ocvent.upperroom
                        Dim lowerroom As Integer = ocvent.lowerroom
                        Dim bx As Integer = 0

                        Do While xmlIn.Name = "vdistribution"
                            Dim odistribution As New oDistribution
                            bx = bx + 1
                            xmlIn.ReadStartElement("vdistribution")
                            odistribution.id = ocvent.id 'saves item id
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

                            If odistribution.varname = "area" Then ocvent.area = odistribution.varvalue
                            If odistribution.varname = "integrity" Then ocvent.integrity = odistribution.varvalue
                            If odistribution.varname = "maxopening" Then ocvent.maxopening = odistribution.varvalue
                            If odistribution.varname = "maxopeningtime" Then ocvent.maxopeningtime = odistribution.varvalue
                            If odistribution.varname = "gastemp" Then ocvent.gastemp = odistribution.varvalue

                            ocventdistributions.Add(odistribution)
                        Loop

                        ocvents.Add(ocvent)
                        xmlIn.ReadEndElement()
                        'NumberCVents(upperroom, lowerroom) = NumberCVents(upperroom, lowerroom) + 1
                        'ReDim Preserve CVentArea(MaxNumberRooms + 1, MaxNumberRooms + 1, NumberCVents(upperroom, lowerroom))
                        'CVentArea(upperroom, lowerroom, NumberCVents(upperroom, lowerroom)) = ocvent.area
                    Loop
                End If
            End If

            xmlIn.Close()

            VentDB.SaveCVents(ocvents, ocventdistributions)

            Return ocvents

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in VentDB.vb GetCVents")
            xmlIn.Close()
            Return ocvents
        End Try
    End Function
    Public Shared Sub SaveCVents(ByVal ocvents As List(Of oCVent), ByVal ocventdistributions As List(Of oDistribution))
        'these are distinct items listed and stored in the cvents.xml file.
        If IsNothing(ocvents) Then Exit Sub

        Dim Path As String = RiskDataDirectory & "cvents.xml"
        If Not My.Computer.FileSystem.DirectoryExists(Path) Then
            My.Computer.FileSystem.CreateDirectory(RiskDataDirectory)
        End If

        Dim xmlOut As New XmlTextWriter(Path, System.Text.Encoding.UTF8)
        Try
            xmlOut.Formatting = Formatting.Indented
            xmlOut.WriteStartDocument()

            xmlOut.WriteStartElement("CVents")
            xmlOut.WriteElementString("version", Version.ToString)

            Dim ocVent As oCVent

            For i As Integer = 0 To ocvents.Count - 1
                ocVent = CType(ocvents(i), oCVent)
                xmlOut.WriteStartElement("CVent")
                xmlOut.WriteElementString("id", CStr(i + 1))
                xmlOut.WriteElementString("description", ocVent.description)
                xmlOut.WriteElementString("upperroom", ocVent.upperroom.ToString)
                xmlOut.WriteElementString("lowerroom", ocVent.lowerroom.ToString)
                'xmlOut.WriteElementString("area", ocVent.area.ToString)
                xmlOut.WriteElementString("opentime", ocVent.opentime.ToString)
                xmlOut.WriteElementString("closetime", ocVent.closetime.ToString)
                xmlOut.WriteElementString("autoopen", ocVent.autoopenvent.ToString)
                'xmlOut.WriteElementString("integrity", ocVent.integrity.ToString)
                'xmlOut.WriteElementString("maxopening", ocVent.maxopening.ToString)
                'xmlOut.WriteElementString("maxopeningtime", ocVent.maxopeningtime.ToString)
                'xmlOut.WriteElementString("gastemp", ocVent.gastemp.ToString)
                xmlOut.WriteElementString("criteria", ocVent.FRcriteria.ToString)
                xmlOut.WriteElementString("triggerFR", ocVent.triggerFR.ToString)

                'added 25/11/14
                xmlOut.WriteElementString("sdtriggerroom", ocVent.sdtriggerroom.ToString)
                xmlOut.WriteElementString("HRRthreshold", ocVent.HRRthreshold.ToString)
                xmlOut.WriteElementString("HRRventopendelay", ocVent.HRRventopendelay.ToString)
                xmlOut.WriteElementString("HRRventopenduration", ocVent.HRRventopenduration.ToString)
                xmlOut.WriteElementString("triggerventopendelay", ocVent.triggerventopendelay.ToString)
                xmlOut.WriteElementString("triggerventopenduration", ocVent.triggerventopenduration.ToString)
                xmlOut.WriteElementString("autoopenvent", ocVent.autoopenvent.ToString)
                xmlOut.WriteElementString("triggerHD", ocVent.triggerHD.ToString)
                xmlOut.WriteElementString("triggerHRR", ocVent.triggerHRR.ToString)
                xmlOut.WriteElementString("triggerSD", ocVent.triggerSD.ToString)
                xmlOut.WriteElementString("triggerFO", ocVent.triggerFO.ToString)
                xmlOut.WriteElementString("triggerVL", ocVent.triggerVL.ToString)
                xmlOut.WriteElementString("triggerHO", ocVent.triggerHO.ToString)
                xmlOut.WriteElementString("triggerFR", ocVent.triggerFR.ToString)

                'added 1/7/2016
                xmlOut.WriteElementString("dischargecoeff", ocVent.dischargecoeff.ToString)

                Dim bx As Integer = 0
                For Each odistribution In ocventdistributions
                    bx = bx + 1

                    If odistribution.id = i + 1 Then
                        xmlOut.WriteStartElement("vdistribution")
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

            NumCVents = ocvents.Count

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in VentDB.vb SaveCVents")
            xmlOut.WriteEndDocument()
            xmlOut.Close()
        End Try

    End Sub
    Public Shared Function GetCVentDistributions() As List(Of oDistribution)
        Dim ocvents As New List(Of oCVent)
        Dim ocventdistributions As New List(Of oDistribution)
        Dim ocventdistributions2 As New List(Of oDistribution)
        Dim Path As String = RiskDataDirectory & "cvents.xml"
        Dim ver As Single
        Dim xmlIn As New XmlTextReader(Path)
        Dim cvflag As Boolean = False

        Try


            If My.Computer.FileSystem.DirectoryExists(RiskDataDirectory) = True Then
                If My.Computer.FileSystem.FileExists(RiskDataDirectory & "cvents.xml") Then

                    xmlIn.WhitespaceHandling = WhitespaceHandling.None

                    Do While xmlIn.Name <> "CVent"
                        xmlIn.Read()
                        If xmlIn.Name = "version" Then
                            ver = CSng(xmlIn.ReadElementString("version"))
                        End If
                        If xmlIn.EOF = True Then 'no items
                            xmlIn.Close()
                            Return ocventdistributions
                            Exit Function
                        End If
                    Loop


                    Do While xmlIn.Name = "CVent"
                        Dim ocvent As New oCVent

                        xmlIn.ReadStartElement("CVent")

                        ocvent.id = CInt(xmlIn.ReadElementString("id"))

                        Do While xmlIn.Name <> "vdistribution"
                            xmlIn.Read()
                            If xmlIn.EOF Then
                                MsgBox("distributions not found in cvents.xml file")

                                xmlIn.Close()

                                Dim odistribution As New oDistribution

                                odistribution.id = ocvent.id
                                odistribution.varname = "area"
                                odistribution.distribution = "None"
                                odistribution.varvalue = 1
                                ocventdistributions.Add(odistribution)

                                odistribution.id = ocvent.id
                                odistribution.varname = "integrity"
                                odistribution.distribution = "None"
                                odistribution.varvalue = 0
                                ocventdistributions.Add(odistribution)

                                odistribution.id = ocvent.id
                                odistribution.varname = "maxopening"
                                odistribution.distribution = "None"
                                odistribution.varvalue = 100
                                ocventdistributions.Add(odistribution)

                                odistribution.id = ocvent.id
                                odistribution.varname = "maxopeningtime"
                                odistribution.distribution = "None"
                                odistribution.varvalue = 10
                                ocventdistributions.Add(odistribution)

                                odistribution.id = ocvent.id
                                odistribution.varname = "gastemp"
                                odistribution.distribution = "None"
                                odistribution.varvalue = 200
                                ocventdistributions.Add(odistribution)

                                Return ocventdistributions
                                xmlIn.Close()
                                Exit Function

                            End If
                        Loop

                        Do While xmlIn.Name = "vdistribution"

                            Dim odistribution As New oDistribution

                            xmlIn.ReadStartElement("vdistribution")
                            odistribution.id = ocvent.id 'saves item id

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

                            ocventdistributions.Add(odistribution)


                        Loop

                        xmlIn.ReadEndElement()
                        ocvents.Add(ocvent)

                    Loop
                    If cvflag = True Then ocventdistributions = ocventdistributions2
                End If
            End If

            xmlIn.Close()

            Return ocventdistributions

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in VentDB.vb GetCVentDistributions")
            xmlIn.Close()
            Return ocventdistributions
        End Try
    End Function
End Class
