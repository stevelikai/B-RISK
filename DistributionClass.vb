Imports System.Xml
Imports System.Collections.Generic

Public Class oDistribution
    Private m_varname As String
    Private m_distribution As String
    Private m_varvalue As Double
    Private m_mean As Double
    Private m_variance As Double
    Private m_lbound As Double
    Private m_ubound As Double
    Private m_alpha As Double
    Private m_beta As Double
    Private m_mode As Double
    Private m_id As Integer
    Private m_units As String

    Public Sub New()
    End Sub

    Public Sub New(ByVal varname As String, ByVal units As String, ByVal distribution As String, ByVal varvalue As Double, ByVal mean As Double, ByVal variance As Double, ByVal lbound As Double, ByVal ubound As Double, ByVal mode As Double, ByVal alpha As Double, ByVal beta As Double)

        Me.varname = varname
        Me.units = units
        Me.distribution = distribution
        Me.varvalue = varvalue
        Me.mean = mean
        Me.variance = variance
        Me.lbound = lbound
        Me.ubound = ubound
        Me.mode = mode
        Me.alpha = alpha
        Me.beta = beta

    End Sub
    Public Property units() As String
        Get
            Return m_units
        End Get
        Set(ByVal value As String)
            m_units = value
        End Set
    End Property
    Public Property beta() As Double
        Get
            Return m_beta
        End Get
        Set(ByVal value As Double)
            m_beta = value
        End Set
    End Property
    Public Property alpha() As Double
        Get
            Return m_alpha
        End Get
        Set(ByVal value As Double)
            m_alpha = value
        End Set
    End Property
    Public Property mode() As Double
        Get
            Return m_mode
        End Get
        Set(ByVal value As Double)
            m_mode = value
        End Set
    End Property
    Public Property varname() As String
        Get
            Return m_varname
        End Get
        Set(ByVal value As String)
            m_varname = value
        End Set
    End Property
    Public Property distribution() As String
        Get
            Return m_distribution
        End Get
        Set(ByVal value As String)
            m_distribution = value
        End Set
    End Property
    Public Property varvalue() As Double
        Get
            Return m_varvalue
        End Get
        Set(ByVal value As Double)
            m_varvalue = value
        End Set
    End Property
    Public Property mean() As Double
        Get
            Return m_mean
        End Get
        Set(ByVal value As Double)
            m_mean = value
        End Set
    End Property
    Public Property variance() As Double
        Get
            Return m_variance
        End Get
        Set(ByVal value As Double)
            m_variance = value
        End Set
    End Property
    Public Property lbound() As Double
        Get
            Return m_lbound
        End Get
        Set(ByVal value As Double)
            m_lbound = value
        End Set
    End Property
    Public Property ubound() As Double
        Get
            Return m_ubound
        End Get
        Set(ByVal value As Double)
            m_ubound = value
        End Set
    End Property
    Public Property id() As Integer
        Get
            Return m_id
        End Get
        Set(ByVal value As Integer)
            m_id = value
        End Set
    End Property
End Class

Public Class DistributionClass
    Public Shared Function GetDistributions() As List(Of oDistribution)

        'get distribution details from distributions.xml file
        Dim oDistributions As New List(Of oDistribution)
        If RiskDataDirectory Is Nothing Then
            Call New_File()
            ProjectDirectory = RiskDataDirectory
        End If
        ProjectDirectory = RiskDataDirectory

        Dim Path As String = ProjectDirectory & "distributions.xml"
        Dim xmlIn As New XmlTextReader(Path)

        Try

            If My.Computer.FileSystem.FileExists(Path) = False Then
                MsgBox("The file distributions.xml is missing", MsgBoxStyle.Critical, "Missing File")
                xmlIn.Close()
                Return oDistributions

                Exit Function
            End If

            If My.Computer.FileSystem.FileExists(Path) Then


                xmlIn.WhitespaceHandling = WhitespaceHandling.None

                Do While xmlIn.Name <> "Distribution"
                    xmlIn.Read()
                    If xmlIn.EOF = True Then 'no distributions
                        xmlIn.Close()
                        Return oDistributions
                        Exit Function
                    End If
                Loop

                Do While xmlIn.Name = "Distribution"
                    Dim oDistribution As New oDistribution
                    xmlIn.ReadStartElement("Distribution")

                    oDistribution.id = CSng(xmlIn.ReadElementString("id"))
                    oDistribution.varname = CStr(xmlIn.ReadElementString("varname"))
                    oDistribution.units = CStr(xmlIn.ReadElementString("units"))
                    oDistribution.distribution = CStr(xmlIn.ReadElementString("distribution"))
                    If oDistribution.distribution = "none" Then oDistribution.distribution = "None"
                    oDistribution.varvalue = CDbl(xmlIn.ReadElementString("varvalue"))
                    oDistribution.mean = CDbl(xmlIn.ReadElementString("mean"))
                    oDistribution.variance = CDbl(xmlIn.ReadElementString("variance"))
                    oDistribution.lbound = CDbl(xmlIn.ReadElementString("lbound"))
                    oDistribution.ubound = CDbl(xmlIn.ReadElementString("ubound"))
                    oDistribution.mode = CDbl(xmlIn.ReadElementString("mode"))
                    oDistribution.alpha = CDbl(xmlIn.ReadElementString("alpha"))
                    oDistribution.beta = CDbl(xmlIn.ReadElementString("beta"))
                    xmlIn.ReadEndElement()

                    oDistributions.Add(oDistribution)

                Loop

                'if reading older file, add any missing distributions
                If oDistributions.Count = 5 Then
                    'add Soot Preflashover Yield
                    Dim oDistribution As New oDistribution
                    oDistribution.id = 6
                    oDistribution.varname = "Soot Preflashover Yield"
                    oDistribution.distribution = "None"
                    oDistribution.units = "g/g"
                    oDistribution.varvalue = preSoot
                    oDistributions.Add(oDistribution)
                End If
                If oDistributions.Count = 6 Then
                    'add CO Preflashover Yield
                    Dim oDistribution As New oDistribution
                    oDistribution.id = 7
                    oDistribution.varname = "CO Preflashover Yield"
                    oDistribution.distribution = "None"
                    oDistribution.units = "g/g"
                    oDistribution.varvalue = preCO
                    oDistributions.Add(oDistribution)
                End If
                If oDistributions.Count = 7 Then
                    'add sprinkler reliability
                    Dim oDistribution As New oDistribution
                    oDistribution.id = 8
                    oDistribution.varname = "Sprinkler Reliability"
                    oDistribution.distribution = "None"
                    oDistribution.units = "-"
                    oDistribution.varvalue = SprReliability
                    oDistributions.Add(oDistribution)
                End If
                If oDistributions.Count = 8 Then
                    'add sprinkler supresssion
                    Dim oDistribution As New oDistribution
                    oDistribution.id = 9
                    oDistribution.varname = "Sprinkler Suppression Probability"
                    oDistribution.distribution = "None"
                    oDistribution.units = "-"
                    oDistribution.varvalue = SprSuppressionProb
                    oDistributions.Add(oDistribution)
                End If
                If oDistributions.Count = 9 Then
                    'add sprinkler cooling coefficient
                    Dim oDistribution As New oDistribution
                    oDistribution.id = 10
                    oDistribution.varname = "Sprinkler Cooling Coefficient"
                    oDistribution.distribution = "None"
                    oDistribution.units = "-"
                    oDistribution.varvalue = SprCooling
                    oDistributions.Add(oDistribution)
                End If
                If oDistributions.Count = 10 Then
                    'add fuel heat of gasification
                    Dim oDistribution As New oDistribution
                    oDistribution.id = 11
                    oDistribution.varname = "Fuel Heat of Gasification"
                    oDistribution.distribution = "None"
                    oDistribution.units = "kJ/g"
                    'oDistribution.varvalue = FuelHeatofGasification
                    oDistributions.Add(oDistribution)
                End If
                If oDistributions.Count = 11 Then
                    'add alpha T
                    Dim oDistribution As New oDistribution
                    oDistribution.id = 12
                    oDistribution.varname = "Alpha T"
                    oDistribution.distribution = "None"
                    oDistribution.units = "kW/s2"
                    oDistribution.varvalue = AlphaT
                    oDistributions.Add(oDistribution)
                End If
                If oDistributions.Count = 12 Then
                    'add peak hrr
                    Dim oDistribution As New oDistribution
                    oDistribution.id = 13
                    oDistribution.varname = "Peak HRR"
                    oDistribution.distribution = "None"
                    oDistribution.units = "kW"
                    oDistribution.varvalue = PeakHRR
                    oDistributions.Add(oDistribution)
                End If
                If oDistributions.Count = 13 Then
                    'add sd reliability
                    Dim oDistribution As New oDistribution
                    oDistribution.id = 14
                    oDistribution.varname = "Smoke Detector Reliability"
                    oDistribution.distribution = "None"
                    oDistribution.units = "-"
                    oDistribution.varvalue = SDReliability
                    oDistributions.Add(oDistribution)
                End If
                If oDistributions.Count = 14 Then
                    'add mech vent reliability
                    Dim oDistribution As New oDistribution
                    oDistribution.id = 15
                    oDistribution.varname = "Mechanical Ventilation Reliability"
                    oDistribution.distribution = "None"
                    oDistribution.units = "-"
                    oDistribution.varvalue = FanReliability
                    oDistributions.Add(oDistribution)
                End If
            Else

                MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Distribution.xml file is missing")

            End If

            xmlIn.Close()
            Return oDistributions
            Exit Function

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in DistributionClass.GetDistributions")
            xmlIn.Close()
            Return oDistributions
        End Try
    End Function

    Public Shared Sub SaveDistributions(ByVal oDistributions As List(Of oDistribution))

        Dim rname As String = RiskDataDirectory & "distributions.xml"

        Dim Path As String = rname
        Dim xmlOut As New XmlTextWriter(Path, System.Text.Encoding.UTF8)

        Try

            xmlOut.Formatting = Formatting.Indented

            xmlOut.WriteStartDocument()
            xmlOut.WriteStartElement("Distributions")

            Dim oDistribution As oDistribution
            For i As Integer = 0 To oDistributions.Count - 1
                oDistribution = CType(oDistributions(i), oDistribution)
                xmlOut.WriteStartElement("Distribution")

                xmlOut.WriteElementString("id", CStr(i + 1))
                xmlOut.WriteElementString("varname", oDistribution.varname.ToString)
                xmlOut.WriteElementString("units", oDistribution.units.ToString)
                xmlOut.WriteElementString("distribution", oDistribution.distribution.ToString)
                xmlOut.WriteElementString("varvalue", oDistribution.varvalue.ToString)
                xmlOut.WriteElementString("mean", oDistribution.mean.ToString)
                xmlOut.WriteElementString("variance", oDistribution.variance.ToString)
                xmlOut.WriteElementString("lbound", oDistribution.lbound.ToString)
                xmlOut.WriteElementString("ubound", oDistribution.ubound.ToString)
                xmlOut.WriteElementString("mode", oDistribution.mode.ToString)
                xmlOut.WriteElementString("alpha", oDistribution.alpha.ToString)
                xmlOut.WriteElementString("beta", oDistribution.beta.ToString)

                xmlOut.WriteEndElement()
            Next

            xmlOut.WriteEndElement()
            xmlOut.WriteEndDocument()
            xmlOut.Close()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in DistributionClass.vb SaveDistributions")
            xmlOut.WriteEndDocument()
            xmlOut.Close()
        End Try

    End Sub
    Public Shared Sub Read_Distributions()
        'read data from distribution file to update text boxes when loading a new basefile

        Dim odistributions As List(Of oDistribution)
        odistributions = DistributionClass.GetDistributions()

        For Each oDistribution In odistributions
            If oDistribution.varname = "Interior Temperature" Then InteriorTemp = oDistribution.varvalue
            If oDistribution.varname = "Exterior Temperature" Then ExteriorTemp = oDistribution.varvalue
            If oDistribution.varname = "Relative Humidity" Then RelativeHumidity = oDistribution.varvalue
            If oDistribution.varname = "Heat of Combustion PFO" Then HoC_fuel = oDistribution.varvalue
            If oDistribution.varname = "Fire Load Energy Density" Then FLED = oDistribution.varvalue
            If oDistribution.varname = "Soot Preflashover Yield" Then preSoot = oDistribution.varvalue
            If oDistribution.varname = "CO Preflashover Yield" Then preCO = oDistribution.varvalue
            If oDistribution.varname = "Sprinkler Reliability" Then SprReliability = CDec(oDistribution.varvalue)
            If oDistribution.varname = "Sprinkler Suppression Probability" Then SprSuppressionProb = CDec(oDistribution.varvalue)
            If oDistribution.varname = "Sprinkler Cooling Coefficient" Then SprCooling = CDec(oDistribution.varvalue)
            If oDistribution.varname = "Alpha T" Then AlphaT = CSng(oDistribution.varvalue)
            If oDistribution.varname = "Peak HRR" Then PeakHRR = CSng(oDistribution.varvalue)
            If oDistribution.varname = "Smoke Detector Reliability" Then SDReliability = CDec(oDistribution.varvalue)
            If oDistribution.varname = "Mechanical Ventilation Reliability" Then FanReliability = CDec(oDistribution.varvalue)

        Next

        frmOptions1.txtInteriorTemp.Text = InteriorTemp - 273
        frmOptions1.txtInteriorTemp.Text = ExteriorTemp - 273
        frmOptions1.txtRelativeHumidity.Text = RelativeHumidity * 100
        frmOptions1.txtHOCFuel.Text = HoC_fuel
        frmItemList.txtFLED.Text = FLED
        frmOptions1.txtpreSoot.Text = preSoot
        frmOptions1.txtpreCO.Text = preCO
        frmSprinklerList.txtSprReliability.Text = CStr(SprReliability)
        frmSprinklerList.txtSprSuppressProb.Text = CStr(SprSuppressionProb)
        frmSprinklerList.txtSprCoolingCoeff.Text = CStr(SprCooling)
        frmPowerlaw.TxtAlphaT.Text = Format(AlphaT, "0.000")
        frmPowerlaw.txtPeakHRR.Text = CStr(PeakHRR)
        frmSmokeDetList.txtSmokeDetReliability.Text = CStr(SDReliability)
        frmFanList.txtFanReliability.Text = CStr(FanReliability)
    End Sub
    
    Public Sub New()

    End Sub
End Class
