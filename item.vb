Imports System.Collections.Generic
Imports System.Math
Imports CenterSpace.NMath.Core
Imports CenterSpace.NMath.Stats

Public Class oItem
    Public tempdatain() As Double

    Private m_description As String
    Private m_type As String
    Private m_detaileddescription As String
    Private m_userlabel As String
    Private m_constantA As Single
    Private m_constantB As Single
    Private m_LHoG As Single
    Private m_co2 As Single
    Private m_HRRUA As Single
    Private m_length As Single
    Private m_width As Single
    Private m_height As Single
    Private m_elevation As Single
    Private m_hrr As String
    Private m_mlrfreeburn
    Private m_prob As Single
    Private m_mass As Single
    Private m_critflux As Single
    Private m_critfluxauto As Single
    Private m_ftplimitpilot As Single
    Private m_ftplimitauto As Single
    Private m_ftpindexpilot As Single
    Private m_ftpindexauto As Single
    Private m_soot As Single
    Private m_hoc As Single
    Private m_id As Integer
    Private m_ignitiontime As Single
    Private m_maxnumitem As Integer
    Private m_xleft As Single
    Private m_ybottom As Single
    Private m_radlossfrac As Single
    Private m_windeffect As Single

    Private m_pooldiameter As Double
    Private m_poolvolume As Double
    Private m_pooldensity As Double
    Private m_poolramp As Double
    Private m_poolFBMLR As Double
    Private m_poolvaptemp As Double
    Private m_pyrolysisoption As Integer


    Public Sub New()
    End Sub
    Public Sub New(ByVal mlrfreeburn As String, ByVal poolvaptemp As Double, ByVal pooldiameter As Double, ByVal poolvolume As Double, ByVal pooldensity As Double, ByVal poolramp As Double, ByVal poolFBMLR As Double,
                   ByVal windeffect As Single, ByVal hrrua As Single, ByVal detaileddescription As String, ByVal userlabel As String, ByVal critfluxmode As Single, ByVal critfluxalpha As Single, ByVal critfluxbeta As Single, ByVal ftpindexpilotmode As Single, ByVal ftpindexpilotalpha As Single, ByVal ftpindexpilotbeta As Single, ByVal ftplimitpilotmode As Single, ByVal ftplimitpilotalpha As Single, ByVal ftplimitpilotbeta As Single, ByVal sootmode As Single, ByVal sootalpha As Single, ByVal sootbeta As Single, ByVal hocmode As Single, ByVal hocalpha As Single, ByVal hocbeta As Single, ByVal ftpindexauto As Single, ByVal ftpindexpilot As Single, ByVal ftplimitauto As Single, ByVal ftplimitpilot As Single, ByVal igntime As Single, ByVal hrr As String, ByVal mass As Single, ByVal prob As Single, ByVal critflux As Single, ByVal critfluxauto As Single,
                   ByVal height As Single, ByVal elevation As Single, ByVal width As Single, ByVal length As Single, ByVal co As Single,
                   ByVal co2 As Single, ByVal soot As Single, ByVal type As String, ByVal description As String, ByVal hoc As Single,
                   ByVal hocdistribution As String, ByVal hocmean As Single, ByVal hocvariance As Single, ByVal hoclbound As Single,
                   ByVal hocubound As Single, ByVal sootdistribution As String, ByVal sootmean As Single, ByVal sootvariance As Single,
                   ByVal sootlbound As Single, ByVal sootubound As Single, ByVal ftplimitpilotdistribution As String, ByVal ftplimitpilotmean As Single, ByVal ftplimitpilotvariance As Single,
                   ByVal ftplimitpilotlbound As Single, ByVal ftplimitpilotubound As Single, ByVal ftindexpilotdistribution As String, ByVal ftpindexpilotmean As Single, ByVal ftpindexpilotvariance As Single,
                   ByVal ftpindexpilotlbound As Single, ByVal ftpindexpilotubound As Single, ByVal critfluxdistribution As String, ByVal critfluxmean As Single, ByVal critfluxvariance As Single,
                   ByVal critfluxlbound As Single, ByVal critfluxubound As Single, ByVal maxnumitem As Integer, ByVal xleft As Single, ByVal ybottom As Single, ByVal radlossfrac As Single, ByVal constantA As Single, ByVal constantB As Single, ByVal LHoG As Single, ByVal pyrolysisoption As Integer)

        Me.mlrfreeburn = mlrfreeburn
        Me.poolvaptemp = poolvaptemp
        Me.pooldiameter = pooldiameter
        Me.poolvolume = poolvolume
        Me.pooldensity = pooldensity
        Me.poolramp = poolramp
        Me.poolFBMLR = poolFBMLR
        Me.description = description
        Me.detaileddescription = detaileddescription
        Me.userlabel = userlabel
        Me.type = type
        Me.ignitiontime = igntime
        Me.length = length
        Me.width = width
        Me.elevation = elevation
        Me.height = height
        Me.prob = prob
        Me.mass = mass
        Me.hrr = hrr
        Me.maxnumitem = maxnumitem
        Me.xleft = xleft
        Me.ybottom = ybottom
        Me.radlossfrac = radlossfrac
        Me.constantA = constantA
        Me.constantB = constantB
        Me.LHoG = LHoG
        Me.HRRUA = hrrua
        Me.hoc = hoc
        Me.critflux = critflux
        Me.critfluxauto = critfluxauto
        Me.ftplimitpilot = ftplimitpilot
        Me.ftplimitauto = ftplimitauto
        Me.ftpindexpilot = ftpindexpilot
        Me.ftpindexauto = ftpindexauto
        Me.co2 = co2
        Me.soot = soot
        Me.windeffect = windeffect
        Me.pyrolysisoption = pyrolysisoption

    End Sub
    Public Property pyrolysisoption() As Integer
        Get
            Return m_pyrolysisoption
        End Get
        Set(ByVal value As Integer)
            m_pyrolysisoption = value
        End Set
    End Property
    Public Property mlrfreeburn() As String
        Get
            Return m_mlrfreeburn
        End Get
        Set(ByVal value As String)
            m_mlrfreeburn = value
        End Set
    End Property
    Public Property poolvaptemp() As Double
        Get
            Return m_poolvaptemp
        End Get
        Set(ByVal value As Double)
            m_poolvaptemp = value
        End Set
    End Property
    Public Property poolFBMLR() As Double
        Get
            Return m_poolFBMLR
        End Get
        Set(ByVal value As Double)
            m_poolFBMLR = value
        End Set
    End Property
    Public Property poolramp() As Double
        Get
            Return m_poolramp
        End Get
        Set(ByVal value As Double)
            m_poolramp = value
        End Set
    End Property
    Public Property pooldensity() As Double
        Get
            Return m_pooldensity
        End Get
        Set(ByVal value As Double)
            m_pooldensity = value
        End Set
    End Property
    Public Property poolvolume() As Double
        Get
            Return m_poolvolume
        End Get
        Set(ByVal value As Double)
            m_poolvolume = value
        End Set
    End Property
    Public Property pooldiameter() As Double
        Get
            Return m_pooldiameter
        End Get
        Set(ByVal value As Double)
            m_pooldiameter = value
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
    Public Property detaileddescription() As String
        Get
            Return m_detaileddescription
        End Get
        Set(ByVal value As String)
            m_detaileddescription = value
        End Set
    End Property
    Public Property userlabel() As String
        Get
            Return m_userlabel
        End Get
        Set(ByVal value As String)
            m_userlabel = value
        End Set
    End Property
    Public Property windeffect() As Single
        Get
            Return m_windeffect
        End Get
        Set(ByVal value As Single)
            m_windeffect = value
        End Set
    End Property
    Public Property LHoG() As Single
        Get
            Return m_LHoG
        End Get
        Set(ByVal value As Single)
            m_LHoG = value
        End Set
    End Property
    Public Property HRRUA() As Single
        Get
            Return m_HRRUA
        End Get
        Set(ByVal value As Single)
            m_HRRUA = value
        End Set
    End Property
    Public Property constantA() As Single
        Get
            Return m_constantA
        End Get
        Set(ByVal value As Single)
            m_constantA = value
        End Set
    End Property
    Public Property constantB() As Single
        Get
            Return m_constantB
        End Get
        Set(ByVal value As Single)
            m_constantB = value
        End Set
    End Property
    Public Property radlossfrac() As Single
        Get
            Return m_radlossfrac
        End Get
        Set(ByVal value As Single)
            m_radlossfrac = value
        End Set
    End Property
    Public Property maxnumitem() As Integer
        Get
            Return m_maxnumitem
        End Get
        Set(ByVal value As Integer)
            m_maxnumitem = value
        End Set
    End Property
    Public Property xleft() As Single
        Get
            Return m_xleft
        End Get
        Set(ByVal value As Single)
            m_xleft = value
        End Set
    End Property
    Public Property ybottom() As Single
        Get
            Return m_ybottom
        End Get
        Set(ByVal value As Single)
            m_ybottom = value
        End Set
    End Property
    Public Property description() As String
        Get
            Return m_description
        End Get
        Set(ByVal value As String)
            m_description = value
        End Set
    End Property
    Public Property type() As String
        Get
            Return m_type
        End Get
        Set(ByVal value As String)
            m_type = value
        End Set
    End Property
    Public Property ftplimitauto() As Single
        Get
            Return m_ftplimitauto
        End Get
        Set(ByVal value As Single)
            m_ftplimitauto = value
        End Set
    End Property
    Public Property ftplimitpilot() As Single
        Get
            Return m_ftplimitpilot
        End Get
        Set(ByVal value As Single)
            m_ftplimitpilot = value
        End Set
    End Property
    Public Property ftpindexpilot() As Single
        Get
            Return m_ftpindexpilot
        End Get
        Set(ByVal value As Single)
            m_ftpindexpilot = value
        End Set
    End Property
    Public Property ftpindexauto() As Single
        Get
            Return m_ftpindexauto
        End Get
        Set(ByVal value As Single)
            m_ftpindexauto = value
        End Set
    End Property
    Public Property ignitiontime() As Single
        Get
            Return m_ignitiontime
        End Get
        Set(ByVal value As Single)
            m_ignitiontime = value
        End Set
    End Property
    Public Property mass() As Single
        Get
            Return m_mass
        End Get
        Set(ByVal value As Single)
            m_mass = value
        End Set
    End Property
    Public Property co2() As Single
        Get
            Return m_co2
        End Get
        Set(ByVal value As Single)
            m_co2 = value
        End Set
    End Property

    Public Property soot() As Single
        Get
            Return m_soot
        End Get
        Set(ByVal value As Single)
            m_soot = value
        End Set
    End Property
    Public Property critflux() As Single
        Get
            Return m_critflux
        End Get
        Set(ByVal value As Single)
            m_critflux = value
        End Set
    End Property
    Public Property critfluxauto() As Single
        Get
            Return m_critfluxauto
        End Get
        Set(ByVal value As Single)
            m_critfluxauto = value
        End Set
    End Property
    Public Property hoc() As Single
        Get
            Return m_hoc
        End Get
        Set(ByVal value As Single)
            m_hoc = value
        End Set
    End Property
    Public Property length() As Single
        Get
            Return m_length
        End Get
        Set(ByVal value As Single)
            m_length = value
        End Set
    End Property
    Public Property width() As Single
        Get
            Return m_width
        End Get
        Set(ByVal value As Single)
            m_width = value
        End Set
    End Property
    Public Property elevation() As Single
        Get
            Return m_elevation
        End Get
        Set(ByVal value As Single)
            m_elevation = value
        End Set
    End Property
    Public Property height() As Single
        Get
            Return m_height
        End Get
        Set(ByVal value As Single)
            m_height = value
        End Set
    End Property
    Public Property prob() As Single
        Get
            Return m_prob
        End Get
        Set(ByVal value As Single)
            m_prob = value
        End Set
    End Property
    Public Property hrr() As String
        Get
            Return m_hrr
        End Get
        Set(ByVal value As String)
            m_hrr = value
        End Set
    End Property

    Public Function GetDisplayText(ByVal sep As String) As String
        Dim text As String
        'text = room & sep & rti & sep & cfactor & sep & acttemp & sep & sprx & sep & spry & sep & sprz & sep & sprr & sep & sprdensity
        text = description

        Return text
    End Function
End Class
