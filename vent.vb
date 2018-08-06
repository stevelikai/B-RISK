Imports System.Collections.Generic
Imports System.Math
Imports CenterSpace.NMath.Core
Imports CenterSpace.NMath.Stats

Public Class oVent
    Implements ICloneable

    Private m_id As Integer
    Private m_fromroom As Integer
    Private m_toroom As Integer
    Private m_face As Integer
    Private m_offset As Double
    Private m_height As Double
    Private m_width As Double
    Private m_sillheight As Double
    Private m_opentime As Double
    Private m_closetime As Double
    Private m_walllength1 As Double
    Private m_walllength2 As Double
    Private m_autobreakglass As Boolean
    Private m_description As String
    Private m_glassconductivity As Single
    Private m_glassemissivity As Single
    Private m_glassexpansion As Single
    Private m_glassthickness As Single
    Private m_glassshading As Single
    Private m_glassbreakingstress As Single
    Private m_glassalpha As Single
    Private m_glassYoungsModulus As Single
    Private m_glassflameflux As Boolean
    Private m_glassdistance As Single
    Private m_glassfalloutime As Single
    Private m_spillplume As Boolean
    Private m_downstand As Single
    Private m_spillplumebalc As Boolean
    Private m_spillplumemodel As Short
    Private m_spillbalconyprojection As Single
    Private m_sdtriggerroom As Integer
    Private m_HRRthreshold As Double
    Private m_HRRventopendelay As Double
    Private m_HRRventopenduration As Double
    Private m_triggerventopendelay As Double
    Private m_triggerventopenduration As Double
    Private m_autoopenvent As Boolean
    Private m_triggerSD As Boolean
    Private m_triggerHD As Boolean
    Private m_triggerHRR As Boolean
    Private m_triggerFO As Boolean
    Private m_triggerVL As Boolean
    Private m_triggerHO As Boolean
    Private m_triggerFR As Boolean
    Private m_cd As Double
    Private m_probventclosed As Single
    Private m_horeliability As Single
    Private m_integrity As Single
    Private m_maxopening As Single
    Private m_maxopeningtime As Single
    Private m_gastemp As Double
    Private m_FRcriteria As Integer
    Private m_enzsimtime As Double
    Private m_enzopen As Boolean

    Public Sub New()
    End Sub

    Public Sub New(ByVal FRcriteria As Integer, ByVal gastemp As Double, ByVal trigger_FR As Boolean, ByVal maxopeningtime As Single, ByVal maxopening As Single, ByVal integrity As Single, ByVal offset As Single, ByVal horeliability As Single, ByVal triggerHO As Boolean, ByVal triggerVL As Boolean, ByVal triggerFO As Boolean, ByVal probventclosed As Single, ByVal cd As Double, ByVal triggerHRR As Boolean, ByVal triggerHD As Boolean, ByVal triggerSD As Boolean, ByVal autoopenvent As Boolean, ByVal triggerventopenduration As Double, ByVal triggerventopendelay As Double, ByVal HRRventopenduration As Double, ByVal HRRventopendelay As Double, ByVal HRRthreshold As Double, ByVal sdtriggerroom As Integer, ByVal spillbalconyprojection As Single, ByVal spillplumemodel As Short, ByVal spillplumebalc As Boolean, ByVal downstand As Single, ByVal glassfalloutime As Single, ByVal glassdistance As Single, ByVal glassflameflux As Boolean, ByVal glassYoungsModulus As Single, ByVal glassalpha As Single, ByVal glassbreakingstress As Single, ByVal glassshading As Single, ByVal glassthickness As Single, ByVal glassexpansion As Single, ByVal glassemissivity As Single, ByVal glassconductivity As Single, ByVal description As String, ByVal fromroom As Integer, ByVal toroom As Integer, ByVal face As Integer, ByVal height As Double, ByVal width As Double, _
                   ByVal sillheight As Double, ByVal opentime As Double, ByVal closetime As Double, ByVal walllength1 As Double, ByVal walllength2 As Double, ByVal autobreakglass As Boolean, ByVal spillplume As Boolean)

        Me.fromroom = fromroom
        Me.toroom = toroom
        Me.face = face
        Me.height = height
        Me.width = width
        Me.sillheight = sillheight
        Me.opentime = opentime
        Me.closetime = closetime
        Me.description = description
        Me.walllength1 = walllength1
        Me.walllength2 = walllength2
        Me.offset = offset
        Me.spillplume = spillplume
        Me.spillplumemodel = spillplumemodel
        Me.spillplumebalc = spillplumebalc
        Me.downstand = downstand
        Me.spillbalconyprojection = spillbalconyprojection
        Me.autobreakglass = autobreakglass
        Me.glassconductivity = glassconductivity
        Me.glassemissivity = glassemissivity
        Me.glassexpansion = glassexpansion
        Me.glassthickness = glassthickness
        Me.glassshading = glassshading
        Me.glassbreakingstress = glassbreakingstress
        Me.glassalpha = glassalpha
        Me.glassYoungsModulus = glassYoungsModulus
        Me.glassflameflux = glassflameflux
        Me.glassdistance = glassdistance
        Me.glassfalloutime = glassfalloutime

        Me.sdtriggerroom = sdtriggerroom
        Me.HRRthreshold = HRRthreshold
        Me.HRRventopendelay = HRRventopendelay
        Me.HRRventopenduration = HRRventopenduration
        Me.triggerventopendelay = triggerventopendelay
        Me.triggerventopenduration = triggerventopenduration
        Me.autoopenvent = autoopenvent
        Me.triggerHD = triggerHD
        Me.triggerSD = triggerSD
        Me.triggerHRR = triggerHRR
        Me.triggerFO = triggerFO
        Me.triggerVL = triggerVL
        Me.triggerHO = triggerHO
        Me.triggerFR = triggerFR
        Me.cd = cd
        Me.probventclosed = probventclosed
        Me.horeliability = horeliability
        Me.integrity = integrity
        Me.maxopening = maxopening
        Me.maxopeningtime = maxopeningtime
        Me.gastemp = gastemp
        Me.FRcriteria = FRcriteria

    End Sub
    Public Property enzsimtime() As Double
        Get
            Return m_enzsimtime
        End Get
        Set(ByVal value As Double)
            m_enzsimtime = value
        End Set
    End Property
    Public Property enzopen() As Boolean
        Get
            Return m_enzopen
        End Get
        Set(ByVal value As Boolean)
            m_enzopen = value
        End Set
    End Property
    Public Property FRcriteria() As Integer
        Get
            Return m_FRcriteria
        End Get
        Set(ByVal value As Integer)
            m_FRcriteria = value
        End Set
    End Property
    Public Property triggerFR() As Boolean
        Get
            Return m_triggerFR
        End Get
        Set(ByVal value As Boolean)
            m_triggerFR = value
        End Set
    End Property
    Public Property triggerVL() As Boolean
        Get
            Return m_triggerVL
        End Get
        Set(ByVal value As Boolean)
            m_triggerVL = value
        End Set
    End Property
    Public Property triggerHO() As Boolean
        Get
            Return m_triggerHO
        End Get
        Set(ByVal value As Boolean)
            m_triggerHO = value
        End Set
    End Property
    Public Property triggerFO() As Boolean
        Get
            Return m_triggerFO
        End Get
        Set(ByVal value As Boolean)
            m_triggerFO = value
        End Set
    End Property
    Public Property maxopeningtime() As Single
        Get
            Return m_maxopeningtime
        End Get
        Set(ByVal value As Single)
            m_maxopeningtime = value
        End Set
    End Property
    Public Property maxopening() As Single
        Get
            Return m_maxopening
        End Get
        Set(ByVal value As Single)
            m_maxopening = value
        End Set
    End Property
    Public Property gastemp() As Double
        Get
            Return m_gastemp
        End Get
        Set(ByVal value As Double)
            m_gastemp = value
        End Set
    End Property
    Public Property integrity() As Single
        Get
            Return m_integrity
        End Get
        Set(ByVal value As Single)
            m_integrity = value
        End Set
    End Property
    Public Property horeliability() As Single
        Get
            Return m_horeliability
        End Get
        Set(ByVal value As Single)
            m_horeliability = value
        End Set
    End Property
    Public Property offset() As Double
        Get
            Return m_offset
        End Get
        Set(ByVal value As Double)
            m_offset = value
        End Set
    End Property
    Public Property probventclosed() As Single
        Get
            Return m_probventclosed
        End Get
        Set(ByVal value As Single)
            m_probventclosed = value
        End Set
    End Property
    Public Property cd() As Double
        Get
            Return m_cd
        End Get
        Set(ByVal value As Double)
            m_cd = value
        End Set
    End Property
    Public Property triggerHRR() As Boolean
        Get
            Return m_triggerHRR
        End Get
        Set(ByVal value As Boolean)
            m_triggerHRR = value
        End Set
    End Property
    Public Property triggerSD() As Boolean
        Get
            Return m_triggerSD
        End Get
        Set(ByVal value As Boolean)
            m_triggerSD = value
        End Set
    End Property
    Public Property triggerHD() As Boolean
        Get
            Return m_triggerHD
        End Get
        Set(ByVal value As Boolean)
            m_triggerHD = value
        End Set
    End Property
    Public Property autoopenvent() As Boolean
        Get
            Return m_autoopenvent
        End Get
        Set(ByVal value As Boolean)
            m_autoopenvent = value
        End Set
    End Property
    Public Property triggerventopenduration() As Double
        Get
            Return m_triggerventopenduration
        End Get
        Set(ByVal value As Double)
            m_triggerventopenduration = value
        End Set
    End Property
    Public Property triggerventopendelay() As Double
        Get
            Return m_triggerventopendelay
        End Get
        Set(ByVal value As Double)
            m_triggerventopendelay = value
        End Set
    End Property
    Public Property HRRventopenduration() As Double
        Get
            Return m_HRRventopenduration
        End Get
        Set(ByVal value As Double)
            m_HRRventopenduration = value
        End Set
    End Property
    Public Property HRRventopendelay() As Double
        Get
            Return m_HRRventopendelay
        End Get
        Set(ByVal value As Double)
            m_HRRventopendelay = value
        End Set
    End Property
    Public Property HRRthreshold() As Double
        Get
            Return m_HRRthreshold
        End Get
        Set(ByVal value As Double)
            m_HRRthreshold = value
        End Set
    End Property
    Public Property sdtriggerroom() As Integer
        Get
            Return m_sdtriggerroom
        End Get
        Set(ByVal value As Integer)
            m_sdtriggerroom = value
        End Set
    End Property
    Public Property spillbalconyprojection() As Single
        Get
            Return m_spillbalconyprojection
        End Get
        Set(ByVal value As Single)
            m_spillbalconyprojection = value
        End Set
    End Property
    Public Property spillplumemodel() As Short
        Get
            Return m_spillplumemodel
        End Get
        Set(ByVal value As Short)
            m_spillplumemodel = value
        End Set
    End Property
    Public Property spillplumebalc() As Boolean
        Get
            Return m_spillplumebalc
        End Get
        Set(ByVal value As Boolean)
            m_spillplumebalc = value
        End Set
    End Property
    Public Property downstand() As Single
        Get
            Return m_downstand
        End Get
        Set(ByVal value As Single)
            m_downstand = value
        End Set
    End Property
    Public Property glassflameflux() As Boolean
        Get
            Return m_glassflameflux
        End Get
        Set(ByVal value As Boolean)
            m_glassflameflux = value
        End Set
    End Property
    Public Property glassfalloutime() As Single
        Get
            Return m_glassfalloutime
        End Get
        Set(ByVal value As Single)
            m_glassfalloutime = value
        End Set
    End Property
    Public Property glassdistance() As Single
        Get
            Return m_glassdistance
        End Get
        Set(ByVal value As Single)
            m_glassdistance = value
        End Set
    End Property
    Public Property glassYoungsModulus() As Single
        Get
            Return m_glassYoungsModulus
        End Get
        Set(ByVal value As Single)
            m_glassYoungsModulus = value
        End Set
    End Property
    Public Property glassalpha() As Single
        Get
            Return m_glassalpha
        End Get
        Set(ByVal value As Single)
            m_glassalpha = value
        End Set
    End Property
    Public Property glassbreakingstress() As Single
        Get
            Return m_glassbreakingstress
        End Get
        Set(ByVal value As Single)
            m_glassbreakingstress = value
        End Set
    End Property
    Public Property glassshading() As Single
        Get
            Return m_glassshading
        End Get
        Set(ByVal value As Single)
            m_glassshading = value
        End Set
    End Property
    Public Property glassthickness() As Single
        Get
            Return m_glassthickness
        End Get
        Set(ByVal value As Single)
            m_glassthickness = value
        End Set
    End Property
    Public Property glassexpansion() As Single
        Get
            Return m_glassexpansion
        End Get
        Set(ByVal value As Single)
            m_glassexpansion = value
        End Set
    End Property
    Public Property glassemissivity() As Single
        Get
            Return m_glassemissivity
        End Get
        Set(ByVal value As Single)
            m_glassemissivity = value
        End Set
    End Property
    Public Property glassconductivity() As Single
        Get
            Return m_glassconductivity
        End Get
        Set(ByVal value As Single)
            m_glassconductivity = value
        End Set
    End Property
    Public Property spillplume() As Boolean
        Get
            Return m_spillplume
        End Get
        Set(ByVal value As Boolean)
            m_spillplume = value
        End Set
    End Property
    Public Property autobreakglass() As Boolean
        Get
            Return m_autobreakglass
        End Get
        Set(ByVal value As Boolean)
            m_autobreakglass = value
        End Set
    End Property
    Public Property walllength2() As Double
        Get
            Return m_walllength2
        End Get
        Set(ByVal value As Double)
            m_walllength2 = value
        End Set
    End Property
    Public Property walllength1() As Double
        Get
            Return m_walllength1
        End Get
        Set(ByVal value As Double)
            m_walllength1 = value
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
    Public Property id() As Integer
        Get
            Return m_id
        End Get
        Set(ByVal value As Integer)
            m_id = value
        End Set
    End Property
    Public Property fromroom() As Integer
        Get
            Return m_fromroom
        End Get
        Set(ByVal value As Integer)
            m_fromroom = value
        End Set
    End Property
    Public Property toroom() As Integer
        Get
            Return m_toroom
        End Get
        Set(ByVal value As Integer)
            m_toroom = value
        End Set
    End Property
    Public Property face() As Integer
        Get
            Return m_face
        End Get
        Set(ByVal value As Integer)
            m_face = value
        End Set
    End Property
    Public Property height() As Double
        Get
            Return m_height
        End Get
        Set(ByVal value As Double)
            m_height = value
        End Set
    End Property
    Public Property width() As Double
        Get
            Return m_width
        End Get
        Set(ByVal value As Double)
            m_width = value
        End Set
    End Property
    Public Property sillheight() As Double
        Get
            Return m_sillheight
        End Get
        Set(ByVal value As Double)
            m_sillheight = value
        End Set
    End Property
    Public Property opentime() As Double
        Get
            Return m_opentime
        End Get
        Set(ByVal value As Double)
            m_opentime = value
        End Set
    End Property
    Public Property closetime() As Double
        Get
            Return m_closetime
        End Get
        Set(ByVal value As Double)
            m_closetime = value
        End Set
    End Property
    Public Function GetDisplayText(ByVal j As String) As String
        Dim text As String

        If toroom <= NumberRooms Then
            text = LeftJustified(j, 2) & " " & _
            RightJustified(description, 24) & _
             RightJustified(fromroom.ToString, 8) & _
            RightJustified(toroom.ToString, 8)
        Else
            text = LeftJustified(j, 2) & " " & _
            RightJustified(description, 24) & _
            RightJustified(fromroom.ToString, 8) & _
           RightJustified("outside", 8)
        End If

        Return text
    End Function
    Public Function Clone() As Object Implements System.ICloneable.Clone

        Dim v As New oVent

        v.description = Me.description
        v.fromroom = Me.fromroom
        v.toroom = Me.toroom
        v.face = Me.face
        v.height = Me.height
        v.width = Me.width
        v.sillheight = Me.sillheight
        v.opentime = Me.opentime
        v.closetime = Me.closetime
        v.description = Me.description
        v.walllength1 = Me.walllength1
        v.walllength2 = Me.walllength2
        v.offset = Me.offset
        v.spillplume = Me.spillplume
        v.spillplumemodel = Me.spillplumemodel
        v.spillplumebalc = Me.spillplumebalc
        v.downstand = Me.downstand
        v.spillbalconyprojection = Me.spillbalconyprojection
        v.autobreakglass = Me.autobreakglass
        v.glassconductivity = Me.glassconductivity
        v.glassemissivity = Me.glassemissivity
        v.glassexpansion = Me.glassexpansion
        v.glassthickness = Me.glassthickness
        v.glassshading = Me.glassshading
        v.glassbreakingstress = Me.glassbreakingstress
        v.glassalpha = Me.glassalpha
        v.glassYoungsModulus = Me.glassYoungsModulus
        v.glassflameflux = Me.glassflameflux
        v.glassdistance = Me.glassdistance
        v.glassfalloutime = Me.glassfalloutime

        v.sdtriggerroom = Me.sdtriggerroom
        v.HRRthreshold = Me.HRRthreshold
        v.HRRventopendelay = Me.HRRventopendelay
        v.HRRventopenduration = Me.HRRventopenduration
        v.triggerventopendelay = Me.triggerventopendelay
        v.triggerventopenduration = Me.triggerventopenduration
        v.autoopenvent = Me.autoopenvent
        v.triggerHD = Me.triggerHD
        v.triggerSD = Me.triggerSD
        v.triggerHRR = Me.triggerHRR
        v.triggerFO = Me.triggerFO
        v.triggerVL = Me.triggerVL
        v.triggerHO = Me.triggerHO
        v.triggerFR = Me.triggerFR
        v.cd = Me.cd
        v.probventclosed = Me.probventclosed
        v.horeliability = Me.horeliability
        v.integrity = Me.integrity
        v.maxopening = Me.maxopening
        v.maxopeningtime = Me.maxopeningtime
        v.gastemp = Me.gastemp
        v.FRcriteria = Me.FRcriteria
        v.enzopen = Me.enzopen
        v.enzsimtime = Me.enzsimtime

        Return v

    End Function
End Class