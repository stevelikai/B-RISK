Imports System.Collections.Generic
Imports System.Math

Public Class oSmokeDet
    Private m_od As Single
    Private m_sdr As Double
    Private m_sdz As Single
    Private m_sdx As Single
    Private m_sdy As Single
    Private m_sdid As Integer
    Private m_room As Integer
    Private m_responsetime As Single
    Private m_transit As Single
    Private m_charlength As Single
    Private m_sdinside As Boolean

    Private m_sdbeam As Boolean
    Private m_sdbeampathlength As Double
    Private m_sdbeamalarmtrans As Double


    Public Sub New()
    End Sub

    Public Sub New(ByVal room As Integer, ByVal sdx As Single, ByVal sdy As Single, ByVal sdz As Single, ByVal sdr As Double, ByVal od As Single, ByVal charlength As Single, ByVal inside As Boolean, ByVal sdbeam As Boolean, ByVal sdbeampathlength As Double, ByVal sdbeamalarmtrans As Double)

        Me.sdx = sdx
        Me.sdy = sdy
        Me.room = room
        Me.responsetime = Nothing
        Me.od = od
        Me.sdr = sdr
        Me.sdz = sdz
        Me.sdinside = sdinside
        Me.charlength = charlength
        Me.transit = Nothing
        Me.sdbeam = sdbeam
        Me.sdbeampathlength = sdbeampathlength
        Me.sdbeamalarmtrans = sdbeamalarmtrans

    End Sub
    Public Property sdbeampathlength() As Double
        Get
            Return m_sdbeampathlength
        End Get
        Set(ByVal value As Double)
            m_sdbeampathlength = value
        End Set
    End Property
    Public Property sdbeamalarmtrans() As Double
        Get
            Return m_sdbeamalarmtrans
        End Get
        Set(ByVal value As Double)
            m_sdbeamalarmtrans = value
        End Set
    End Property
    Public Property sdinside() As Boolean
        Get
            Return m_sdinside
        End Get
        Set(ByVal value As Boolean)
            m_sdinside = value
        End Set
    End Property
    Public Property sdbeam() As Boolean
        Get
            Return m_sdbeam
        End Get
        Set(ByVal value As Boolean)
            m_sdbeam = value
        End Set
    End Property
    Public Property charlength() As Single
        Get
            Return m_charlength
        End Get
        Set(ByVal value As Single)
            m_charlength = value
        End Set
    End Property
    Public Property responsetime() As Single
        Get
            Return m_responsetime
        End Get
        Set(ByVal value As Single)
            m_responsetime = value
        End Set
    End Property
    Public Property transit() As Single
        Get
            Return m_transit
        End Get
        Set(ByVal value As Single)
            m_transit = value
        End Set
    End Property
    Public Property od() As Single
        Get
            Return m_od
        End Get
        Set(ByVal value As Single)
            m_od = value
        End Set
    End Property
    Public Property room() As Single
        Get
            Return m_room
        End Get
        Set(ByVal value As Single)
            m_room = value
        End Set
    End Property
    Public Property sdx() As Single
        Get
            Return m_sdx
        End Get
        Set(ByVal value As Single)
            m_sdx = value
        End Set
    End Property
    Public Property sdy() As Single
        Get
            Return m_sdy
        End Get
        Set(ByVal value As Single)
            m_sdy = value
        End Set
    End Property
    Public Property sdz() As Single
        Get
            Return m_sdz
        End Get
        Set(ByVal value As Single)
            m_sdz = value
        End Set
    End Property
    Public Property sdr() As Double
        Get
            Return m_sdr
        End Get
        Set(ByVal value As Double)
            m_sdr = value
        End Set
    End Property
    Public Property sdid() As Integer
        Get
            Return m_sdid
        End Get
        Set(ByVal value As Integer)
            m_sdid = value
        End Set
    End Property

    Public Function GetDisplayText(ByVal sep As String) As String
        Dim text As String
        'text = room & sep & VB6.Format(sdx, "0.000") & sep & VB6.Format(sdy, "0.000")

        If sdbeam = -1 Then
            text = room & sep & Format(sdbeam, "beam")
        Else
            text = room & sep & Format(sdbeam, "point")
        End If


        Return text
    End Function

End Class

