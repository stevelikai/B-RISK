Imports System.Collections.Generic
Imports System.Math

Public Class oFan
    Private m_fanflowrate As Single
    Private m_fanelevation As Single
    Private m_fanpressurelimit As Single
    Private m_fanstarttime As Single
    Private m_fanreliability As Single
    Private m_fanroom As Integer
    Private m_fanid As Integer
    Private m_fanextract As Boolean
    Private m_fanmanual As Integer
    Private m_fancurve As Boolean


    Public Sub New()
    End Sub

    Public Sub New(ByVal fanreliability As Single, ByVal fanroom As Integer, ByVal fanid As Integer, ByVal fanflowrate As Single, ByVal fanelevation As Single, ByVal fanpressurelimit As Single, ByVal fanstarttime As Single, ByVal fanextract As Boolean, ByVal fanmanual As Integer, ByVal fancurve As Boolean)

        Me.fanflowrate = fanflowrate
        Me.fanelevation = fanelevation
        Me.fanroom = fanroom
        Me.fanpressurelimit = fanpressurelimit
        Me.fanstarttime = fanstarttime
        Me.fanid = fanid
        Me.fanextract = fanextract
        Me.fanmanual = fanmanual
        Me.fancurve = fancurve
        Me.fanreliability = fanreliability

    End Sub
    Public Property fancurve() As Boolean
        Get
            Return m_fancurve
        End Get
        Set(ByVal value As Boolean)
            m_fancurve = value
        End Set
    End Property
    Public Property fanmanual() As Integer
        Get
            Return m_fanmanual
        End Get
        Set(ByVal value As Integer)
            m_fanmanual = value
        End Set
    End Property
    Public Property fanextract() As Boolean
        Get
            Return m_fanextract
        End Get
        Set(ByVal value As Boolean)
            m_fanextract = value
        End Set
    End Property
    Public Property fanreliability() As Single
        Get
            Return m_fanreliability
        End Get
        Set(ByVal value As Single)
            m_fanreliability = value
        End Set
    End Property
    Public Property fanflowrate() As Single
        Get
            Return m_fanflowrate
        End Get
        Set(ByVal value As Single)
            m_fanflowrate = value
        End Set
    End Property
    Public Property fanelevation() As Single
        Get
            Return m_fanelevation
        End Get
        Set(ByVal value As Single)
            m_fanelevation = value
        End Set
    End Property
    Public Property fanpressurelimit() As Single
        Get
            Return m_fanpressurelimit
        End Get
        Set(ByVal value As Single)
            m_fanpressurelimit = value
        End Set
    End Property
    Public Property fanstarttime() As Single
        Get
            Return m_fanstarttime
        End Get
        Set(ByVal value As Single)
            m_fanstarttime = value
        End Set
    End Property
    Public Property fanroom() As Integer
        Get
            Return m_fanroom
        End Get
        Set(ByVal value As Integer)
            m_fanroom = value
        End Set
    End Property
    Public Property fanid() As Integer
        Get
            Return m_fanid
        End Get
        Set(ByVal value As Integer)
            m_fanid = value
        End Set
    End Property

    Public Function GetDisplayText(ByVal sep As String) As String
        Dim text As String
        text = fanroom & sep & VB6.Format(fanflowrate, "0.000") & sep & sep & VB6.Format(fanelevation, "0.000")
        Return text
    End Function

End Class
