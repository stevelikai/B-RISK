Imports System.Collections.Generic
Public Class oRoom
    Private m_width As Double
    Private m_length As Double
    Private m_minheight As Double
    Private m_maxheight As Double
    Private m_elevation As Double
    Private m_absx As Single
    Private m_absy As Single
    Private m_cslope As Boolean
    Private m_description As String
    Private m_num As Integer

    Public Sub New()

    End Sub

    Public Sub New(ByVal num As Integer, ByVal description As String, ByVal cslope As Boolean, ByVal absx As Single, ByVal absy As Single, ByVal elevation As Double, ByVal maxheight As Double, ByVal minheight As Double, ByVal length As Double, ByVal width As Double)

        Me.width = width
        Me.length = length
        Me.minheight = minheight
        Me.maxheight = maxheight
        Me.elevation = elevation
        Me.absx = absx
        Me.absy = absy
        Me.cslope = cslope
        Me.description = description
        Me.num = num

    End Sub
    Public Property num() As Integer
        Get
            Return m_num
        End Get
        Set(ByVal value As Integer)
            m_num = value
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
    Public Property cslope() As Boolean
        Get
            Return m_cslope
        End Get
        Set(ByVal value As Boolean)
            m_cslope = value
        End Set
    End Property
    Public Property absx() As Single
        Get
            Return m_absx
        End Get
        Set(ByVal value As Single)
            m_absx = value
        End Set
    End Property
    Public Property absy() As Single
        Get
            Return m_absy
        End Get
        Set(ByVal value As Single)
            m_absy = value
        End Set
    End Property
    Public Property elevation() As Double
        Get
            Return m_elevation
        End Get
        Set(ByVal value As Double)
            m_elevation = value
        End Set
    End Property
    Public Property maxheight() As Double
        Get
            Return m_maxheight
        End Get
        Set(ByVal value As Double)
            m_maxheight = value
        End Set
    End Property
    Public Property minheight() As Double
        Get
            Return m_minheight
        End Get
        Set(ByVal value As Double)
            m_minheight = value
        End Set
    End Property
    Public Property length() As Double
        Get
            Return m_length
        End Get
        Set(ByVal value As Double)
            m_length = value
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
    Public Function GetDisplayText(ByVal sep As String) As String
        Dim text As String

        text = String.Format("{0,-4} {1,-20} {2,10} {3,10} {4,10}", num, description, length, width, maxheight)

        Return text

    End Function
End Class
