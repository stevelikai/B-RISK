Imports System.Collections.Generic
Imports System.Math
Public Class oparameter

    Private m_param_name As String
    Private m_param_value As Single
    Private m_param_distribution As String
    Private m_param_mean As Single
    Private m_param_variance As Single
    Private m_param_lbound As Single
    Private m_param_ubound As Single
    Private m_param_alpha As Single
    Private m_param_beta As Single

    Public Sub New()
    End Sub

    Public Sub New(ByVal param_name As String, ByVal param_value As Single, ByVal param_distribution As String, ByVal param_mean As Single, ByVal param_variance As Single, ByVal param_lowerbound As Single, ByVal param_upperbound As Single, ByVal param_alpha As Single, ByVal param_beta As Single)

        Me.pname = param_name
        Me.pvalue = param_value
        Me.pdistribution = param_distribution
        Me.pmean = param_mean
        Me.plbound = param_lowerbound
        Me.pubound = param_upperbound
        Me.palpha = param_alpha
        Me.pbeta = param_beta

    End Sub
    Public Property pbeta() As Single
        Get
            Return m_param_beta
        End Get
        Set(ByVal value As Single)
            m_param_beta = value
        End Set
    End Property
    Public Property palpha() As Single
        Get
            Return m_param_alpha
        End Get
        Set(ByVal value As Single)
            m_param_alpha = value
        End Set
    End Property
    Public Property pubound() As Single
        Get
            Return m_param_ubound
        End Get
        Set(ByVal value As Single)
            m_param_ubound = value
        End Set
    End Property
    Public Property plbound() As Single
        Get
            Return m_param_lbound
        End Get
        Set(ByVal value As Single)
            m_param_lbound = value
        End Set
    End Property
    Public Property pmean() As Single
        Get
            Return m_param_mean
        End Get
        Set(ByVal value As Single)
            m_param_mean = value
        End Set
    End Property
    Public Property pdistribution() As String
        Get
            Return m_param_distribution
        End Get
        Set(ByVal value As String)
            m_param_distribution = value
        End Set
    End Property
    Public Property pvalue() As Single
        Get
            Return m_param_value
        End Get
        Set(ByVal value As Single)
            m_param_value = value
        End Set
    End Property
    Public Property pname() As String
        Get
            Return m_param_name
        End Get
        Set(ByVal value As String)
            m_param_name = value
        End Set
    End Property
End Class
