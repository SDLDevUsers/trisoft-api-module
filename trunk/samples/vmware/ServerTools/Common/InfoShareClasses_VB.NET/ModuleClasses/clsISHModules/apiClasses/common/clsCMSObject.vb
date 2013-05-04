Public Class CMSObject
    Private _GUID As String
    Private _Version As String
    Private _IshType As String
    Private _ReportedItems As String
    Public Property GUID As String
        Get
            Return _GUID
        End Get
        Set(ByVal value As String)
            _GUID = value
        End Set
    End Property
    Public Property Version As String
        Get
            Return _Version
        End Get
        Set(ByVal value As String)
            _Version = value
        End Set
    End Property
    Public Property IshType As String
        Get
            Return _IshType
        End Get
        Set(ByVal value As String)
            _IshType = value
        End Set
    End Property
    Public Property ReportedItems As String
        Get
            Return _ReportedItems
        End Get
        Set(ByVal value As String)
            _ReportedItems = value
        End Set
    End Property

    Public Sub New(ByVal strGUID As String, ByVal strVersion As String, ByVal strIshType As String, Optional ByVal strReportedItems As String = "<reporteditems/>")
        _GUID = strGUID
        _Version = strVersion
        _IshType = strIshType
        _ReportedItems = strReportedItems
    End Sub
End Class
