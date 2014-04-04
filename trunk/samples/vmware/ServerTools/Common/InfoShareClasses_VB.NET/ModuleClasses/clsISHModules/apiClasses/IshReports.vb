Imports System.Xml
Imports ErrorHandlerNS
Public Class IshReports
    Inherits mainAPIclass
#Region "Private Members"

#End Region
#Region "Constructors"
    Public Sub New(ByVal Username As String, ByVal Password As String, ByVal ServerURL As String)
        'Make sure to use the FQDN up to the "WS" portion of your URL: "https://yourserver/InfoShareWS"
        oISHAPIObjs = New ISHObjs(Username, Password, ServerURL)
        'oISHAPIObjs.ISHAppObj.Login("InfoShareAuthor", Username, Password, Context)
    End Sub
#End Region
#Region "Properties"
    Private ReadOnly strModuleName As String = "IshReports"
#End Region
#Region "Methods"
   
#End Region
End Class
