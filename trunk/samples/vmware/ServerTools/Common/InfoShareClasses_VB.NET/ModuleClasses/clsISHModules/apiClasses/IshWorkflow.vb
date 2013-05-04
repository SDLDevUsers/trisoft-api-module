Imports System.Xml
Imports ErrorHandlerNS

Public Class IshWorkflow
    Inherits mainAPIclass
#Region "Private Members"
    Private ReadOnly strModuleName As String = "IshWorkFlow"
#End Region
#Region "Constructors"
    Public Sub New(ByVal Username As String, ByVal Password As String, ByVal ServerURL As String)
        oISHAPIObjs = New ISHObjs(Username, Password, ServerURL)
        oISHAPIObjs.ISHAppObj.Login("InfoShareAuthor", Username, Password, Context)
    End Sub
#End Region
#Region "Properties"

#End Region
#Region "Methods"

#End Region

End Class
