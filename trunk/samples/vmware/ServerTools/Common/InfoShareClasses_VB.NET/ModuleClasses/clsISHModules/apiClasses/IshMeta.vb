Imports System.Xml
Imports ErrorHandlerNS

Public Class IshMeta
    Inherits mainAPIclass
#Region "Private Members"
    Private ReadOnly strModuleName As String = "ISHMeta"
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

    ''' <summary>
    ''' Determines if a user has priviledges of a particular role and belongs to a specified group.
    ''' </summary>
    ''' <param name="Username">Username in the CMS.</param>
    ''' <param name="Role">Role priviledge such as "Administrator", "Author", "Illustrator", etc.</param>
    ''' <param name="UserGroup">(OPTIONAL) The group to search within ("TCL", for instance).</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UserHasRole(ByVal Username As String, ByVal Role As String, Optional ByVal UserGroup As String = "TCL") As Boolean
        Dim returneduserlist() As String
        Dim userlist As New ArrayList

        oISHAPIObjs.ISHMetaObj.GetUsers(Context, Role, UserGroup, returneduserlist)

        For Each uname As String In returneduserlist
            userlist.Add(uname)
        Next
        If userlist.Contains(Username) Then
            Return True
        Else
            Return False
        End If


    End Function
#End Region
End Class
