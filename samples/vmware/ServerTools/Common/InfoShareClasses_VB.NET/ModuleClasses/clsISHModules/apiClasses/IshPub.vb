Imports System.Xml
Imports ErrorHandlerNS
Public Class IshPub
    Inherits mainAPIclass
#Region "Private Members"
    Private ReadOnly strModuleName As String = "ISHPub"
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
    ''' Saves all objects from a specified publication, version, and language
    ''' </summary>
    ''' <param name="PubGUID"></param>
    ''' <param name="PubVer"></param>
    ''' <param name="Language"></param>
    ''' <param name="SavePath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExportPublicationbyBaseline(ByVal PubGUID As String, ByVal PubVer As String, ByVal Language As String, ByVal SavePath As String) As Boolean
        'Get the baseline objects
        Dim myBaseline As Dictionary(Of String, CMSObject)
        myBaseline = GetBaselineObjects(PubGUID, PubVer, Language)
        Dim CurRes As New ArrayList
        CurRes.Add("High")
        CurRes.Add("Low")
        'for each baseline object, save the files to the specified path (getobjbyid with path)
        For Each myObject As KeyValuePair(Of String, CMSObject) In myBaseline
            If myObject.Value.IshType = "ISHIllustration" Then
                For Each resolution As String In CurRes
                    If ObjectExists(myObject.Value.GUID, myObject.Value.Version, Language, resolution) Then
                        GetObjByID(myObject.Value.GUID, myObject.Value.Version, Language, resolution, SavePath)
                    End If
                Next
            Else
                GetObjByID(myObject.Value.GUID, myObject.Value.Version, Language, "", SavePath)
            End If

        Next
    End Function

    ''' <summary>
    ''' Resets all map and topic title properties to match their title elements. 
    ''' </summary>
    ''' <param name="PubGUID"></param>
    ''' <param name="PubVer"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UpdateAllTitleProperties(ByVal PubGUID As String, ByVal PubVer As String) As Boolean
        Dim dictBaseLine As Dictionary(Of String, CMSObject) = GetBaselineObjects(PubGUID, PubVer)
        For Each entry As KeyValuePair(Of String, CMSObject) In dictBaseLine
            If entry.Value.IshType = "ISHModule" Or entry.Value.IshType = "ISHMasterDoc" Or entry.Value.IshType = "ISHLibrary" Then
                UpdateTitleProperty(entry.Key, entry.Value.Version)
            End If
        Next
        Return True
    End Function

    

    Private Function GetPubObjByID(ByVal GUID As String, ByVal Version As String) As XmlDocument
        Dim MyNode As XmlNode = Nothing
        Dim MyDoc As New XmlDocument
        Dim MyMeta As New XmlDocument
        Dim XMLString As String = ""
        Dim ISHMeta As String = ""
        Dim ISHResult As String = ""

        'Call the CMS to get our content!
        Try
            ISHResult = oISHAPIObjs.ISHPubOutObj25.Find(Context, _
                                                         PublicationOutput25.eISHStatusgroup.ISHNoStatusFilter, _
                                                         oCommonFuncs.BuildPubMetaDataFilter(GUID, Version).ToString, _
                                                         oCommonFuncs.BuildFullPubMetadata.ToString, _
                                                         XMLString)
        Catch ex As Exception
            'modErrorHandler.Errors.PrintMessage(3, "Failed to retrieve XML from CMS server: " + ex.Message, strModuleName)
            Return Nothing
        End Try

        'Load the XML and get the metadata:
        Try
            MyDoc.LoadXml(XMLString)
            Return MyDoc
        Catch ex As Exception
            modErrorHandler.Errors.PrintMessage(3, "Failed to retrieve object metadata from returned XML: " + ex.Message, strModuleName)
            Return Nothing
        End Try
    End Function

    
#End Region
End Class
