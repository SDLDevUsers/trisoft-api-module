Imports System.Xml
Imports System
Imports ErrorHandlerNS
Imports System.IO
Imports System.Text

Public Class IshPubOutput
    Inherits mainAPIclass
#Region "Private Members"
    Private ReadOnly strModuleName As String = "IshPubOutput"
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
    Public Function DownloadOutput(ByVal PubGUID As String, ByVal PubVer As String, ByVal OutLang As String, ByVal OutType As String, ByVal myFolder As String) As String
        Dim strEdtID As String = ""
        Dim strEdtType As String = ""
        Dim strFileSize As String = ""
        Dim strMimeType As String = ""
        Dim strFileExt As String = ""
        Dim lngFileSize As Long = 0
        Dim lngIshLngRef As Long = 0
        Dim strPubTitle As String = ""
        Dim strBuildUser As String = ""
        Dim strBuildDate As String = ""
        Dim strPubServ As String = ""
        Try
            GetOutputInfo(PubGUID, PubVer, OutLang, OutType, strEdtID, strEdtType, strFileSize, strMimeType, strFileExt, lngIshLngRef, strPubTitle, strBuildDate, strBuildUser, strPubServ)
            lngFileSize = Convert.ToInt64(strFileSize)
            Dim remaining As Long = lngFileSize
            Dim plOff As Long = 0
            Dim chunks() As Byte = {}

            Dim chunk_size As Long = 256000
            'gather chunks until file is complete
            While (remaining > 0)


                'This is to make sure we don't ask for a bigger chunk than there is left
                If (chunk_size > remaining) Then
                    chunk_size = remaining
                End If
                Dim pboutbytes() As Byte


                oISHAPIObjs.ISHPubOutObj25.GetNextDataObjectChunkByIshLngRef(Context, lngIshLngRef, strEdtID, plOff, chunk_size, pboutbytes)


                remaining = lngFileSize - plOff
                'No need to update the offset, it appears GetNextDataObjectChunk... does it automatically.
                'plOff = plOff + chunk_size
                'append new chunk to current chunks
                Dim byteList As List(Of Byte) = New List(Of Byte)(chunks)
                byteList.AddRange(pboutbytes)
                Dim byteArrayAll() As Byte = byteList.ToArray
                chunks = byteArrayAll

            End While
            'Create the storage folder:
            If Not Directory.Exists(myFolder) Then
                Directory.CreateDirectory(myFolder)
            End If
            'Default CMS naming convention:
            'Lists and Steps=1=PDF - Press size with registration marks=en.pdf
            strPubTitle = oCommonFuncs.RemoveWindowsIllegalChars(strPubTitle)
            Dim filename As String = strPubTitle & "=" & PubVer & "=" & OutType & "=" & OutLang & "." & strFileExt
            Dim fullfilepath As String = ""
            Try
                fullfilepath = myFolder + filename
                My.Computer.FileSystem.WriteAllBytes(fullfilepath, chunks, False)
            Catch ex As Exception
                fullfilepath = myFolder + PubGUID & "=" & PubVer & "=" & OutType & "=" & OutLang & "." & strFileExt
                My.Computer.FileSystem.WriteAllBytes(fullfilepath, chunks, False)
            End Try
            Return fullfilepath
        Catch ex As Exception
            modErrorHandler.Errors.PrintMessage(3, "Failure downloading content for specified output: " + PubGUID + "-v" + PubVer + "-" + OutLang + "-" + OutType & ". Reason: " & ex.Message, strModuleName + "-DownloadOutput")
            Return False
        End Try

        '// these are php header declarations for page load over http
        'header('Content-Type: ' . $mime);
        'header('Content-Disposition: attachment; filename="FILENAME' . '.'.$fileextension.'"');

        '// printing total $chunks to web page
        'echo $chunks;

    End Function
    ''' <summary>
    ''' ''' Gets an output's metadata along with a lot of file-specific info for downloading the file.
    ''' </summary>
    ''' <param name="PubGUID">Publication GUID</param>
    ''' <param name="PubVer">Publication Version</param>
    ''' <param name="OutLang">Language of the output</param>
    ''' <param name="OutputType">Output type (PDF - Online, WebWorks, etc.)</param>
    ''' <param name="outEdGUID"></param>
    ''' <param name="outEDTType"></param>
    ''' <param name="outFileSize">Returns the exact size of the file in bytes.</param>
    ''' <param name="outMimeType"></param>
    ''' <param name="outFileExt"></param>
    ''' <param name="outIshLngRef"></param>
    ''' <param name="outPubTitle"></param>
    ''' <param name="outBuildDate"></param>
    ''' <param name="outBuildUser"></param>
    ''' <param name="OutPubServ"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetOutputInfo(ByVal PubGUID As String, ByVal PubVer As String, ByVal OutLang As String, ByVal OutputType As String, ByRef outEdGUID As String, ByRef outEDTType As String, ByRef outFileSize As String, ByRef outMimeType As String, ByRef outFileExt As String, ByRef outIshLngRef As Long, ByRef outPubTitle As String, ByRef outBuildDate As String, ByRef outBuildUser As String, ByRef OutPubServ As String) As Boolean
        'Get publication title:
        Dim myfilter As String = "<ishfields><ishfield name=""FMAPID"" level=""lng"">" & PubGUID & "</ishfield><ishfield name=""DOC-LANGUAGE"" level=""lng"">" & OutLang & "</ishfield></ishfields>"
        Dim myrequest As String = "<ishfields><ishfield name=""VERSION"" level=""version""/><ishfield name=""FISHPUBSTATUS"" level=""lng""/><ishfield name=""FISHISRELEASED"" level=""version""/><ishfield name=""DOC-LANGUAGE"" level=""lng""/><ishfield name=""FISHOUTPUTFORMATREF"" level=""lng""/><ishfield name=""FTITLE"" level=""logical""/></ishfields>"
        Dim responsexml As String = ""
        Try
            oISHAPIObjs.ISHPubOutObj25.Find(Context, PublicationOutput25.eISHStatusgroup.ISHNoStatusFilter, myfilter, myrequest, responsexml)
            Dim pubinfo As New XmlDocument
            pubinfo.LoadXml(responsexml)
            outPubTitle = pubinfo.SelectSingleNode("//ishfield[@name='FTITLE']").InnerText
        Catch ex As Exception
            outPubTitle = PubGUID
        End Try



        'Get output's plLngRef num:
        Dim alloutputs As XmlDocument

        alloutputs = GetPubOutputsByISHRef(PubGUID, PubVer, OutLang, OutputType)
        If IsNothing(alloutputs) Then
            Return Nothing
        End If
        Dim strMyISHLangRef As String
        Dim lngMyISHLangRef As Long
        Dim outputs As XmlNodeList = alloutputs.SelectNodes("//ishobject")
        'For any returned outputs, get the ishlangref and convert it to a Long integer.
        For Each myoutput As XmlNode In outputs
            lngMyISHLangRef = 0
            strMyISHLangRef = ""
            strMyISHLangRef = myoutput.Attributes.GetNamedItem("ishlngref").InnerText
            lngMyISHLangRef = Convert.ToInt64(strMyISHLangRef)
        Next

        'Get who and when output info:
        Dim strWhoWhenRequest As String = "<ishfields><ishfield name=""VERSION"" level=""version""/><ishfield name=""DOC-LANGUAGE"" level=""lng""/><ishfield name=""FISHPUBLNGCOMBINATION"" level=""lng""/><ishfield name=""FISHOUTPUTFORMATREF"" level=""lng""/><ishfield name=""FISHEVENTID"" level=""lng""/><ishfield name=""FISHPUBLISHER"" level=""lng""/></ishfields>"
        Dim strWhoWhenResult As String = ""
        oISHAPIObjs.ISHPubOutObj25.GetMetaDataByIshLngRef(Context, lngMyISHLangRef, strWhoWhenRequest, strWhoWhenResult)
        Dim whowhen As New XmlDocument
        whowhen.LoadXml(strWhoWhenResult)


        'Get output's metadata
        Dim requesteddata As String = ""
        oISHAPIObjs.ISHPubOutObj25.GetDataObjectInfoByIshLngRef(Context, lngMyISHLangRef, requesteddata)
        Dim metadata As New XmlDocument
        metadata.LoadXml(requesteddata)




        Try
            outBuildUser = whowhen.SelectSingleNode("//ishfield[@name='FISHPUBLISHER']").InnerText
            Dim EventID As String = whowhen.SelectSingleNode("//ishfield[@name='FISHEVENTID']").InnerText
            'Split the event ID and return the various parts: e.g.: "593 cms-dev-app 20110518 11:00:46" where "<eventid> <pubServ> <Date> <Time>
            'TODO: Note that we're dropping the event ID here. Might be useful for sys admin to know that but most users don't need it.
            Dim eventids As String() = Nothing
            eventids = EventID.Split(" ")
            OutPubServ = eventids(1)
            outBuildDate = eventids(2) + " " + eventids(3)
            outBuildDate = outBuildDate.Insert(6, "-")
            outBuildDate = outBuildDate.Insert(4, "-")
            outEdGUID = metadata.SelectSingleNode("/ishdataobjects/ishdataobject").Attributes.GetNamedItem("ed").InnerText.ToString
            outEDTType = metadata.SelectSingleNode("/ishdataobjects/ishdataobject").Attributes.GetNamedItem("edt").InnerText.ToString
            outFileExt = metadata.SelectSingleNode("/ishdataobjects/ishdataobject").Attributes.GetNamedItem("fileextension").InnerText.ToString
            outFileSize = metadata.SelectSingleNode("/ishdataobjects/ishdataobject").Attributes.GetNamedItem("size").InnerText.ToString
            outMimeType = metadata.SelectSingleNode("/ishdataobjects/ishdataobject").Attributes.GetNamedItem("mimetype").InnerText.ToString
            outIshLngRef = lngMyISHLangRef
            Return (True)
        Catch ex As Exception
            modErrorHandler.Errors.PrintMessage(3, "Unable to find output for specified GUID: " + PubGUID + "-v" + PubVer + "-" + OutLang + "-" + OutputType & ". Reason: " & ex.Message, strModuleName + "-GetOutputInfo")
            Return False
        End Try
    End Function



    Public Function GetOutputState(ByVal PubGUID As String, ByVal Version As String, ByVal Language As String, ByVal OutputType As String) As String
        'Get the output:
        Dim alloutputs As XmlDocument

        alloutputs = GetPubOutputsByISHRef(PubGUID, Version, Language, OutputType)
        If IsNothing(alloutputs) Then
            Return Nothing
        End If
        Dim strMyISHLangRef As String
        Dim lngMyISHLangRef As Long
        Dim outputs As XmlNodeList = alloutputs.SelectNodes("//ishobject")
        'Should only return one output, but just in case, let's log an error if there are more than one found.
        If outputs.Count > 1 Then
            modErrorHandler.Errors.PrintMessage(3, "Multiple outputs found. Expected one unique output. PubGUID: " + PubGUID, strModuleName + "-GetOutputState")
            Return Nothing
        End If

        Dim Status As String = "UNKNOWN"
        For Each myoutput As XmlNode In outputs
            lngMyISHLangRef = 0
            strMyISHLangRef = ""
            strMyISHLangRef = myoutput.Attributes.GetNamedItem("ishlngref").InnerText
            lngMyISHLangRef = Convert.ToInt64(strMyISHLangRef)
            'Build a requesting filter.
            Dim requestedMeta As String = "<ishfields><ishfield name=""FISHPUBSTATUS"" level=""lng""/></ishfields>"
            'Place to store resulting requested info.
            Dim strRequestedObjects As String = ""
            'Get the status of the requested output.
            oISHAPIObjs.ISHPubOutObj25.GetMetaDataByIshLngRef(Context, lngMyISHLangRef, requestedMeta, strRequestedObjects)
            '.ISHPubOutObj25.GetMetaDataByIshLngRef(Context, lngMyISHLangRef, requestedMeta, strRequestedObjects)
            Dim mydoc As New XmlDocument
            mydoc.LoadXml(strRequestedObjects)
            'Record the state:
            Try
                Status = mydoc.SelectSingleNode("//ishfield[@name=""FISHPUBSTATUS""]").InnerText
            Catch ex As Exception
                Status = "NOTFOUND"
            End Try

        Next
        Return Status



    End Function
    Public Function CanBePublished(ByVal IshLangRef As Long) As Boolean
        'Build a requesting filter.
        Dim requestedMeta As String = "<ishfields><ishfield name=""FISHPUBSTATUS"" level=""lng""/></ishfields>"
        'Place to store resulting requested info.
        Dim strRequestedObjects As String = ""
        'Build a list of forbidden states that don't allow publishing.
        Dim mylist As New List(Of String)
        mylist.Add("Publish Pending")
        mylist.Add("Publishing")
        mylist.Add("Released")
        'Get the status of the requested output.
        oISHAPIObjs.ISHPubOutObj25.GetMetaDataByIshLngRef(Context, IshLangRef, requestedMeta, strRequestedObjects)
        Dim mydoc As New XmlDocument
        mydoc.LoadXml(strRequestedObjects)
        Dim Status As String

        Status = mydoc.SelectSingleNode("//ishfield[@name=""FISHPUBSTATUS""]").InnerText

        'Check to see if our status won't allow publishing.
        If mylist.Contains(Status) Then
            Return False
        Else
            Return True
        End If
    End Function

    ''' <summary>
    ''' Get all the outputs according to the specified metadata being used as a filter.  
    ''' If a filter is left out, it will grab all outputs regardless of what that meta is set to, except for Version.
    ''' If Version is not specified, the latest will be used.
    ''' Start any returned outputs if they can, report any that can't be started.
    ''' </summary>
    ''' <param name="PubGUID">Publication GUID</param>
    ''' <param name="Version">Version</param>
    ''' <param name="Language">Output Language</param>
    ''' <param name="OutputType">Output Type</param>
    ''' <returns>True or False</returns>
    ''' <remarks></remarks>
    Public Function StartPubOutput(ByVal PubGUID As String, Optional ByVal Version As String = "latest", Optional ByVal Language As String = "", Optional ByVal OutputType As String = "all") As Boolean
        Dim alloutputs As XmlDocument

        alloutputs = GetPubOutputsByISHRef(PubGUID, Version, Language, OutputType)
        If IsNothing(alloutputs) Then
            Return Nothing
        End If
        Dim strMyISHLangRef As String
        Dim lngMyISHLangRef As Long
        Dim outputs As XmlNodeList = alloutputs.SelectNodes("//ishobject")
        'For any returned outputs, start the publishing process or report any that can't be started for whatever reason.
        For Each myoutput As XmlNode In outputs
            lngMyISHLangRef = 0
            strMyISHLangRef = ""
            strMyISHLangRef = myoutput.Attributes.GetNamedItem("ishlngref").InnerText
            lngMyISHLangRef = Convert.ToInt64(strMyISHLangRef)
            Dim strOutEventID As String = ""
            Try
                If CanBePublished(lngMyISHLangRef) Then
                    oISHAPIObjs.ISHPubOutObj20.PublishByIshLngRef(Context, lngMyISHLangRef, strOutEventID)
                Else
                    modErrorHandler.Errors.PrintMessage(2, "Output is already printing or is released. PubGUID: " + PubGUID + " ISHLangRef of output: " + lngMyISHLangRef.ToString, strModuleName + "-StartAllPubOutputs")
                End If
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(3, "Failure while attempting to start an output. PubGUID: " + PubGUID + " ISHLangRef of output: " + lngMyISHLangRef.ToString, strModuleName + "-StartAllPubOutputs")
            End Try

        Next
        Return True
    End Function

    ''' <summary>
    ''' Gets an XML list of all publishing outputs of a given publication GUID. If no version is specified, the latest version is used.  Likewise, if no language is specified, all languages are retrieved.
    ''' </summary>
    ''' <param name="GUID">Publication GUID</param>
    ''' <param name="Version">Publication Version</param>
    ''' <param name="Language">Publication Language</param>
    ''' <returns>XmlDocument</returns>
    ''' <remarks></remarks>
    Public Function GetPubOutputsByISHRef(ByVal GUID As String, Optional ByVal Version As String = "latest", Optional ByVal Language As String = "", Optional ByVal OutputType As String = "all") As XmlDocument
        Dim strXMLMetaDataFilter As StringBuilder = oCommonFuncs.BuildPubMetaDataFilter(GUID)
        Dim strXMLRequestedMetadata As StringBuilder = oCommonFuncs.BuildFullPubMetadata()
        Dim strOutXMLObjList As String
        Dim GUIDs(0) As String
        GUIDs(0) = GUID
        If Version = "latest" Then
            Version = GetLatestPubVersionNumber(GUID)
        End If
        Dim requestedmeta As String = oCommonFuncs.BuildFullPubMetadata.ToString
        Dim metafilter As String = oCommonFuncs.BuildPubMetaDataFilter(GUID, Version, Language, OutputType).ToString
        Try
            Dim result As String = oISHAPIObjs.ISHPubOutObj25.RetrieveMetadata(Context, _
                                                               GUIDs, _
                                                               PublicationOutput25.eISHStatusgroup.ISHNoStatusFilter, _
                                                               metafilter, _
                                                               requestedmeta, _
                                                               strOutXMLObjList)
            'ISHPubOutObj25.Find(Context, PublicationOutput25.eISHStatusgroup.ISHNoStatusFilter, strXMLMetaDataFilter.ToString, strXMLRequestedMetadata.ToString, strOutXMLObjList)
            Dim ListofObjects As New XmlDocument()
            ListofObjects.LoadXml(strOutXMLObjList)
            Return ListofObjects
        Catch ex As Exception
            modErrorHandler.Errors.PrintMessage(2, "Error retrieving metadata while getting pub outputs. Message: " + ex.Message, strModuleName + "-GetPubOutputsByISHRef")
            Return Nothing
        End Try

    End Function
    Public Function CancelPublishOperation(ByVal strPubGUID As String, ByVal strPubVer As String, ByVal strOutType As String, ByVal strLanguage As String) As Boolean

        Dim EdGuid As String = ""
        Dim EDTType As String = ""
        Dim FileSize As Long = 0
        Dim mimetype As String = ""
        Dim fileext As String = ""
        Dim ishlngref As Long = 0
        Dim PubTitle As String = ""
        Dim state As String = ""
        Dim strBuildUser As String = ""
        Dim strBuildDate As String = ""
        Dim strPubServ As String = ""
        Try
            state = GetOutputState(strPubGUID, strPubVer, strLanguage, strOutType)
            If state = "Pending" Or state = "Publishing" Then
                GetOutputInfo(strPubGUID, strPubVer, strLanguage, strOutType, EdGuid, EDTType, FileSize, mimetype, fileext, ishlngref, PubTitle, strBuildDate, strBuildUser, strPubServ)
                oISHAPIObjs.ISHPubOutObj20.CancelPublishByIshLngRef(Context, ishlngref)
            End If
        Catch ex As Exception
            modErrorHandler.Errors.PrintMessage(2, "Unable to cancel publishing on specified output. Skipping. Message: " + ex.Message, strModuleName + "-CancelPublishOperation")
        End Try
    End Function
#End Region
End Class
