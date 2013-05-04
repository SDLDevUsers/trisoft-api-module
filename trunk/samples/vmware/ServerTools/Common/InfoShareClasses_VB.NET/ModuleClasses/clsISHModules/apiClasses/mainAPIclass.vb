Imports System
Imports System.Xml
Imports System.IO
Imports System.Xml.XmlReader
Imports System.Text

Imports System.Collections
Imports Microsoft.VisualBasic.ControlChars
Imports Microsoft.VisualBasic.FileIO
Imports Microsoft.VisualBasic.FileIO.FileSystem
Imports System.Convert
Imports System.Text.RegularExpressions
Imports ErrorHandlerNS

Public Class mainAPIclass
    'This class contains:
    '1. An explicit object that hooks up all of the default APIs as objects
    '2. Shared methods used across multiple subclasses.
    'All sub API classes inherit this class to include the mutually used above two pieces.
    'Use of the subclasses must be done explictly.  All universally needed functionality
    'should be moved out to the instantiated class and shared there, not in these subs.

#Region "Private Members"
    Private ReadOnly strModuleName As String = "CustomCMSFuncs"
#End Region
#Region "Constructors"
    Sub New()
        'Can be overridden by subs. Should report an error if ever called.
        'Do nothing.
    End Sub

#End Region
#Region "Properties"
    Public oISHAPIObjs As ISHObjs
    Public Context As String = ""
    Public oCommonFuncs As New clsCommonFuncs
    Public DeletedGUIDs As New ArrayList
    Public DeleteFailedGUIDs As New ArrayList

    ''' <summary>
    ''' Structure used to store common metadata from returned CMS objects.
    ''' </summary>
    Public Structure ObjectData
        Dim GUID As String
        Dim IshType As String
        Dim Title As String
        Dim Version As String
        Dim Language As String
        Dim Status As String
        Dim Resolution As String
        Dim MetaData As XmlNode
    End Structure

#End Region
#Region "Shared Methods"
#Region "Baseline Funcs"
    Public Function GetBaselineObjects(ByVal GUID As String, ByVal Version As String, Optional ByVal Language As String = "en") As Dictionary(Of String, CMSObject)
        'Get the pub's baseline ID from the pub object
        Dim outObjectList As String = ""
        Dim GUIDs(0) As String
        GUIDs(0) = GUID
        Dim Languages(0) As String
        Languages(0) = Language
        Dim Resolutions() As String
        oISHAPIObjs.ISHMetaObj.GetLOVValues(Context, "DRESOLUTION", Resolutions)

        'Get the existing publication content at the specified version.
        oISHAPIObjs.ISHPubOutObj25.RetrieveVersionMetadata(Context, _
                        GUIDs, _
                        Version, _
                        oCommonFuncs.BuildMinPubMetadata.ToString, _
                        outObjectList)

        Dim VerDoc As New XmlDocument
        VerDoc.LoadXml(outObjectList)
        If VerDoc Is Nothing Or VerDoc.HasChildNodes = False Then
            modErrorHandler.Errors.PrintMessage(3, "Unable to find publication for specified GUID: " + GUID, strModuleName + "-GetBaselineObjects")
            Return Nothing
        End If
        'Get the Baseline ID from the publication:
        Dim baselineID As String = ""
        Dim baselinename As String
        Dim ishfields As XmlNode = VerDoc.SelectSingleNode("//ishfields")
        baselinename = ishfields.SelectSingleNode("ishfield[@name='FISHBASELINE']").InnerText
        'Pull the baseline info
        Dim myBaseline As String = ""
        oISHAPIObjs.ISHBaselineObj.GetBaselineId(Context, baselinename, baselineID)
        oISHAPIObjs.ISHBaselineObj.GetReport(Context, baselineID, Nothing, Languages, Languages, Languages, Resolutions, outObjectList)
        'Load the resulting baseline string as an xml document
        Dim baselineDoc As New XmlDocument
        baselineDoc.LoadXml(outObjectList)
        Dim dictBaselineObjects As New Dictionary(Of String, CMSObject)
        'for each object referenced, store the various info in an object and then store them in the hashtable. GUIDs are the keys.
        For Each baselineObject As XmlNode In baselineDoc.SelectNodes("/baseline/objects/object")
            'create a new CMSObject storage container
            Dim refGuid As String = baselineObject.Attributes.GetNamedItem("ref").Value
            Dim ishtype As String = baselineObject.Attributes.GetNamedItem("type").Value
            If ishtype = "ISHNone" Then
                Continue For
            End If
            Dim refver As String = baselineObject.Attributes.GetNamedItem("versionnumber").Value
            Dim reportitems As String = baselineObject.SelectSingleNode("reportitems").OuterXml.ToString
            Dim CMSObject As New CMSObject(refGuid, refver, ishtype, reportitems)
            'save the object to the hash using the GUID as the key.
            dictBaselineObjects.Add(refGuid, CMSObject)
        Next
        Return dictBaselineObjects
    End Function
#End Region

#Region "Condition Functions"


#End Region

#Region "Document Functions"
    Public Function UpdateTitleProperty(ByVal GUID As String, ByVal Version As String) As Boolean

        'Get object the XML
        Dim ObjectXML As XmlDocument
        ObjectXML = GetObjByID(GUID, Version, "en", "")
        'get the topic xml out of the CDATA from the object
        Dim topicXML As New XmlDocument
        topicXML = oCommonFuncs.GetXMLOut(ObjectXML.SelectSingleNode("//ishdata"))
        Dim ENTitle As String = topicXML.SelectSingleNode("/*/title").InnerText
        ENTitle = ENTitle.Replace(",", "‚")
        ENTitle = ENTitle.Replace("\", "/")
        ENTitle = Replace(ENTitle, "*", "")
        ENTitle = Replace(ENTitle, "?", "")
        ENTitle = Replace(ENTitle, ">", "")
        ENTitle = Replace(ENTitle, "<", "")
        ENTitle = Replace(ENTitle, ":", "")
        ENTitle = Replace(ENTitle, "|", "")
        ENTitle = Replace(ENTitle, "#", "")
        ENTitle = Replace(ENTitle, "!", "")
        'Get the title out of the XML
        Dim strMetaTitle As String = "<ishfields><ishfield name=""FTITLE"" level=""logical"">" + ENTitle + "</ishfield></ishfields>"
        'Set the title on the object
        If SetMeta(strMetaTitle, GUID, Version, "") Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function SetMeta(ByVal strMeta As String, ByVal GUID As String, ByVal Version As String, ByVal strResolution As String, Optional ByVal Language As String = "en") As Boolean
        Try
            ' Clear variable for the result
            oISHAPIObjs.ISHDocObj.SetMetaData(Context, GUID, Version, Language, strResolution, strMeta, "")
        Catch ex As Exception
            modErrorHandler.Errors.PrintMessage(2, "Error setting meta for " + GUID + "" + Version + "" + Language + ". Message: " + ex.Message.ToString, strModuleName + "-SetMeta")
            Return False
        End Try
        Return True
    End Function

    ''' <summary>
    ''' Pulls the specified object from the CMS and saves it at the specified location.  Returns true if successful and file has been saved.
    ''' </summary>
    Public Function GetObjByID(ByVal GUID As String, ByVal Version As String, ByVal Language As String, ByVal Resolution As String, ByVal SavePath As String) As Boolean
        Dim MyNode As XmlNode = Nothing
        Dim MyDoc As New XmlDocument
        Dim MyMeta As New XmlDocument
        Dim XMLString As String = ""
        Dim ISHMeta As String = ""
        Dim ISHResult As String = ""
        Dim filename As String = "BROKEN_FILENAME"
        Dim extension As String = "FIX"
        Dim requestedmeta As StringBuilder = oCommonFuncs.BuildRequestedMetadata()
        'Call the CMS to get our content!
        Try
            ISHResult = oISHAPIObjs.ISHDocObj.GetDocObj(Context, GUID, Version, Language, Resolution, "", "", requestedmeta.ToString, XMLString)
        Catch ex As Exception
            modErrorHandler.Errors.PrintMessage(3, "Failed to retrieve XML from CMS server: " + ex.Message, strModuleName + "-GetObjByID")
            Return False
        End Try

        'Load the XML and get the metadata:
        Try
            MyDoc.LoadXml(XMLString)
            filename = oCommonFuncs.GetFilenameFromIshMeta(MyDoc)
            'Remove any characters not allowed by windows operating system on filenames.
            filename = filename.Replace("\", "")
            filename = filename.Replace("/", "")
            filename = filename.Replace(":", "")
            filename = filename.Replace("*", "")
            filename = filename.Replace("?", "")
            filename = filename.Replace("""", "")
            filename = filename.Replace("<", "")
            filename = filename.Replace(">", "")
            filename = filename.Replace("|", "")
            MyNode = MyDoc.SelectSingleNode("//ishdata")
            'Get the extension:
            For Each ishattrib As XmlAttribute In MyNode.Attributes
                If ishattrib.Name = "fileextension" Then
                    extension = ishattrib.Value
                End If
            Next
        Catch ex As Exception
            modErrorHandler.Errors.PrintMessage(3, "Failed to retrieve object metadata from returned XML: " + ex.Message, strModuleName)
            Return False
        End Try
        'check to see if it's already been exported first, exit if it has...
        If File.Exists(SavePath + "\" + filename + "." + extension.ToLower) Then
            modErrorHandler.Errors.PrintMessage(2, "File already exists. Skipping: " + SavePath + "\" + filename + "." + extension.ToLower, strModuleName)
            Return True
        End If
        'Convert the CDATA to byte array
        Dim finalfile() As Byte
        Try
            'Convert CDATA Blob to Byte array
            finalfile = oCommonFuncs.GetBinaryOut(MyNode)
        Catch ex As Exception
            modErrorHandler.Errors.PrintMessage(3, "Failed to convert CDATA Blob to binary stream - no content returned from CMS: " + ex.Message, strModuleName)
            Return False
        End Try
        'Save the content out to a file:
        Try
            'Create the save path if it doesn't exist:
            If Directory.Exists(SavePath) = False Then
                Directory.CreateDirectory(SavePath)
            End If
            'write to filename, bytes we extracted, don't append

            My.Computer.FileSystem.WriteAllBytes(SavePath + "\" + filename + "." + extension.ToLower, finalfile, False)
            Return True
        Catch ex As Exception
            modErrorHandler.Errors.PrintMessage(3, "Failed to save returned object to a file: " + ex.Message, strModuleName)
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Pulls the specified object from the CMS and saves it at the specified location.  Returns true if successful and file has been saved.
    ''' </summary>
    Public Function GetObjByID(ByVal GUID As String, ByVal Version As String, ByVal Language As String, ByVal Resolution As String) As XmlDocument
        Dim MyNode As XmlNode = Nothing
        Dim MyDoc As New XmlDocument
        Dim MyMeta As New XmlDocument
        Dim XMLString As String = ""
        Dim ISHMeta As String = ""
        Dim ISHResult As String = ""
        Dim filename As String = "BROKEN_FILENAME"
        Dim extension As String = "FIX"
        Dim requestedmeta As StringBuilder = oCommonFuncs.BuildRequestedMetadata()
        'Call the CMS to get our content!
        Try
            ISHResult = oISHAPIObjs.ISHDocObj.GetDocObj(Context, GUID, Version, Language, Resolution, "", "", requestedmeta.ToString, XMLString)
        Catch ex As Exception
            'modErrorHandler.Errors.PrintMessage(3, "Failed to retrieve XML from CMS server: " + ex.Message, strModuleName)
            Return Nothing
        End Try

        'Load the XML and get the metadata:
        Try
            MyDoc.LoadXml(XMLString)
            Return MyDoc
            'filename = GetFilenameFromIshMeta(MyDoc)
            'MyNode = MyDoc.SelectSingleNode("//ishdata")
            ''Get the extension:
            'For Each ishattrib As XmlAttribute In MyNode.Attributes
            '    If ishattrib.Name = "fileextension" Then
            '        extension = ishattrib.Value
            '    End If
            'Next
        Catch ex As Exception
            modErrorHandler.Errors.PrintMessage(3, "Failed to retrieve object metadata from returned XML: " + ex.Message, strModuleName)
            Return Nothing
        End Try
    End Function
    Private Function ObliterateGUID(ByVal GUID As String, ByVal Version As String, ByVal Language As String, Optional ByVal Resolution As String = "High") As Boolean
        Try
            If Resolution.Length > 0 Then
                Try
                    oISHAPIObjs.ISHDocObj.Delete(Context, GUID, Version, Language, "Low")
                    modErrorHandler.Errors.PrintMessage(1, "Deleting low resolution for: " + GUID, strModuleName + "-ObliterateGUID")

                Catch ex As Exception

                End Try
                If Resolution = "High" Then
                    Try
                        oISHAPIObjs.ISHDocObj.Delete(Context, GUID, Version, Language, "High")
                        modErrorHandler.Errors.PrintMessage(1, "Deleting high resolution for: " + GUID, strModuleName + "-ObliterateGUID")

                    Catch ex As Exception

                    End Try
                End If

                Try
                    oISHAPIObjs.ISHDocObj.Delete(Context, GUID, Version, Language, "Thumbnail")
                    modErrorHandler.Errors.PrintMessage(1, "Deleting thumbnail for: " + GUID, strModuleName + "-ObliterateGUID")

                Catch ex As Exception

                End Try

                Try
                    oISHAPIObjs.ISHDocObj.Delete(Context, GUID, Version, Language, "Source")
                    modErrorHandler.Errors.PrintMessage(1, "Deleting source resolution for: " + GUID, strModuleName + "-ObliterateGUID")

                Catch ex As Exception

                End Try

            End If
            Try
                oISHAPIObjs.ISHDocObj.Delete(Context, GUID, Version, "", "")
                modErrorHandler.Errors.PrintMessage(1, "Deleting language level for: " + GUID, strModuleName + "-ObliterateGUID")

            Catch ex As Exception

            End Try
            Try
                oISHAPIObjs.ISHDocObj.Delete(Context, GUID, "", "", "")
                modErrorHandler.Errors.PrintMessage(1, "Deleting version level for: " + GUID, strModuleName + "-ObliterateGUID")

            Catch ex As Exception

            End Try
            'Deleted everything we could!  Let's see if we did it.
            If ObjectExists(GUID, Version, Language, Resolution) Then
                modErrorHandler.Errors.PrintMessage(2, "Deletions performed on GUID failed. GUID: " + GUID + " still exists. Delete manually.", strModuleName + "-ObliterateGUID")
                DeleteFailedGUIDs.Add(GUID)

                Return False
            End If
            DeletedGUIDs.Add(GUID)
            Return True
        Catch ex As Exception
            If ex.Message.Contains("-115") Then
                'The module is referenced by some other module...
                modErrorHandler.Errors.PrintMessage(3, "Unable to delete module: " + GUID + " Referenced by other module(s).", strModuleName)
                Return False
            ElseIf ex.Message.Contains("-102") Then
                oISHAPIObjs.ISHDocObj.Delete(Context, GUID, "", "", "")
            Else
                Return False
            End If

        End Try

    End Function
    ''' <summary>
    ''' Check to see if an object with the specified parameters exists in the CMS.  Returns true if exists.
    ''' </summary>
    Public Function ObjectExists(ByVal GUID As String, ByVal Version As String, ByVal Language As String, Optional ByVal Resolution As String = "") As Boolean
        Dim MyNode As XmlNode = Nothing
        Dim MyDoc As New XmlDocument
        Dim MyMeta As New XmlDocument
        Dim XMLString As String = ""
        Dim ISHMeta As String = ""
        Dim ISHResult As String = ""
        Dim filename As String = "BROKEN_FILENAME"
        Dim extension As String = "FIX"
        Dim requestedmeta As StringBuilder = oCommonFuncs.BuildRequestedMetadata()

        'Call the CMS to get our content!
        Try
            ISHResult = oISHAPIObjs.ISHDocObj.GetDocObj(Context, GUID, Version, Language, Resolution, "", "", requestedmeta.ToString, XMLString)
            Return True
        Catch ex As Exception
            'modErrorHandler.Errors.PrintMessage(3, "Failed to retrieve XML from CMS server: " + ex.Message, strModuleName)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Replaces a specified module with templated content.  
    ''' Most commonly used before attempting to recursively delete referencing modules to prevent circular references.
    ''' </summary>
    Public Function ReplaceWithTemplatedContent(ByVal GUID As String, ByVal Version As String, ByVal Language As String) As Boolean
        'Set the GUIDs in our templates.
        oCommonFuncs.SetGUIDinTemplates(GUID)

        'Start the replacement process:
        Dim requestedmetadata As StringBuilder = oCommonFuncs.BuildRequestedMetadata()
        Dim RequestedXMLObject As String = ""
        Dim doc As New XmlDocument
        Dim IshType As String
        Dim TopicType As String
        Dim Data() As Byte

        Try
            'check out the module (must be map or topic)
            oISHAPIObjs.ISHDocObj.CheckOut(Context, GUID, Version, Language, "", requestedmetadata.ToString, RequestedXMLObject)
        Catch ex As Exception
            'If, for some reason, we already have an object checked out, great.  otherwise, we can't check it out for some reason.
            'Exit Code for already checking an object out is -132
            If ex.Message.Contains("-132") Then
                'we have it checked out already, but we still need to get the object CData:
                oISHAPIObjs.ISHDocObj.GetDocObj(Context, GUID, Version, Language, "", "", "", requestedmetadata.ToString, RequestedXMLObject)
            Else
                modErrorHandler.Errors.PrintMessage(3, "Unable to checkout GUID: " + GUID + " Error: " + ex.Message, strModuleName + "-ReplaceWithTemplatedContent")
                Return False
            End If
        End Try

        'Load the XML and get the metadata:
        doc.LoadXml(RequestedXMLObject)
        'get the ISHType from the meta
        IshType = oCommonFuncs.GetISHTypeFromMeta(doc)
        Select Case IshType
            Case "ISHMasterDoc"
                'if a map, replace the content with our template content
                Data = oCommonFuncs.XMLTemplates.mapblob

            Case "ISHModule"
                'if a topic, find out what kind
                TopicType = oCommonFuncs.GetTopicTypeFromMeta(doc)
                Select Case TopicType
                    Case "task"
                        Data = oCommonFuncs.XMLTemplates.taskblob
                    Case "concept"
                        Data = oCommonFuncs.XMLTemplates.conceptblob
                    Case "reference"
                        Data = oCommonFuncs.XMLTemplates.referenceblob
                    Case "troubleshooting"
                        Data = oCommonFuncs.XMLTemplates.troubleshootingblob
                    Case Else
                        modErrorHandler.Errors.PrintMessage(2, "Unexpected TopicType used for GUID: " + GUID, strModuleName + "-ReplaceWithTemplatedContent")
                        Return False
                End Select
                'replace the content with our template content
            Case Else
                'Returned an unexpected type...
                'modErrorHandler.Errors.PrintMessage(2, "Unable to determine ISHType for GUID: " + GUID, strModuleName + "-ReplaceWithTemplatedContent")
                Return False
        End Select
        If IsNothing(Data) = False Then
            Try
                oISHAPIObjs.ISHDocObj.CheckIn(Context, GUID, Version, Language, "", "", "EDTXML", Data)
                Return True
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(3, "Unable to checkin after replacing content. GUID: " + GUID + " Object is still checked out to user!", strModuleName + "-ReplaceWithTemplatedContent")
                Return False
            End Try
        Else
            modErrorHandler.Errors.PrintMessage(3, "No Data blob was created from the template file.  Can't replace content in CMS with nothing.  Failed to replace GUID:" + GUID, strModuleName + "-ReplaceWithTemplatedContent")
            Return False
        End If
    End Function

    ''' <summary>
    ''' Deletes a given object and referencing parents (if possible).  Optionally allows deleting all referenced children (does not apply to parents).
    ''' </summary>
    Public Function DeleteObjectRecursivelyByGUID(ByVal GUID As String, Optional ByVal Version As String = "1", Optional ByVal Language As String = "en", Optional ByVal Resolution As String = "", Optional ByVal DeleteSubs As Boolean = True) As Boolean
        Dim ParentModules As New Hashtable
        Dim ChildrenModules As New Hashtable
        Dim requestedmetadata As StringBuilder = oCommonFuncs.BuildRequestedMetadata()
        'first, find out if the GUID exists in the CMS:
        If ObjectExists(GUID, Version, Language, Resolution) Then

            'We're going to delete it anyway.  Use template to replace the contents completely (if not an image).
            If Resolution = "" Then
                If ReplaceWithTemplatedContent(GUID, Version, Language) = False Then
                    'modErrorHandler.Errors.PrintMessage(2, "Unable to replace content in GUID: " + GUID + " with default, template content. May not be able to delete due if the topic contains circular references to referencing modules.", strModuleName + "-RecursiveDeletion")
                End If
            End If


            'if the guid has owners, use the list to recurse into them
            If GetReferencingModules(GUID, ParentModules, Version, Language, requestedmetadata.ToString) Then
                For Each parentmodule As DictionaryEntry In ParentModules
                    ' Make sure we don't recurse into ourselves here...
                    If Not parentmodule.Value.GUID = GUID Then
                        Dim result As Boolean = DeleteObjectRecursivelyByGUID(parentmodule.Value.GUID, parentmodule.Value.Version, parentmodule.Value.Language, parentmodule.Value.Resolution, False)
                        If result = False Then
                            'There was a problem deleting the parent!
                            modErrorHandler.Errors.PrintMessage(3, "Unable to delete parent of GUID: " + GUID + " Parent that returned an error was GUID: " + parentmodule.Value.GUID, strModuleName + "-RecursiveDeletion")
                            Return False
                        End If
                    End If
                Next
            Else
                'otherwise, if it doesn't, get the children
                If GetReferencedModules(GUID, ChildrenModules, Version, Language, requestedmetadata.ToString) Then 'if true, has children
                    'then, delete the current GUID (if it can be deleted), 
                    If CanBeDeleted(GUID, Version, Language, Resolution) Then
                        ObliterateGUID(GUID, Version, Language, Resolution)
                    Else
                        'Can't be deleted for some reason...
                        modErrorHandler.Errors.PrintMessage(3, "Unable to delete GUID: " + GUID, strModuleName + "RecursiveDeletion")
                        Return False
                    End If
                    'check to see if we should also delete children of the current module
                    If DeleteSubs = True Then
                        'recursedelete into each of the children
                        For Each childmodule As DictionaryEntry In ChildrenModules
                            'Make sure we don't recurse on ourselves
                            If Not childmodule.Value.GUID = GUID Then
                                If DeleteObjectRecursivelyByGUID(childmodule.Value.GUID, childmodule.Value.Version, childmodule.Value.language, childmodule.Value.resolution, DeleteSubs) = False Then
                                    modErrorHandler.Errors.PrintMessage(3, "Failed to delete descendant of: " + GUID, strModuleName + "RecursiveDeletion")
                                    Return False
                                End If
                            End If
                        Next
                    End If
                    Return True
                Else
                    'no children?  Just delete the current GUID
                    If CanBeDeleted(GUID, Version, Language) Then
                        ObliterateGUID(GUID, Version, Language, Resolution)
                        Return True
                    Else
                        'Can't be deleted for some reason...
                        modErrorHandler.Errors.PrintMessage(3, "Unable to delete GUID: " + GUID, strModuleName + "RecursiveDeletion")
                        Return False
                    End If
                End If
            End If
            'Object existed but had parents before and needed children trimmed.  That's done.
            'Shouldn't have any parents at this point, so we can likely delete it.  let's try:
            If CanBeDeleted(GUID, Version, Language) Then
                ObliterateGUID(GUID, Version, Language, Resolution)
                Return True
            Else
                'Can't be deleted for some reason...
                modErrorHandler.Errors.PrintMessage(3, "Unable to delete GUID: " + GUID, strModuleName + "RecursiveDeletion")
                Return False
            End If
            'made it this far - everything deleted successfully without returning False.
            Return True
        Else 'End "if GUID exists"
            'GUID doesn't exist.  Mission accomplished - it is already deleted!
            modErrorHandler.Errors.PrintMessage(1, "GUID doesn't exist in the CMS - no need to delete.  GUID: " + GUID, strModuleName + "-DeleteObjectRecursivelyByGUID")
            Return True
        End If

    End Function
    ''' <summary>
    ''' Checks to see if a specified module has no referencing modules and is not in a released state. Returns true if both conditions are met.
    ''' </summary>
    ''' <param name="GUID">GUID of object in CMS</param>
    ''' <param name="Version">Version of object in CMS</param>
    ''' <param name="Language">Language of object in CMS</param>
    ''' <param name="Resolution">Resolution of object in CMS</param>
    ''' <returns>Boolean</returns>
    Public Function CanBeDeleted(ByVal GUID As String, Optional ByVal Version As String = "1", Optional ByVal Language As String = "en", Optional ByVal Resolution As String = "") As Boolean
        Dim ModuleHash As New Hashtable

        If GetReferencingModules(GUID, ModuleHash, Version, Language) = True Then
            'Has parents, can't be deleted
            Return False
        Else
            'This object has no parents.  first check passed.  Continue

        End If
        Dim SearchResult As String = ""
        oISHAPIObjs.ISHDocObj.GetMetaData(Context, GUID, Version, Language, Resolution, "<ishfields><ishfield name=""FSTATUS"" level=""lng""/></ishfields>", SearchResult)
        If SearchResult.Contains("""FSTATUS"" level=""lng"">Released") Then
            'Status is released.  Can't delete
            Return False
        Else
            'Status is something else, can delete
            Return True
        End If




    End Function
#End Region

#Region "Folder Functions"
    Private Function FolderContainsGUID(ByVal GUID As String, ByVal FolderID As Long) As Boolean
        If Not FolderID = 0 Then
            Dim requestedobjects As String = ""


            Dim RealRequestedMeta As New StringBuilder
            RealRequestedMeta = oCommonFuncs.BuildRequestedMetadata()



            'Returns a list of referencing objects to "requestedobjects" string as xml:
            Try
                oISHAPIObjs.ISHFolderObj.GetContents(Context, FolderID, "", RealRequestedMeta.ToString, requestedobjects)
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(2, "Unable to get contents of folder " + FolderID, strModuleName + "-FolderContainsGUID")
            End Try


            'Load the string as an xmldoc
            Dim doc As New XmlDocument
            doc.LoadXml(requestedobjects)
            'get our returned objects into a hashtable (key=GUID, Value=Information + XMLNode of returned data)
            Dim SubObjects As Hashtable = GetReportedObjects(doc)
            'Key is a mashup of the GUID, version, and other info so it's not going to match against our GUID... We'll need to spin through the hash to find what we want.
            For Each myentry As DictionaryEntry In SubObjects
                If myentry.Value.GUID = GUID Then
                    Return True
                End If
            Next
            'No matches found, exit.
            Return False
        Else
            'can't do searches for content in root folder.
            Return False
        End If


    End Function

    Public Function FindFolderIDforObjbyGUID(ByVal GUID As String, ByVal ParentFolderID As Long) As Long
        'Recursively searches for a GUID and returns the ID when found.  returns -1 if not found.

        'First, check to see if our GUID exists in this folder (only if not root folder):
        If FolderContainsGUID(GUID, ParentFolderID) Then
            'if we found it, return this as the valid parent!
            Return ParentFolderID
            Exit Function
        Else

            'if not, we need to dive deeper.
            'get a folderlist of all children of currentfolderid
            Dim subfolderlistXML As String = ""
            Dim CMSReply As String
            Try
                CMSReply = oISHAPIObjs.ISHFolderObj.GetSubFolders(Context, ParentFolderID.ToString, 1, subfolderlistXML)
            Catch ex As Exception

            End Try
            'Load the subfolderlistXML into an xml document
            ' Create the reader from the string.
            Dim strReader As New StringReader(subfolderlistXML)

            Dim reader As XmlReader = XmlReader.Create(strReader)
            ' Create the new XMLdoc and load the content into it.
            Dim doc As New XmlDocument()
            doc.PreserveWhitespace = True
            doc.Load(reader)

            'recurse into each subid
            Dim ReturnedFolderID As Long
            Try
                For Each subid As XmlNode In doc.SelectNodes("//ishfolders/ishfolder/@ishfolderref")
                    ReturnedFolderID = FindFolderIDforObjbyGUID(GUID, Convert.ToInt64(subid.Value.ToString))
                    If ReturnedFolderID > 0 Then
                        Exit For
                    End If
                Next
                Return ReturnedFolderID
            Catch ex As Exception
                ' No Sub Folders.  
                Return -1
            End Try
        End If
    End Function
#End Region

#Region "Meta Functions"

#End Region

#Region "Output Functions"

#End Region

#Region "Pub Functions"
    Public Function GetLatestPubVersionNumber(ByVal GUID As String) As String
        Dim outObjectList As String = ""
        Dim GUIDs(0) As String
        GUIDs(0) = GUID

        'Get the existing content at the latest version.
        oISHAPIObjs.ISHPubOutObj25.RetrieveVersionMetadata(Context, _
                        GUIDs, _
                        "latest", _
                        oCommonFuncs.BuildMinPubMetadata.ToString, _
                        outObjectList)

        Dim VerDoc As New XmlDocument
        VerDoc.LoadXml(outObjectList)
        If VerDoc Is Nothing Or VerDoc.HasChildNodes = False Then
            modErrorHandler.Errors.PrintMessage(3, "Unable to find publication for specified GUID: " + GUID, strModuleName + "-GetLatestPubVersionNumber")
            Return Nothing
        End If
        'Iterate through each returned obj and figure out the highest number value for 'Version returned'
        Dim ishfields As XmlNode = VerDoc.SelectSingleNode("//ishfields")
        Dim VersionNode As XmlNode = ishfields.SelectSingleNode("ishfield[@name='VERSION']")
        Return VersionNode.InnerText
    End Function

    Public Function GetMasterMapGUID(ByVal PubGUID As String, ByVal PubVer As String) As String
        'Get the pub's mastermap GUID from the pub object
        Dim outObjectList As String = ""
        Dim GUIDs(0) As String
        GUIDs(0) = PubGUID
        Dim Languages(0) As String
        Languages(0) = "en"
        Dim Resolutions() As String
        oISHAPIObjs.ISHMetaObj.GetLOVValues(Context, "DRESOLUTION", Resolutions)

        'Get the existing publication content at the specified version.
        oISHAPIObjs.ISHPubOutObj25.RetrieveVersionMetadata(Context, _
                        GUIDs, _
                        PubVer, _
                        oCommonFuncs.BuildMinPubMetadata.ToString, _
                        outObjectList)

        Dim VerDoc As New XmlDocument
        VerDoc.LoadXml(outObjectList)
        If VerDoc Is Nothing Or VerDoc.HasChildNodes = False Then
            modErrorHandler.Errors.PrintMessage(3, "Unable to find publication for specified GUID: " + PubGUID, strModuleName + "-GetMasterMapGUID")
            Return Nothing
        End If
        'Get the Master Map GUID from the publication:
        Dim ishfields As XmlNode = VerDoc.SelectSingleNode("//ishfields")
        Return ishfields.SelectSingleNode("ishfield[@name='FISHMASTERREF']").InnerText
    End Function
#End Region

#Region "PubOutput Functions"

#End Region

#Region "Reports Functions"
    

    ''' <summary>
    ''' Given a commonly returned "ishobjects" XML Document returned from most CMS queries, 
    ''' this function converts each entry to Dictionary Entries in a hashtable.  Each entry 
    ''' uses a combined GUID+meta key that appears the same as the commonly used CMS filenames. 
    ''' The value is the ObjectData structure found in this class.
    ''' </summary>
    ''' <param name="XMLDoc">A standard "ishobjects" XML Document returned from a query to the CMS.</param>
    ''' <returns>A Hashtable containing all of the ishobjects and their metadata.</returns>
    Public Function GetReportedObjects(ByVal XMLDoc As XmlDocument) As Hashtable
        'Grab each returned ishobject and build our structure out of it.  Add it to the hash we'll be returning.
        Dim ReturnHash As New Hashtable
        For Each CMSModule As XmlNode In XMLDoc.SelectNodes("/ishobjects/ishobject")
            Dim returnobj As New ObjectData
            With returnobj
                .GUID = CMSModule.Attributes.GetNamedItem("ishref").Value.ToString
                .IshType = CMSModule.Attributes.GetNamedItem("ishtype").Value.ToString
                If .IshType = "ISHNotFound" Then
                    'Ack, it's a broken link ishobject! skip it here.  
                    'TODO: would need to rewrite this if we want to track the broken links for anything.
                    Continue For
                End If
                Dim ishfields As XmlNode = CMSModule.SelectSingleNode("ishfields")

                .Title = ishfields.SelectSingleNode("ishfield[contains(@name, 'FTITLE')]").InnerText
                .Version = ishfields.SelectSingleNode("ishfield[contains(@name, 'VERSION')]").InnerText
                .Language = ishfields.SelectSingleNode("ishfield[contains(@name, 'DOC-LANGUAGE')]").InnerText
                .Status = ishfields.SelectSingleNode("ishfield[contains(@name, 'FSTATUS')]").InnerText
                If .IshType = "ISHIllustration" Then
                    .Resolution = ishfields.SelectSingleNode("ishfield[contains(@name, 'FRESOLUTION')]").InnerText
                Else
                    .Resolution = ""
                End If
                .MetaData = CMSModule
            End With
            If ReturnHash.Contains(returnobj.Title + "=" + returnobj.GUID + "=" + returnobj.Version + "=" + returnobj.Language + "=" + returnobj.Resolution) = False Then
                ReturnHash.Add(returnobj.Title + "=" + returnobj.GUID + "=" + returnobj.Version + "=" + returnobj.Language + "=" + returnobj.Resolution, returnobj)
            Else
                'MsgBox("Hit a Duplicate entry! Impossible")
            End If
        Next
        Return ReturnHash
    End Function

    ''' <summary>
    ''' Finds all objects that are referred to by the specified ISHObject as children.
    ''' </summary>
    ''' <param name="GUID"></param>
    ''' <param name="ReferencedModules">Returns a hashtable of all referenced modules.</param>
    ''' <param name="Version"></param>
    ''' <param name="Language"></param>
    ''' <param name="RequestedMetadata">You can limit the metadata to retrieve from the CMS using the common templating structure for CMS metadata requests.  Expected format is string of XML content.</param>        
    Public Function GetReferencedModules(ByVal GUID As String, ByRef ReferencedModules As Hashtable, Optional ByVal Version As String = "1", Optional ByVal Language As String = "en", Optional ByVal RequestedMetadata As String = "") As Boolean
        Dim requestedobjects As String = ""
        Dim RealRequestedMeta As New StringBuilder
        If RequestedMetadata = "" Then
            RealRequestedMeta = oCommonFuncs.BuildRequestedMetadata()
        Else
            RealRequestedMeta.Append(RequestedMetadata)
        End If


        'Returns a list of referencing objects to "requestedobjects" string as xml:
        oISHAPIObjs.ISHReportsObj.GetReferencedDocObj(Context, GUID, Version, Language, False, RealRequestedMeta.ToString, requestedobjects)
        'Load the string as an xmldoc
        Dim doc As New XmlDocument
        doc.LoadXml(requestedobjects)
        'get our returned objects into a hashtable (key=GUID, Value=Information + XMLNode of returned data)
        ReferencedModules = GetReportedObjects(doc)
        Dim TrimmedReferencedModules As New Hashtable(ReferencedModules)
        'drop any objects that match our current object's GUID (could be more than one entry, depending on whether or not it's an image)
        For Each obtainedmodule As DictionaryEntry In ReferencedModules
            'if we find an entry that has the same GUID as our currently examined module's GUID, let's drop it.
            If obtainedmodule.Value.guid = GUID Then
                TrimmedReferencedModules.Remove(obtainedmodule.Key)
            End If
        Next
        ReferencedModules = TrimmedReferencedModules
        'we've removed all the reflective references.  Should be greater than 0 if we have any matches.
        If ReferencedModules.Count > 0 Then

            Return True
        Else
            'only the requested GUID was found so there aren't any referenced children...
            Return False
        End If
    End Function

    ''' <summary>
    ''' Finds all objects that refer to the specified ISHObject as parents.
    ''' </summary>
    ''' <param name="GUID"></param>
    ''' <param name="ReferencingModules">Returns a hashtable of all referencing objects.</param>
    ''' <param name="Version"></param>
    ''' <param name="Language"></param>
    ''' <param name="RequestedMetadata">You can limit the metadata to retrieve from the CMS using the common templating structure for CMS metadata requests.  Expected format is string of XML content.</param>        
    Public Function GetReferencingModules(ByVal GUID As String, ByRef ReferencingModules As Hashtable, Optional ByVal Version As String = "1", Optional ByVal Language As String = "en", Optional ByVal RequestedMetadata As String = "") As Boolean
        Dim requestedobjects As String = ""
        Dim RealRequestedMeta As New StringBuilder
        If RequestedMetadata = "" Then
            RealRequestedMeta = oCommonFuncs.BuildRequestedMetadata()
        Else
            RealRequestedMeta.Append(RequestedMetadata)
        End If

        'Returns a list of referencing objects to "requestedobjects" string as xml:
        oISHAPIObjs.ISHReportsObj.GetReferencedByDocObj(Context, GUID, Version, Language, False, RealRequestedMeta.ToString, requestedobjects)
        'Load the string as an xmldoc
        Dim doc As New XmlDocument
        doc.LoadXml(requestedobjects)
        'get our returned objects into a hashtable (key=GUID, Value=Information + XMLNode of returned data)
        ReferencingModules = GetReportedObjects(doc)
        Dim TrimmedReferencingModules As New Hashtable(ReferencingModules)
        'drop any objects that match our current object's GUID (could be more than one entry, depending on whether or not it's an image)
        For Each obtainedmodule As DictionaryEntry In ReferencingModules
            'if we find an entry that has the same GUID as our currently examined module's GUID, let's drop it.
            If obtainedmodule.Value.guid = GUID Then
                TrimmedReferencingModules.Remove(obtainedmodule.Key)
            End If
        Next
        ReferencingModules = TrimmedReferencingModules
        'We've removed all reflective references from the list.
        If ReferencingModules.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region

#Region "Search Functions"

#End Region

#Region "Workflow Functions"

#End Region

#End Region




End Class
