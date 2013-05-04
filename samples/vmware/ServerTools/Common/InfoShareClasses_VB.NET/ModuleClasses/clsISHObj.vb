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


<ComClass(clsISHObj.ClassId, clsISHObj.InterfaceId, clsISHObj.EventsId)> _
Public Class clsISHObj

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "990ade3a-03fb-49d3-86ee-7b7c21f958d0"
    Public Const InterfaceId As String = "759d75cb-111a-4f9f-9508-887bdc9dafc0"
    Public Const EventsId As String = "36982ff3-d385-4e70-81a3-b7dbeab2acf5"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        'Load up our XML Templates for each type:
        oDocument.LoadXMLTemplates()

    End Sub






    Public oApplication As New IshApplication
    Public oDocument As New IshDocument
    Public oBaseline As New IshBaseline
    Public oConditions As New IshConditions
    Public oMeta As New IshMeta
    Public oOutput As New IshOutput
    Public oPub As New IshPub
    Public oPubOutput As New IshPubOutput
    Public oReports As New IshReports
    Public oSearch As New IshSearch
    Public oWorkflow As New IshWorkflow
    Public oFolder As New IshFolder
    Public CMSServerURL As New String("")
    Private Shared m_Context As New String("")
    Private Shared m_recursion_hash As New Hashtable

    Public ReadOnly Property DeletedGUIDs() As ArrayList
        Get
            DeletedGUIDs = oDocument.DeletedGUIDs
        End Get
    End Property
    Public ReadOnly Property DeleteFailedGUIDs() As ArrayList
        Get
            DeleteFailedGUIDs = oDocument.DeleteFailedGUIDs
        End Get
    End Property


    Private Shared ReadOnly Property ISHApp() As String
        Get
            ISHApp = "InfoShareAuthor"
        End Get
    End Property
    Private Shared ReadOnly Property strModuleName() As String
        Get
            strModuleName = "clsISHObj"
        End Get
    End Property
    ''' <summary>
    ''' Context of the current user's credentials on the specified CMS URL.
    ''' The easiest way of setting this value is to run SetContext after instantiating the object.
    ''' However, it can also be set directly.
    ''' </summary>
    Public Property Context() As String
        Get
            Context = m_Context
        End Get
        Set(ByVal value As String)
            m_Context = value
        End Set
    End Property

    ''' <summary>
    ''' Used to set the context of the current user's credentials on the specified CMS URL.
    ''' </summary>
    Public Function SetContext(ByVal uname As String, ByVal passwd As String, ByVal RepositoryURL As String) As Boolean
        'First, test the App web reference with the passed URL.
        oApplication.ISHAppObj.Url = RepositoryURL + "/InfoShareWS/Application20.asmx"
        'Set the IshAppObj to use the passed URL.
        m_Context = ""
        CMSServerURL = "" 'Reset because we're trying a new connection.  if it fails, we don't want to keep this around.
        'test logging in (m_context gets set here):
        Try
            oApplication.ISHAppObj.Login(ISHApp, uname, passwd, m_Context)
            CMSServerURL = RepositoryURL
        Catch ex As Exception
            modErrorHandler.Errors.PrintMessage(3, "Login failed: " + ex.Message.ToString, strModuleName)
            Return False
        End Try

        If Not m_Context = "" Then
            'Set all the ISHObj web refs to use this passed URL...
            Try
                oApplication.ISHAppObj.Url = RepositoryURL + "/InfoShareWS/Application20.asmx"
                oDocument.ISHDocObj.Url = RepositoryURL + "/InfoShareWS/DocumentObj20.asmx"
                oDocument.ISHDocObj25.Url = RepositoryURL + "/InfoShareWS/DocumentObj25.asmx"
                oBaseline.ISHBaselineObj.Url = RepositoryURL + "/InfoShareWS/Baseline20.asmx"
                oConditions.ISHCondObj.Url = RepositoryURL + "/InfoShareWS/Condition20.asmx"
                oFolder.ISHFolderObj.Url = RepositoryURL + "/InfoShareWS/Folder20.asmx"
                oMeta.ISHMetaObj.Url = RepositoryURL + "/InfoShareWS/MetaDataAssist20.asmx"
                oOutput.ISHOutputObj.Url = RepositoryURL + "/InfoShareWS/OutputFormat20.asmx"
                oPub.ISHPubObj.Url = RepositoryURL + "/InfoShareWS/Publication20.asmx"
                oPubOutput.ISHPubOutObj.Url = RepositoryURL + "/InfoShareWS/PublicationOutput20.asmx"
                oReports.ISHReportsObj.Url = RepositoryURL + "/InfoShareWS/Reports20.asmx"
                oSearch.ISHSearchObj.Url = RepositoryURL + "/InfoShareWS/Search20.asmx"
                oWorkflow.ISHWorkflowObj.Url = RepositoryURL + "/InfoShareWS/Workflow20.asmx"
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(1, "Unable to set Web Reference URL.  Message: " + ex.Message, strModuleName)
                Return False
            End Try

        Else
            'ErrorMessage Out.
            modErrorHandler.Errors.PrintMessage(1, "Login Failed: One or more parameters are incorrect", strModuleName)
            Return False
        End If
        CopyDTDFile()
        Return True
    End Function



    Public Class IshApplication
        Public Shared ISHAppObj As New Application20.Application20
    End Class
    Public Class IshDocument
        Public Shared DeletedGUIDs As New ArrayList
        Public Shared DeleteFailedGUIDs As New ArrayList
        Public Shared ISHDocObj As New DocumentObj20.DocumentObj20
        Public Shared ISHDocObj25 As New DocumentObj25.DocumentObj25
        Public Shared Function MoveObject(ByVal GUID As String, ByVal ToFolderID As Long) As Boolean
            'TODO: Has to search through the entire CMS Structure just to figure out the current folder ID of the specified GUID...  Terribly inefficient!
            'The problem is that FolderID is not tracked with objects as part of its metadata.  Looking up an object gives you NO information about where it exists in the CMS.
            Try
                Dim CurrentFolder As Long = IshFolder.FindFolderIDforObjbyGUID(GUID, 0)
                If CurrentFolder > 0 Then
                    ISHDocObj.Move(m_Context, GUID, CurrentFolder.ToString, ToFolderID.ToString)
                Else
                    modErrorHandler.Errors.PrintMessage(3, "Unable to move object " + GUID + " to specified location!", strModuleName + "-MoveObject")
                End If
                Return True
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(3, "Unable to move object " + GUID + " to specified location!", strModuleName + "-MoveObject")
                Return False
            End Try
        End Function
        
        Public Shared Function ChangeState(ByVal strDesiredState As String, ByVal strGUID As String, ByVal strVer As String, ByVal strLanguage As String, ByVal strResolution As String) As Boolean
            Dim myCurrentState = GetCurrentState(strGUID, strVer, strResolution, strLanguage)
            Dim processingresult As Boolean = True
            Dim strMetaState As String = "<ishfields><ishfield name=""FSTATUS"" level=""lng"">" + strDesiredState + "</ishfield></ishfields>"
            ''Could be used to update person assigned to a specific role as well...
            'Dim strMetaRole As String = "<ishfields><ishfield name=""FEDITOR"" level=""lng"">" + strEditorName + "</ishfield></ishfields>"
            Dim result As Boolean = True

            ' Generic move status drive used to change the status of a topic and COULD be used to update the name associated with the status
            If (m_Context = "") Then
                modErrorHandler.Errors.PrintMessage(3, "Context is not set. Unable to continue.", strModuleName + "-ChangeState")
                Return False
            End If

            If (CanMoveToState(strDesiredState, strGUID, strVer, strResolution, strLanguage)) Then
                ' Change the state
                If (SetMeta(strMetaState, strGUID, strVer, strResolution, strLanguage)) Then
                    modErrorHandler.Errors.PrintMessage(1, "Changed state for " + strGUID + "=" + strVer + "=" + strResolution + "=" + strLanguage + ".", strModuleName + "-ChangeState")
                    '' Check to see if we'r esupposed to assign the editor name as well...

                    'If (GetCurrentState(strGUID, strVer, strResolution, strLanguage) = "Editing") And strEditorName.Length > 0 Then
                    '    ' Now set the metadata on the role to the current user's name
                    '    ' For example, set "lgalindo" as FEDITOR
                    '    If SetMeta(strMetaRole, strGUID, strVer, strResolution, strLanguage) Then
                    '        modErrorHandler.Errors.PrintMessage(1, "Changed editor for " + strGUID + "=" + strVer + "=" + strResolution + "=" + strLanguage + " to " + strEditorName + ".", strModuleName + "-ChangeState")
                    '    Else
                    '        'Something happened when changing the Editor!
                    '        modErrorHandler.Errors.PrintMessage(2, "Failed to change editor for " + strGUID + "=" + strVer + "=" + strResolution + "=" + strLanguage + " to " + strEditorName + ".", strModuleName + "-ChangeState")
                    '        processingresult = False
                    '    End If

                    'End If



                Else
                    'Something happened when changing the state!
                    modErrorHandler.Errors.PrintMessage(3, "Failed to change state for " + strGUID + "=" + strVer + "=" + strResolution + "=" + strLanguage + ". Current state is: " + myCurrentState, strModuleName + "-ChangeState")
                    processingresult = False
                End If
            Else
                '' If the current state is editing, just make sure that the editor name is updated
                'If (GetCurrentState(strGUID, strVer, strResolution) = "Editing") And strEditorName.Length > 0 Then
                '    ' Now set the metadata on the role to the current user's name
                '    ' For example, set "lgalindo" as FEDITOR
                '    If SetMeta(strMetaRole, strGUID, strVer, strResolution) Then
                '        modErrorHandler.Errors.PrintMessage(1, "Changed editor for " + strGUID + "=" + strVer + "=" + strResolution + " to " + strEditorName + ".", strModuleName + "-ChangeState")
                '    Else
                '        'Something happened when changing the Editor!
                '        modErrorHandler.Errors.PrintMessage(2, "Failed to change editor for " + strGUID + "=" + strVer + "=" + strResolution + " to " + strEditorName + ".", strModuleName + "-ChangeState")
                '        processingresult = False
                '    End If
                'End If
                If myCurrentState = strDesiredState Then
                    modErrorHandler.Errors.PrintMessage(1, "State already set as requested for " + strGUID + "=" + strVer + "=" + strResolution + "=" + strLanguage + ".", strModuleName + "-ChangeState")
                    processingresult = True
                Else
                    modErrorHandler.Errors.PrintMessage(3, "Failed to change state for " + strGUID + "=" + strVer + "=" + strResolution + "=" + strLanguage + ". Current state is: " + myCurrentState, strModuleName + "-ChangeState")
                    processingresult = False
                End If
            End If



            Return processingresult
        End Function
        
        Public Shared Function CheckIn(ByVal PathToCheckInFile As String) As Boolean
            Dim checkinfile As New FileInfo(PathToCheckInFile)
            Dim CMSFilename, GUID, Version, Language, Resolution As New String("")
            GetCommonMetaFromLocalFile(checkinfile.FullName, CMSFilename, GUID, Version, Language, Resolution)
            Dim checkinblob As Byte() = GetIshBlobFromFile(checkinfile.FullName)
            Try
                ISHDocObj.CheckIn(m_Context, GUID, Version, Language, Resolution, "", GetISHEdt(checkinfile.Extension), checkinblob)
                modErrorHandler.Errors.PrintMessage(1, "Checked in object " + checkinfile.FullName + ".", strModuleName + "-CheckIn")
                checkinfile.Attributes = FileAttributes.Normal
                checkinfile.Delete()
                Return True
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(3, "Failed to check in object " + checkinfile.FullName + ". Message: " + ex.Message, strModuleName + "-CheckIn")
                Return False
            End Try
        End Function
        Public Shared Function CheckOut(ByVal GUID As String, ByVal Version As String, ByVal Language As String, ByVal Resolution As String, ByVal LocalStorePath As String) As Boolean
            Dim CheckOutFile As New String("")
            'first, ensure it exists.
            If ObjectExists(GUID, Version, Language, Resolution) Then
                'Check out the object
                Try
                    ISHDocObj.CheckOut(m_Context, GUID, Version, Language, Resolution, "", CheckOutFile)
                    modErrorHandler.Errors.PrintMessage(1, "Checked out object " + GUID + "=" + Version + "=" + Language + "=" + Resolution + ".", strModuleName + "-CheckOut")
                    CheckOutFile = ""
                Catch ex As Exception
                    If GetCurrentState(GUID, Version, Resolution, Language) = "Released" Then
                        CreateNewVersion(GUID, Version, Language, Resolution)
                        'now try checking out the new version:
                        Try
                            ISHDocObj.CheckOut(m_Context, GUID, Version, Language, Resolution, "", CheckOutFile)
                        Catch ex3 As Exception
                            modErrorHandler.Errors.PrintMessage(3, "Failed to check out object " + GUID + "=" + Version + "=" + Language + "=" + Resolution + ". Message: " + ex3.Message, strModuleName + "-CheckOut")
                            Return False
                        End Try
                    End If
                End Try
                'Now try fetching the object to our local system. Don't do this for objects that have a "High" resolution if the current one being processed is Low.
                Try
                    If ObjectExists(GUID, Version, Language, "High") And Resolution = "Low" Then
                        'Skip it.
                    Else
                        'Grab it down locally.
                        GetObjByID(GUID, Version, Language, Resolution, LocalStorePath)
                        modErrorHandler.Errors.PrintMessage(1, "Downloaded object " + GUID + "=" + Version + "=" + Language + "=" + Resolution + ".", strModuleName + "-CheckOut")
                    End If
                Catch ex As Exception
                    modErrorHandler.Errors.PrintMessage(3, "Failed to download checked-out object " + GUID + "=" + Version + "=" + Language + "=" + Resolution + ". Message: " + ex.Message, strModuleName + "-CheckOut")
                End Try
                Return True
            Else
                Return False
            End If

        End Function
        ''' <summary>
        ''' Given the current informaiton of a given CMS Object, creates a new version on the same branch containing the 
        ''' content from the previous version in a Draft state.
        ''' </summary>
        ''' <param name="GUID">Unique object ID</param>
        ''' <param name="Version">Passes in the current version of the object to be versioned, returns the new version number</param>
        ''' <param name="Language">Language</param>
        ''' <param name="Resolution">(Optional) Resolution</param>
        ''' <returns>True if successful, False if failed</returns>
        ''' <remarks></remarks>
        Public Shared Function CreateNewVersion(ByVal GUID As String, ByRef Version As String, ByVal Language As String, Optional ByVal Resolution As String = "") As Boolean
            'Specified version is released so we need to create a new version of the specified branch.
            'Get the existing content at the current version.
            Dim newverDoc As XmlDocument = GetObjByID(GUID, Version, Language, Resolution)
            'if the content is a map or topic, we need to remove the processing instruction before it can be added as a new version.
            Dim datablob As Byte()
            Dim newverIshType As String = GetISHTypeFromMeta(newverDoc)
            If newverIshType = "ISHModule" Or newverIshType = "ISHMasterDoc" Then
                Dim MyNode As XmlNode = newverDoc.SelectSingleNode("//ishdata")
                'get the dita topic out of the CData:
                Dim DITATopic As XmlDocument = GetXMLOut(MyNode)
                'drop the ISH version specific ProcInstr
                Dim ishnode As XmlNode = DITATopic.SelectSingleNode("/processing-instruction('ish')")
                DITATopic.RemoveChild(ishnode)
                DITATopic.Save("c:\temp\deletetopic.xml")
                'load the doc to a datablob:
                'Convert the doc to an ISH blob
                datablob = GetIshBlobFromFile("c:\temp\deletetopic.xml")
                'delete local file
                File.Delete("c:\temp\deletetopic.xml")
            Else
                'get the blob (images only) needed to create the new version:
                datablob = GetBinaryOut(newverDoc.SelectSingleNode("//ishdata"))
            End If

            'get the various required parameters needed to create the new version:
            Dim IshType As DocumentObj20.eISHType = StringToISHType(GetISHTypeFromMeta(newverDoc))
            Dim basefolder As DocumentObj25.eBaseFolder
            Dim folderpath() As String
            Dim folderID() As Long
            ISHDocObj25.FolderLocation(m_Context, GUID, basefolder, folderpath, folderID)
            'Need to first collect the current version info and then drop it.
            Dim ishfields As XmlNode = newverDoc.SelectSingleNode("//ishfields")
            'TODO: This doesn't currently handle branched objects... would need to pull the final value (after the last '.') in the string, increment it, then replace the final value.
            Dim VersionNode As XmlNode = ishfields.SelectSingleNode("ishfield[@name='VERSION']")
            Version = VersionNode.InnerText
            Version = Trim(Str(Int(Version) + 1))
            Dim delfield As XmlNode = ishfields.SelectSingleNode("ishfield[@name='VERSION']")
            ishfields.RemoveChild(delfield)
            Dim XMLMetaData As String = ishfields.OuterXml
            Dim psEDT As String = newverDoc.SelectSingleNode("//ishdata").Attributes.GetNamedItem("edt").Value
            'if the container version exists, just create the resolution within it.  otherwise, create the new version too.
            If ObjectExists(GUID, Version, Language) Then
                Try
                    'Create the new language content on the existing new version (created previously, but not populated with this resolution's content):
                    ISHDocObj.CreateOrUpdate(m_Context, folderID(folderID.Length - 1), IshType, GUID, Version, Language, Resolution, "Draft", XMLMetaData, XMLMetaData, psEDT, datablob)
                Catch ex2 As Exception
                    modErrorHandler.Errors.PrintMessage(3, "Failed to create new version of object " + GUID + "=" + Version + "=" + Language + "=" + Resolution + ". Message: " + ex2.Message, strModuleName + "-CheckOut")
                    Return False
                End Try
            Else
                Try
                    'Create the new version AND language:
                    ISHDocObj.CreateOrUpdate(m_Context, folderID(folderID.Length - 1), IshType, GUID, "new", Language, Resolution, "Draft", XMLMetaData, XMLMetaData, psEDT, datablob)
                Catch ex2 As Exception
                    modErrorHandler.Errors.PrintMessage(3, "Failed to create new version of object " + GUID + "=" + Version + "=" + Language + "=" + Resolution + ". Message: " + ex2.Message, strModuleName + "-CheckOut")
                    Return False
                End Try
            End If

        End Function
        Public Shared Function GetLatestVersionNumber(ByVal GUID As String) As String
            'Get the existing content at the current version.
            Dim VerDoc As XmlDocument = GetObjByID(GUID, "latest", "en", "")
            If VerDoc Is Nothing Then
                VerDoc = GetObjByID(GUID, "latest", "en", "Low")
            End If
            Dim ishfields As XmlNode = VerDoc.SelectSingleNode("//ishfields")
            Dim VersionNode As XmlNode = ishfields.SelectSingleNode("ishfield[@name='VERSION']")
            Return VersionNode.InnerText
        End Function
        ''' <summary>
        ''' Check to see if an object with the specified parameters exists in the CMS.  Returns true if exists.
        ''' </summary>
        Public Shared Function ObjectExists(ByVal GUID As String, ByVal Version As String, ByVal Language As String, Optional ByVal Resolution As String = "") As Boolean
            Dim MyNode As XmlNode = Nothing
            Dim MyDoc As New XmlDocument
            Dim MyMeta As New XmlDocument
            Dim XMLString As String = ""
            Dim ISHMeta As String = ""
            Dim ISHResult As String = ""
            Dim filename As String = "BROKEN_FILENAME"
            Dim extension As String = "FIX"
            Dim requestedmeta As StringBuilder = BuildRequestedMetadata()

            'Call the CMS to get our content!
            Try
                ISHResult = ISHDocObj.GetDocObj(m_Context, GUID, Version, Language, Resolution, "", "", requestedmeta.ToString, XMLString)
                Return True
            Catch ex As Exception
                'modErrorHandler.Errors.PrintMessage(3, "Failed to retrieve XML from CMS server: " + ex.Message, strModuleName)
                Return False
            End Try

        End Function
        ''' <summary>
        ''' Pulls the specified object from the CMS and saves it at the specified location.  Returns true if successful and file has been saved.
        ''' </summary>
        Public Shared Function GetObjByID(ByVal GUID As String, ByVal Version As String, ByVal Language As String, ByVal Resolution As String, ByVal SavePath As String) As Boolean
            Dim MyNode As XmlNode = Nothing
            Dim MyDoc As New XmlDocument
            Dim MyMeta As New XmlDocument
            Dim XMLString As String = ""
            Dim ISHMeta As String = ""
            Dim ISHResult As String = ""
            Dim filename As String = "BROKEN_FILENAME"
            Dim extension As String = "FIX"
            Dim requestedmeta As StringBuilder = BuildRequestedMetadata()
            'Call the CMS to get our content!
            Try
                ISHResult = ISHDocObj.GetDocObj(m_Context, GUID, Version, Language, Resolution, "", "", requestedmeta.ToString, XMLString)
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(3, "Failed to retrieve XML from CMS server: " + ex.Message, strModuleName + "-GetObjByID")
                Return False
            End Try

            'Load the XML and get the metadata:
            Try
                MyDoc.LoadXml(XMLString)
                filename = GetFilenameFromIshMeta(MyDoc)
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

            'Convert the CDATA to byte array
            Dim finalfile() As Byte
            Try
                'Convert CDATA Blob to Byte array
                finalfile = GetBinaryOut(MyNode)
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
        Public Shared Function GetObjByID(ByVal GUID As String, ByVal Version As String, ByVal Language As String, ByVal Resolution As String) As XmlDocument
            Dim MyNode As XmlNode = Nothing
            Dim MyDoc As New XmlDocument
            Dim MyMeta As New XmlDocument
            Dim XMLString As String = ""
            Dim ISHMeta As String = ""
            Dim ISHResult As String = ""
            Dim filename As String = "BROKEN_FILENAME"
            Dim extension As String = "FIX"
            Dim requestedmeta As StringBuilder = BuildRequestedMetadata()
            'Call the CMS to get our content!
            Try
                ISHResult = ISHDocObj.GetDocObj(m_Context, GUID, Version, Language, Resolution, "", "", requestedmeta.ToString, XMLString)
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

            ''Convert the CDATA to byte array
            'Dim finalfile() As Byte
            'Try
            '    'Convert CDATA Blob to Byte array
            '    finalfile = GetBinaryOut(MyNode)
            'Catch ex As Exception
            '    modErrorHandler.Errors.PrintMessage(3, "Failed to convert CDATA Blob to binary stream - no content returned from CMS: " + ex.Message, strModuleName)
            '    Return False
            'End Try
            ''Save the content out to a file:
            'Try
            '    'Create the save path if it doesn't exist:
            '    If Directory.Exists(SavePath) = False Then
            '        Directory.CreateDirectory(SavePath)
            '    End If
            '    'write to filename, bytes we extracted, don't append
            '    My.Computer.FileSystem.WriteAllBytes(SavePath + "\" + filename + "." + extension.ToLower, finalfile, False)
            '    Return True
            'Catch ex As Exception
            '    modErrorHandler.Errors.PrintMessage(3, "Failed to save returned object to a file: " + ex.Message, strModuleName)
            '    Return False
            'End Try
        End Function
        Private Shared Function GetFilenameFromIshMeta(ByVal IshMetaData As XmlDocument) As String
            Dim filename As New StringBuilder
            Try
                'first, check to make sure we actually have attributes we need:
                If IshMetaData.ChildNodes.Count = 0 Then
                    Return ""
                End If
                'Next, return the basic data that all objects have:
                filename.Append(IshMetaData.SelectSingleNode("/ishobjects/ishobject/ishfields/ishfield[contains(@name, 'FTITLE')]").InnerText + "=")
                filename.Append(IshMetaData.SelectSingleNode("/ishobjects/ishobject/@ishref").Value + "=") ' get guid
                filename.Append(IshMetaData.SelectSingleNode("/ishobjects/ishobject/ishfields/ishfield[contains(@name, 'VERSION')]").InnerText + "=")
                filename.Append(IshMetaData.SelectSingleNode("/ishobjects/ishobject/ishfields/ishfield[contains(@name, 'DOC-LANGUAGE')]").InnerText + "=")

                'last, check to see if we have an image.  If we do, we need to get the data and return it as part of the string too.  Otherwise, we just return ""
                Dim resolution As String = ""
                If Not IshMetaData.SelectSingleNode("/ishobjects/ishobject/ishfields/ishfield[contains(@name, 'FRESOLUTION')]") Is Nothing Then
                    'there is text here... let's capture it:
                    resolution = IshMetaData.SelectSingleNode("/ishobjects/ishobject/ishfields/ishfield[contains(@name, 'FRESOLUTION')]").InnerText
                End If
                filename.Append(resolution)

                Return filename.ToString
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(3, "One or more metadata was not found while trying to build a filename from a located GUID. " + ex.Message, strModuleName)
                Return ""
            End Try
            Return True
        End Function
        Private Shared Function GetBinaryOut(ByVal DITANode As XmlNode) As Byte()
            Dim CDataNode As XmlNode
            Dim CData As String = ""
            Dim decodedBytes As Byte()

            Dim settings As XmlReaderSettings
            Dim resolver As New DITAResolver()
            settings = New XmlReaderSettings()
            settings.ProhibitDtd = False
            settings.ValidationType = ValidationType.None
            settings.XmlResolver = resolver
            settings.CloseInput = True



            CDataNode = DITANode.FirstChild
            CData = CDataNode.InnerText
            decodedBytes = Convert.FromBase64String(CData)
            Return decodedBytes
        End Function
        Private Shared Function GetXMLOut(ByVal DITANode As XmlNode, Optional ByVal KeepBom As Boolean = False) As XmlDocument
            Dim CDataNode As XmlNode
            Dim CData As String = ""
            Dim decodedBytes As Byte()
            Dim decodedText As String

            Dim settings As XmlReaderSettings
            Dim resolver As New DITAResolver()
            settings = New XmlReaderSettings()
            settings.ProhibitDtd = False
            settings.ValidationType = ValidationType.None
            settings.XmlResolver = resolver
            settings.CloseInput = True



            CDataNode = DITANode.FirstChild
            CData = CDataNode.InnerText
            decodedBytes = Convert.FromBase64String(CData)
            decodedText = Encoding.Unicode.GetString(decodedBytes)
            Dim strStripBOM As String = ""
            'We're stripping the BOM off here; Disable via optional parameter if you need to keep it.
            strStripBOM = decodedText.Substring(1)



            ' Try creating the reader from the string
            Dim strReader As New StringReader(strStripBOM)

            Dim reader As XmlReader = XmlReader.Create(strReader, settings)
            ' Create the new XMLdoc and load the content into it.
            Dim doc As New XmlDocument()
            doc.Load(reader)

            Return doc
        End Function
        ''' <summary>
        ''' Given an XML Document type, returns a base-64 encoded blob that can be fed directly to the CMS.
        ''' </summary>
        Public Shared Function GetIshBlobFromXMLDoc(ByVal Doc As XmlDocument) As Byte()
            Dim Data() As Byte
            Try
                Dim numBytes As Long = Doc.OuterXml.Length
                Dim myxmlstream As System.IO.MemoryStream

                'Get the encoding of the XML Document so we know how to read it into a binary stream:
                Dim decl As XmlDeclaration
                'assume UTF-8 encoding
                Dim xmlencoding As String
                'find the real encoding:
                Try
                    decl = CType(Doc.FirstChild, XmlDeclaration)
                    'grab the encoding from the declaration
                    xmlencoding = decl.Encoding
                Catch fnf As Exception
                    ' If an error is caught, the encoding is not set.  This means we just go ahead with the UTF-8 encoding.
                    ' document might be missing an xml declaration
                    xmlencoding = "utf-8"
                    modErrorHandler.Errors.PrintMessage(2, "The XMLDocument threw an error when getting the encoding.  This xml document could be missing an XML declaration. Assuming UTF-8 to try anyway." & fnf.Message, strModuleName + "-GetIshBlobFromXMLDoc")
                End Try

                Select Case xmlencoding.ToLower
                    Case "utf-8"
                        myxmlstream = New System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(Doc.OuterXml))
                    Case "utf-16", "Unicode", "unicode"
                        myxmlstream = New System.IO.MemoryStream(System.Text.Encoding.Unicode.GetBytes(Doc.OuterXml))
                    Case "utf-32"
                        myxmlstream = New System.IO.MemoryStream(System.Text.Encoding.UTF32.GetBytes(Doc.OuterXml))
                    Case Else
                        'um. huh?  What encoding is this?
                        Return Nothing
                End Select
                'Read the stream into our binary reader (byte array) and return it.
                Dim br As New BinaryReader(myxmlstream)
                Data = br.ReadBytes(CInt(numBytes))
                br.Close()
                myxmlstream.Close()
                Return Data
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(3, "Failed to convert content to Base64 blob: " + ex.Message, strModuleName + "-GetIshBlobFromFile")
                Return Nothing
            End Try
        End Function
        ''' <summary>
        ''' Given a path to a file of any type, returns a base-64 encoded blob that can be fed directly to the CMS.
        ''' </summary>
        Public Shared Function GetIshBlobFromFile(ByVal FilePath As String) As Byte()
            Dim Data() As Byte
            Try
                Dim fInfo As New FileInfo(FilePath)
                Dim numBytes As Long = fInfo.Length
                Dim fStream As New FileStream(FilePath, FileMode.Open, FileAccess.Read)
                Dim br As New BinaryReader(fStream)
                Data = br.ReadBytes(CInt(numBytes))
                br.Close()
                fStream.Close()
                Return Data
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(3, "Failed to convert content to Base64 blob: " + ex.Message, strModuleName + "-GetIshBlobFromFile")
                Return Nothing
            End Try
        End Function
        ''' <summary>
        ''' Imports specified file using the parameters provided.  Returns true if successful.
        ''' </summary>
        ''' <param name="FilePath">Path to the file to be imported.</param>
        ''' <param name="CMSFolderID">ID of the CMS folder to import to.</param>
        ''' <param name="Author">Author of the object being imported.</param>
        ''' <param name="ISHType">Type of content being imported.  Allowed values are: "ISHIllustration", "ISHLibrary", "ISHMasterDoc", "ISHModule", "ISHNone", "ISHPublication", "ISHReusedObj", and "ISHTemplate"</param>
        ''' <param name="ReturnedGUID">Returns GUID set by CMS upon successful import.</param>
        ''' <param name="CMSTitle">Title to be used for the object within the CMS. This is shown to users of the CMS and can be updated in the object Properties.</param>
        ''' <param name="ObjectMetaType">Specifies the value found in the LOV for a module or image (Graphic, Icon, Screenshot, Concept, Reference, Task, etc.).  Value is arbitrary but is only properly set if found in the CMS.</param>
        Public Shared Function ImportObject(ByVal FilePath As String, ByVal CMSFolderID As Long, ByVal Author As String, ByVal ISHType As String, ByVal strState As String, ByRef ReturnedGUID As String, ByRef CMSTitle As String, Optional ByVal ObjectMetaType As String = "") As Boolean
            If File.Exists(FilePath) Then
                'Dim CDataNode As XmlNode
                'Dim CData As String = ""
                'Dim decodedBytes As Byte()
                'Dim decodedText As String
                Dim CMSFileName As String = ""
                Dim GUID As String = ""
                Dim Version As String = ""
                Dim Language As String = ""
                Dim Resolution As String = ""
                'Get our meta from the file:
                If GetCommonMetaFromLocalFile(FilePath, CMSFileName, CMSTitle, GUID, Version, Language, Resolution) = False Then
                    Return False
                End If
                'check to see if user has commas in their titles, report if they do and fail the import:
                If CMSTitle.Contains(",") Then
                    modErrorHandler.Errors.PrintMessage(2, "Object's title contains character(s) not allowed by the CMS. Replacing with equivalent Unicode character '‚' (Single Low Quotation Mark). Title: '" & CMSTitle & "'. File: " & FilePath & ".", strModuleName + "importobj-getmeta")
                    CMSTitle = CMSTitle.Replace(",", "‚")
                    'Return False
                End If
                'Create the MetaXML string based on our metadata:
                Dim metaxml As String
                metaxml = GetMetaDataXMLStucture(CMSTitle, Version, Author, strState, Resolution, ObjectMetaType)
                If metaxml = "" Then
                    modErrorHandler.Errors.PrintMessage(3, "Failed to generate the xml metadata needed to create the content in the CMS. Aborting import.", strModuleName + "importobj-getmeta")
                    Return False
                End If
                'now that we have the meta, need to get the bytearray data blob
                Dim data As Byte()
                data = IshDocument.GetIshBlobFromFile(FilePath)


                Dim result As String = ""
                ' Import the content if it doesn't already exist in the CMS
                If IshDocument.ObjectExists(GUID, Version, Language, Resolution) = False Then
                    Try
                        ISHDocObj.Create(m_Context, CMSFolderID.ToString, StringToISHType(ISHType), GUID, Version, Language, Resolution, metaxml, GetISHEdt(Path.GetExtension(FilePath)), data)
                        ReturnedGUID = GUID
                        ''if objectmetatype is icon, also import thumbnail as new resolution
                        'If ObjectMetaType = "Icon" Then
                        '    ISHDocObj.Create(m_Context, CMSFolderID.ToString, StringToISHType(ISHType), ReturnedGUID, Version, Language, "Thumbnail", metaxml, GetISHEdt(Path.GetExtension(FilePath)), data)
                        'End If
                        Return True
                    Catch ex As Exception
                        If ex.Message.ToString.Contains("-227") And ex.Message.ToString.Contains("Check that the id in the document is the same as the DocId provided via metadata") Then
                            modErrorHandler.Errors.PrintMessage(3, GUID + " contains an invalid ID.  Most likely, the Public ID in the DOCTYPE declaration is not recognized by the catalog lookup/DTDs. Message from CMS: " + ex.Message, strModuleName + "-importobj-CMSCreate")
                        Else
                            modErrorHandler.Errors.PrintMessage(3, "Failed to import " + GUID + " to the CMS: " + ex.Message, strModuleName + "-importobj-CMSCreate")
                        End If



                        Return False
                    End Try
                Else
                    modErrorHandler.Errors.PrintMessage(2, "Object to be imported already exists in the CMS. Skipping import for " + FilePath, strModuleName + "-ImportObject")
                    ReturnedGUID = GUID
                    Return True ' It's already imported, let the user know.
                End If
            Else
                Return False 'File didn't exist locally...
            End If

        End Function
        Public Shared Function GetISHEdt(ByVal FileExtension As String) As String
            'Valid EDT Values:
            'EDT-PDF
            'EDT-WORD
            'EDT-EXCEL
            'EDT-PPT
            'EDT-TEXT
            'EDT-HTML
            'EDT-REP-S3
            'EDT-TIFF
            'EDT-TRIDOC
            'EDTXML
            'EDTFM
            'EDTJPEG
            'EDTGIF
            'EDTUNDEFINED
            'EDTFLASH
            'EDTCGM
            'EDTBMP
            'EDTMPEG
            'EDTEPS
            'EDTAVI
            'EDTZIP
            'EDTPNG
            'EDTSVG
            'EDTSVGZ
            'EDTWMF
            'EDTRAR
            'EDTTAR
            'EDTHLP
            'EDTRTF
            'EDTHTM
            'EDTCSS
            'EDTPDF
            'EDTDOC
            'EDTXLS
            'EDTCHM
            'EDTVSD
            'EDTPSD
            'EDTPSP
            'EDTEMF
            'EDTAI

            Select Case FileExtension.Replace(".", "").ToLower
                Case "xml", "dita", "ditamap"
                    Return "EDTXML"
                Case "eps"
                    Return "EDTEPS"
                Case "jpg", "jpeg"
                    Return "EDTJPEG"
                Case "fm", "book", "mif"
                    Return "EDTFM"
                Case "xls"
                    Return "EDTXLS"
                Case "chm"
                    Return "EDTCHM"
                Case "VSD"
                    Return "EDTVSD"
                Case "psd"
                    Return "EDTPSD"
                Case "emf"
                    Return "EDTEMF"
                Case "ai"
                    Return "EDTAI"
                Case "tif", "tiff"
                    Return "EDT-TIFF"
                Case "doc"
                    Return "EDTDOC"
                Case "pdf"
                    Return "EDTPDF"
                Case "css"
                    Return "EDTCSS"
                Case "htm", "html"
                    Return "EDTHTM"
                Case "RTF"
                    Return "EDTRTF"
                Case "hlp"
                    Return "EDTHLP"
                Case "tar"
                    Return "EDTTAR"
                Case "rar"
                    Return "EDTRAR"
                Case "wmf"
                    Return "EDTWMF"
                Case "svgz"
                    Return "EDTSVGZ"
                Case "svg"
                    Return "EDTSVG"
                Case "png"
                    Return "EDTPNG"
                Case "zip"
                    Return "EDTZIP"
                Case "avi"
                    Return "EDTAVI"
                Case "mpg", "mpeg"
                    Return "EDTMPEG"
                Case "gif"
                    Return "EDTGIF"
                Case "fla", "swf"
                    Return "EDTFLASH"
                Case "bmp"
                    Return "EDTBMP"
                Case "fla"
                    Return "EDTFLASH"
                Case "pdf"
                    Return "EDTPDF"
                Case Else
                    Return "EDTUNDEFINED"
            End Select
        End Function
        Private Shared Function GetTopicTypeFromMeta(ByVal doc As XmlDocument) As String
            Dim MyNode As XmlNode
            Dim DITATopic As New XmlDocument
            Dim topictype As String = ""
            MyNode = doc.SelectSingleNode("//ishdata")
            'get the dita topic out of the CData:
            DITATopic = GetXMLOut(MyNode)
            topictype = DITATopic.DocumentElement.Name
            'topictype = DITATopic.
            Return topictype
        End Function
        Private Shared Function GetISHTypeFromMeta(ByVal doc As XmlDocument) As String
            Dim MyNode As XmlNode
            Dim ishtype As String = ""
            MyNode = doc.SelectSingleNode("//ishobject")
            'Get the ISHType:
            For Each ishattrib As XmlAttribute In MyNode.Attributes
                If ishattrib.Name = "ishtype" Then
                    ishtype = ishattrib.Value
                End If
            Next
            Return ishtype
        End Function
        Private Shared Function GetMetaDataXMLStucture(ByVal CMSTitle As String, ByVal Version As String, ByVal Author As String, ByVal Status As String, ByVal Resolution As String, Optional ByVal ModuleType As String = "", Optional ByVal Illustrator As String = "mmatus") As String
            If CMSTitle = "" Or Version = "" Or Author = "" Or Status = "" Then
                'if one or more required fields are blank, abort opperation!
                modErrorHandler.Errors.PrintMessage(3, "One or more required Metadata fields are blank. Check the Author, Status, CMSTitle, and Version values.", strModuleName + "GetMetadataXMLStructure")
                Return ""
            End If
            Dim XMLString As New StringBuilder
            XMLString.Append("<ishfields>")
            XMLString.Append("<ishfield name=""FTITLE"" level=""logical"">")
            XMLString.Append(CMSTitle)
            XMLString.Append("</ishfield>")
            'If we have a ModuleType, we need to set the metadata appropriately for the LOV.
            If Not ModuleType = "" Then
                Select Case ModuleType.ToLower
                    Case "concept", "task", "reference", "topic", "dita", "troubleshooting", "glossary", "glossary term"
                        XMLString.Append("<ishfield name=""FMODULETYPE"" level=""logical"">")
                    Case "map", "submap"
                        XMLString.Append("<ishfield name=""FMASTERTYPE"" level=""logical"">")
                    Case "graphic", "icon", "screenshot"
                        XMLString.Append("<ishfield name=""FILLUSTRATIONTYPE"" level=""logical"">")

                End Select
                XMLString.Append(ModuleType)
                XMLString.Append("</ishfield>")
            End If
            'if we have an image, need to set the default illustrator for it.
            If Resolution.Length > 0 Then
                XMLString.Append("<ishfield name=""FILLUSTRATOR"" level=""lng"">")
                XMLString.Append(Illustrator)
                XMLString.Append("</ishfield>")
                XMLString.Append("<ishfield type=""hidden"" name=""FNOTRANSLATIONMGMT"" level=""logical"" label=""Disable translation management"">No</ishfield>")
            End If
            XMLString.Append("<ishfield name=""FAUTHOR"" level=""lng"">")
            XMLString.Append(Author)
            XMLString.Append("</ishfield>")
            XMLString.Append("<ishfield name=""FSTATUS"" level=""lng"">")
            XMLString.Append(Status)
            XMLString.Append("</ishfield>")
            XMLString.Append("</ishfields>")
            Return XMLString.ToString
        End Function
        ''' <summary>
        ''' Converts a string (ISHIllustration, ISHBaseline, etc.) to a valid ISHType object.
        ''' </summary>
        Public Shared Function StringToISHType(ByVal IshType As String) As VMwareISHModulesNS.DocumentObj20.eISHType
            Select Case IshType
                Case "ISHBaseline"
                    Return DocumentObj20.eISHType.ISHBaseline
                Case "ISHIllustration"
                    Return DocumentObj20.eISHType.ISHIllustration
                Case "ISHLibrary"
                    Return DocumentObj20.eISHType.ISHLibrary
                Case "ISHMasterDoc"
                    Return DocumentObj20.eISHType.ISHMasterDoc
                Case "ISHModule"
                    Return DocumentObj20.eISHType.ISHModule
                Case "ISHNone"
                    Return DocumentObj20.eISHType.ISHNone
                Case "ISHPublication"
                    Return DocumentObj20.eISHType.ISHPublication
                Case "ISHReusedObj"
                    Return DocumentObj20.eISHType.ISHReusedObj
                Case "ISHTemplate"
                    Return DocumentObj20.eISHType.ISHTemplate
                Case Else
                    Return DocumentObj20.eISHType.ISHNone
            End Select

        End Function
        ''' <summary>
        ''' Retrieves CMS metadata from a local file including CMSFilename for XML files.  File must be exported from the CMS or preprocessed for the CMS.
        ''' </summary>
        Public Shared Function GetCommonMetaFromLocalFile(ByVal LocalFilePath As String, ByRef CMSFileName As String, ByRef GUID As String, ByRef Version As String, ByRef Language As String, ByRef Resolution As String) As Boolean
            'most commonly used to get parameters for deleting an object in the CMS...
            Dim myfile As New FileInfo(LocalFilePath)
            Dim aryMeta As New ArrayList
            Try
                For Each metapiece As String In myfile.Name.Replace(myfile.Extension, "").Split("=")
                    aryMeta.Add(metapiece)
                Next
                CMSFileName = aryMeta(0)
                GUID = aryMeta(1)
                Version = aryMeta(2)
                Language = aryMeta(3)
                If aryMeta.Count > 4 Then
                    Resolution = aryMeta(4)
                Else
                    Resolution = ""
                End If
                'Extension = myfile.Extension
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' Retrieves CMS metadata from a local file excluding CMSFilename for non-XML files.  File must be exported from the CMS or preprocessed for the CMS.
        ''' </summary>
        Public Shared Function GetCommonMetaFromLocalFile(ByVal LocalFilePath As String, ByRef CMSFileName As String, ByRef CMSTitle As String, ByRef GUID As String, ByRef Version As String, ByRef Language As String, ByRef Resolution As String) As Boolean
            'used to get metadata used for creating or modifying content in the CMS.
            Dim myfile As New FileInfo(LocalFilePath)
            Dim aryMeta As New ArrayList
            Try
                For Each metapiece As String In myfile.Name.Replace(myfile.Extension, "").Split("=")
                    aryMeta.Add(metapiece)
                Next
                CMSFileName = aryMeta(0)
                GUID = aryMeta(1)
                Version = aryMeta(2)
                Language = aryMeta(3)
                If aryMeta.Count > 4 Then
                    Resolution = aryMeta(4)
                Else
                    Resolution = ""
                End If


                Select Case myfile.Extension
                    Case ".xml", ".ditamap", ".dita"
                        ' read the xml file into an xmldoc 
                        Dim doc As New XmlDocument
                        doc = LoadFileIntoXMLDocument(LocalFilePath)
                        'Get the title info from the doc:
                        If Not doc.SelectSingleNode("//title[1]") Is Nothing Then
                            CMSTitle = doc.SelectSingleNode("//title[1]").InnerText
                        ElseIf Not doc.DocumentElement.Attributes.GetNamedItem("title") Is Nothing Then
                            CMSTitle = doc.DocumentElement.Attributes.GetNamedItem("title").InnerText
                        Else
                            CMSTitle = CMSFileName
                        End If
                        CMSTitle = CMSTitle.Replace("&", "&amp;")
                        CMSTitle = CMSTitle.Replace("<", "&lt;")
                        CMSTitle = CMSTitle.Replace(">", "&gt;")
                    Case Else 'just assign the title as the filename portion.
                        CMSTitle = CMSFileName
                End Select

                Return True
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(3, "Failed get metadata based on filename and/or content for: " + LocalFilePath + " Error: " + ex.Message, strModuleName)
                Return False
            End Try


        End Function
        Private Structure XMLTemplateStruct
            Dim map As XmlDocument
            Dim mapblob() As Byte
            Dim concept As XmlDocument
            Dim conceptblob() As Byte
            Dim task As XmlDocument
            Dim taskblob() As Byte
            Dim reference As XmlDocument
            Dim referenceblob() As Byte
            Dim troubleshooting As XmlDocument
            Dim troubleshootingblob() As Byte
        End Structure
        Private Shared XMLTemplates As New XMLTemplateStruct
        Public Sub LoadXMLTemplates()
            Dim TemplateHash As New Hashtable
            TemplateHash.Add("map.ditamap", My.Application.Info.DirectoryPath + "\templateModules\map.ditamap")
            TemplateHash.Add("concept.xml", My.Application.Info.DirectoryPath + "\templateModules\concept.xml")
            TemplateHash.Add("task.xml", My.Application.Info.DirectoryPath + "\templateModules\task.xml")
            TemplateHash.Add("reference.xml", My.Application.Info.DirectoryPath + "\templateModules\reference.xml")
            TemplateHash.Add("troubleshooting.xml", My.Application.Info.DirectoryPath + "\templateModules\troubleshooting.xml")
            For Each templatefile As DictionaryEntry In TemplateHash
                Try
                    Dim doc As XmlDocument = LoadFileIntoXMLDocument(templatefile.Value)
                    With XMLTemplates
                        Select Case templatefile.Key
                            Case "map.ditamap"
                                .map = doc
                                .mapblob = GetIshBlobFromXMLDoc(doc)
                            Case "concept.xml"
                                .concept = doc
                                .conceptblob = GetIshBlobFromXMLDoc(doc)
                            Case "reference.xml"
                                .reference = doc
                                .referenceblob = GetIshBlobFromXMLDoc(doc)
                            Case "task.xml"
                                .task = doc
                                .taskblob = GetIshBlobFromXMLDoc(doc)
                            Case "troubleshooting.xml"
                                .troubleshooting = doc
                                .troubleshootingblob = GetIshBlobFromXMLDoc(doc)
                        End Select
                    End With
                Catch ex As Exception
                    modErrorHandler.Errors.PrintMessage(3, "Unable to load template files into memory! Check that they exist in " + My.Application.Info.DirectoryPath + "\templateModules" + ". Message: " + ex.Message, strModuleName + "-LoadXMLTemplates")
                End Try
            Next

        End Sub
        Private Shared Function SetGUIDinTemplates(ByVal GUID As String) As Boolean
            Try
                With XMLTemplates
                    .map.DocumentElement.Attributes.GetNamedItem("id").Value = GUID
                    .concept.DocumentElement.Attributes.GetNamedItem("id").Value = GUID
                    .task.DocumentElement.Attributes.GetNamedItem("id").Value = GUID
                    .reference.DocumentElement.Attributes.GetNamedItem("id").Value = GUID
                    .troubleshooting.DocumentElement.Attributes.GetNamedItem("id").Value = GUID
                End With
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(2, "Unable to set GUID in XML templates. Message: " + ex.Message, strModuleName)
            End Try

        End Function

        ''' <summary>
        ''' Replaces a specified module with templated content.  
        ''' Most commonly used before attempting to recursively delete referencing modules to prevent circular references.
        ''' </summary>
        Public Shared Function ReplaceWithTemplatedContent(ByVal GUID As String, ByVal Version As String, ByVal Language As String) As Boolean
            'Set the GUIDs in our templates.
            SetGUIDinTemplates(GUID)

            'Start the replacement process:
            Dim requestedmetadata As StringBuilder = BuildRequestedMetadata()
            Dim RequestedXMLObject As String = ""
            Dim doc As New XmlDocument
            Dim IshType As String
            Dim TopicType As String
            Dim Data() As Byte

            Try
                'check out the module (must be map or topic)
                IshDocument.ISHDocObj.CheckOut(m_Context, GUID, Version, Language, "", requestedmetadata.ToString, RequestedXMLObject)
            Catch ex As Exception
                'If, for some reason, we already have an object checked out, great.  otherwise, we can't check it out for some reason.
                'Exit Code for already checking an object out is -132
                If ex.Message.Contains("-132") Then
                    'we have it checked out already, but we still need to get the object CData:
                    IshDocument.ISHDocObj.GetDocObj(m_Context, GUID, Version, Language, "", "", "", requestedmetadata.ToString, RequestedXMLObject)
                Else
                    modErrorHandler.Errors.PrintMessage(3, "Unable to checkout GUID: " + GUID + " Error: " + ex.Message, strModuleName + "-ReplaceWithTemplatedContent")
                    Return False
                End If
            End Try

            'Load the XML and get the metadata:
            doc.LoadXml(RequestedXMLObject)
            'get the ISHType from the meta
            IshType = GetISHTypeFromMeta(doc)
            Select Case IshType
                Case "ISHMasterDoc"
                    'if a map, replace the content with our template content
                    Data = XMLTemplates.mapblob

                Case "ISHModule"
                    'if a topic, find out what kind
                    TopicType = GetTopicTypeFromMeta(doc)
                    Select Case TopicType
                        Case "task"
                            Data = XMLTemplates.taskblob
                        Case "concept"
                            Data = XMLTemplates.conceptblob
                        Case "reference"
                            Data = XMLTemplates.referenceblob
                        Case "troubleshooting"
                            Data = XMLTemplates.troubleshootingblob
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
                    IshDocument.ISHDocObj.CheckIn(m_Context, GUID, Version, Language, "", "", "EDTXML", Data)
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
        Public Shared Function DeleteObjectRecursivelyByGUID(ByVal GUID As String, Optional ByVal Version As String = "1", Optional ByVal Language As String = "en", Optional ByVal Resolution As String = "", Optional ByVal DeleteSubs As Boolean = True) As Boolean
            Dim ParentModules As New Hashtable
            Dim ChildrenModules As New Hashtable
            Dim requestedmetadata As StringBuilder = BuildRequestedMetadata()
            'first, find out if the GUID exists in the CMS:
            If IshDocument.ObjectExists(GUID, Version, Language, Resolution) Then

                'We're going to delete it anyway.  Use template to replace the contents completely (if not an image).
                If Resolution = "" Then
                    If IshDocument.ReplaceWithTemplatedContent(GUID, Version, Language) = False Then
                        'modErrorHandler.Errors.PrintMessage(2, "Unable to replace content in GUID: " + GUID + " with default, template content. May not be able to delete due if the topic contains circular references to referencing modules.", strModuleName + "-RecursiveDeletion")
                    End If
                End If


                'if the guid has owners, use the list to recurse into them
                If IshReports.GetReferencingModules(GUID, ParentModules, Version, Language, requestedmetadata.ToString) Then
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
                    If IshReports.GetReferencedModules(GUID, ChildrenModules, Version, Language, requestedmetadata.ToString) Then 'if true, has children
                        'then, delete the current GUID (if it can be deleted), 
                        If CanBeDeleted(GUID, Version, Language, Resolution) Then
                            IshDocument.ObliterateGUID(GUID, Version, Language, Resolution)
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
                            IshDocument.ObliterateGUID(GUID, Version, Language, Resolution)
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
                    IshDocument.ObliterateGUID(GUID, Version, Language, Resolution)
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
        Private Shared Function ObliterateGUID(ByVal GUID As String, ByVal Version As String, ByVal Language As String, Optional ByVal Resolution As String = "High") As Boolean
            Try
                If Resolution.Length > 0 Then
                    Try
                        ISHDocObj.Delete(m_Context, GUID, Version, Language, "Low")
                        modErrorHandler.Errors.PrintMessage(1, "Deleting low resolution for: " + GUID, strModuleName + "-ObliterateGUID")

                    Catch ex As Exception

                    End Try
                    If Resolution = "High" Then
                        Try
                            ISHDocObj.Delete(m_Context, GUID, Version, Language, "High")
                            modErrorHandler.Errors.PrintMessage(1, "Deleting high resolution for: " + GUID, strModuleName + "-ObliterateGUID")

                        Catch ex As Exception

                        End Try
                    End If

                    Try
                        ISHDocObj.Delete(m_Context, GUID, Version, Language, "Thumbnail")
                        modErrorHandler.Errors.PrintMessage(1, "Deleting thumbnail for: " + GUID, strModuleName + "-ObliterateGUID")

                    Catch ex As Exception

                    End Try

                    Try
                        ISHDocObj.Delete(m_Context, GUID, Version, Language, "Source")
                        modErrorHandler.Errors.PrintMessage(1, "Deleting source resolution for: " + GUID, strModuleName + "-ObliterateGUID")

                    Catch ex As Exception

                    End Try

                End If
                Try
                    ISHDocObj.Delete(m_Context, GUID, Version, "", "")
                    modErrorHandler.Errors.PrintMessage(1, "Deleting language level for: " + GUID, strModuleName + "-ObliterateGUID")

                Catch ex As Exception

                End Try
                Try
                    ISHDocObj.Delete(m_Context, GUID, "", "", "")
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
                    ISHDocObj.Delete(m_Context, GUID, "", "", "")
                Else
                    Return False
                End If

            End Try

        End Function

        ''' <summary>
        ''' Checks to see if a specified module has no referencing modules and is not in a released state. Returns true if both conditions are met.
        ''' </summary>
        ''' <param name="GUID">GUID of object in CMS</param>
        ''' <param name="Version">Version of object in CMS</param>
        ''' <param name="Language">Language of object in CMS</param>
        ''' <param name="Resolution">Resolution of object in CMS</param>
        ''' <returns>Boolean</returns>
        Public Shared Function CanBeDeleted(ByVal GUID As String, Optional ByVal Version As String = "1", Optional ByVal Language As String = "en", Optional ByVal Resolution As String = "") As Boolean
            Dim ModuleHash As New Hashtable

            If IshReports.GetReferencingModules(GUID, ModuleHash, Version, Language) = True Then
                'Has parents, can't be deleted
                Return False
            Else
                'This object has no parents.  first check passed.  Continue

            End If
            Dim SearchResult As String = ""
            IshDocument.ISHDocObj.GetMetaData(m_Context, GUID, Version, Language, Resolution, "<ishfields><ishfield name=""FSTATUS"" level=""lng""/></ishfields>", SearchResult)
            If SearchResult.Contains("""FSTATUS"" level=""lng"">Released") Then
                'Status is released.  Can't delete
                Return False
            Else
                'Status is something else, can delete
                Return True
            End If




        End Function
        Public Shared Function _ResetRecursionHash()
            m_recursion_hash.Clear()
            Return True
        End Function
        Public Shared Function ChangeAssigneeRecursively(ByVal GUID As String, ByVal Version As String, ByVal NewAuthor As String, ByVal Role As String, Optional ByVal Resolution As String = "") As Boolean
            'Dim requestmetadata As StringBuilder = BuildRequestedMetadata()
            Dim RequestedXMLObject As String = ""
            Dim doc As New XmlDocument
            Dim docorig As New XmlDocument
            Try
                VMwareISHModulesNS.clsISHObj.IshDocument.ISHDocObj.GetMetaData(m_Context, GUID, Version, "en", Resolution, BuildRequestedMetadata().ToString, RequestedXMLObject)
            Catch ex As Exception
                'failed to get object
                modErrorHandler.Errors.PrintMessage(2, "Failed to get object in DB. Info: " + GUID + "=" + Version + ". Message: " + ex.Message.ToString, strModuleName + "-ChangeAssigneeRecursively")
                Return False
            End Try

            'Load the XML and get the metadata:
            doc.LoadXml(RequestedXMLObject)

            'keep the original for matching later.
            docorig.LoadXml(RequestedXMLObject)
            Dim IshType As String = GetISHTypeFromMeta(doc)

            'Get the children and recurse if applicable (by type).
            Select Case IshType
                Case "ISHMasterDoc", "ISHModule"
                    'if a map or topic, get children
                    'Dim CurMeta As Object = IshReports.GetReportedObjects(doc)
                    Dim children As New Hashtable
                    IshReports.GetReferencedModules(GUID, children, Version, "en")
                    For Each childmodule As DictionaryEntry In children
                        If Not childmodule.Value.GUID = GUID And Not m_recursion_hash.Contains(childmodule.Value.GUID) Then
                            'Track the GUID we're diving into so that we don't accidently try to process it again later (endless looping).
                            m_recursion_hash.Add(childmodule.Value.GUID, childmodule.Value.GUID)
                            ChangeAssigneeRecursively(childmodule.Value.GUID, childmodule.Value.version, NewAuthor, Role, childmodule.Value.Resolution)
                        End If
                    Next
                Case Else
                    'Returned something that we don't need to parse for children...
                    'Return False
            End Select

            'change owner by GUID Stuff
            Dim ishfields As XmlNode = doc.SelectSingleNode("//ishfields")
            Dim ishfieldsorig As XmlNode = docorig.SelectSingleNode("//ishfields")
            Dim funame As XmlNode = ishfields.SelectSingleNode("ishfield[@name='" + Role + "']")
            'If there's no currently assigned name to the role specified, we need to insert it.
            If funame Is Nothing Then
                Dim funamedoc As New XmlDocument()
                funamedoc.LoadXml("<funame><ishfield name='" + Role + "' level='lng'>" + NewAuthor + "</ishfield></funame>")
                funame = funamedoc.FirstChild
                ishfields.AppendChild(doc.ImportNode(funame.FirstChild, True))
                funame = ishfields.LastChild
            End If
            funame.InnerText = NewAuthor
            Dim ishver As XmlNode = ishfields.SelectSingleNode("ishfield[@name='VERSION']")
            ishfields.RemoveChild(ishver)
            'change the owner of the GUID at the specified ver
            Select Case IshType
                Case "ISHMasterDoc", "ISHModule"
                    If Not Role = "FILLUSTRATOR" Then
                        'if a map or topic, update the guid with simple command
                        Try
                            VMwareISHModulesNS.clsISHObj.IshDocument.ISHDocObj.SetMetaData(m_Context, GUID, Version, "en", "", ishfields.OuterXml.ToString, RequestedXMLObject)
                        Catch ex As Exception
                            modErrorHandler.Errors.PrintMessage(2, "Unable to change assignee on object " + GUID + "=" + Version + ". Message: " + ex.Message.ToString, strModuleName + "-ChangeAssigneeRecursively")
                            Return False
                        End Try
                    End If
                Case "ISHIllustration"
                    If Not Role = "FEDITOR" And Not Role = "FCODEREVIEWER" Then
                        'If illustration, need to update all possible resolutions.
                        'For Images, we need to replace ALL instances which means we need to trick it into thinking that we've returned the matching resolution for each type.

                        'Start by making the original res match the High resolution.
                        Dim res As XmlNode = doc.SelectSingleNode("//ishfields/ishfield[@name='FRESOLUTION']")
                        Dim resOrig As XmlNode = docorig.SelectSingleNode("//ishfields/ishfield[@name='FRESOLUTION']")
                        res.InnerText = "High"
                        resOrig.InnerText = "High"


                        Try
                            VMwareISHModulesNS.clsISHObj.IshDocument.ISHDocObj.SetMetaData(m_Context, GUID, Version, "en", "High", ishfields.OuterXml.ToString, ishfieldsorig.OuterXml.ToString)
                        Catch ex As Exception

                        End Try
                        res.InnerText = "Low"
                        resOrig.InnerText = "Low"
                        Try
                            VMwareISHModulesNS.clsISHObj.IshDocument.ISHDocObj.SetMetaData(m_Context, GUID, Version, "en", "Low", ishfields.OuterXml.ToString, ishfieldsorig.OuterXml.ToString)
                        Catch ex As Exception

                        End Try
                        res.InnerText = "Thumbnail"
                        resOrig.InnerText = "Thumbnail"
                        Try
                            VMwareISHModulesNS.clsISHObj.IshDocument.ISHDocObj.SetMetaData(m_Context, GUID, Version, "en", "Thumbnail", ishfields.OuterXml.ToString, ishfieldsorig.OuterXml.ToString)
                        Catch ex As Exception

                        End Try
                        res.InnerText = "Source"
                        resOrig.InnerText = "Source"
                        Try
                            VMwareISHModulesNS.clsISHObj.IshDocument.ISHDocObj.SetMetaData(m_Context, GUID, Version, "en", "Source", ishfields.OuterXml.ToString, ishfieldsorig.OuterXml.ToString)
                        Catch ex As Exception

                        End Try
                    End If
                Case Else
                    'Something else altogether... Not sure what to do here.
            End Select
            'Track the GUID we're diving into so that we don't accidently try to process it again later (endless looping).
            If Not m_recursion_hash.Contains(GUID) Then
                m_recursion_hash.Add(GUID, GUID)
            End If
            Return True


        End Function

        Public Shared Function CanMoveToState(ByVal strState As String, ByVal strGUID As String, ByVal strVersion As String, ByVal strResolution As String, Optional ByVal strLanguage As String = "en") As Boolean

            Dim OutStates As String()
            Try
                ' Declare variable for the Application service
                'Dim DocService As ISDoc.DocumentObj20 = New ISDoc.DocumentObj20()

                ' Clear variable for the result
                ISHDocObj.GetPossibleTransitionStates(m_Context, strGUID, strVersion, strLanguage, strResolution, OutStates)

            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(2, "Error getting possible transition states for " + strGUID + "" + strVersion + "" + strLanguage + ". Message: " + ex.Message.ToString, strModuleName + "-CanMoveToState")
                Return False
            End Try

            ' Now check to see if we can move to desired state:
            If (OutStates.Length > 0) Then
                Dim s As String
                For Each s In OutStates
                    If (s = strState) Then
                        CanMoveToState = True
                        Exit Function
                    End If
                Next
            End If
            Return False
        End Function

        Public Shared Function SetMeta(ByVal strMeta As String, ByVal GUID As String, ByVal Version As String, ByVal strResolution As String, Optional ByVal Language As String = "en") As Boolean



            Try
                ' Declare variable for the Application service
                'Dim DocService As ISDoc.DocumentObj20 = New ISDoc.DocumentObj20()

                ' Clear variable for the result
                ISHDocObj.SetMetaData(m_Context, GUID, Version, Language, strResolution, strMeta, "")



            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(2, "Error setting meta for " + GUID + "" + Version + "" + Language + ". Message: " + ex.Message.ToString, strModuleName + "-SetMeta")
                Return False
            End Try


            Return True
        End Function
        Public Shared Function GetCurrentState(ByVal GUID As String, ByVal Version As String, ByVal strResolution As String, Optional ByVal Language As String = "en") As String

            Dim state As String = "nothing"
            Dim OutXML As String = ""
            Try
                '' Declare variable for the Application service
                'Dim DocService As IshDocument.DocumentObj20 = New ISDoc.DocumentObj20()

                Dim strMeta As String = "<ishfields><ishfield name=""FSTATUS"" level=""lng""/></ishfields>"

                ISHDocObj.GetMetaData(m_Context, GUID, Version, Language, strResolution, strMeta, OutXML)

            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(2, "Error getting current state for " + GUID + "" + Version + "" + Language + ". Message: " + ex.Message.ToString, strModuleName + "-GetCurrentState")
                Return False
            End Try

            Dim strFind As String = "<ishfield name=""FSTATUS"" level=""lng"">"
            state = OutXML.Substring(OutXML.LastIndexOf(strFind) + strFind.Length)
            state = state.Remove(state.LastIndexOf("</ishfield>"))


            Return state
        End Function


        'Public Shared Function GetChildren(ByVal GUID As String, ByVal Version As String) As ArrayList
        '    Dim RealRequestedMeta As New StringBuilder
        '    If RequestedMetadata = "" Then
        '        RealRequestedMeta = BuildRequestedMetadata()
        '    Else
        '        RealRequestedMeta.Append(RequestedMetadata)
        '    End If
        '    Dim ChildrenModules As New Hashtable
        '    If IshReports.GetReferencedModules(GUID, ChildrenModules, Version, "en", requestedmetadata.ToString) Then 'if true, has children
        '        'then, delete the current GUID (if it can be deleted), 
        '        If CanBeDeleted(GUID, Version, Language, Resolution) Then
        '            IshDocument.ObliterateGUID(GUID, Version, Language, Resolution)
        '        Else
        '            'Can't be deleted for some reason...
        '            modErrorHandler.Errors.PrintMessage(3, "Unable to delete GUID: " + GUID, strModuleName + "RecursiveDeletion")
        '            Return False
        '        End If
        '        'check to see if we should also delete children of the current module
        '        If DeleteSubs = True Then
        '            'recursedelete into each of the children
        '            For Each childmodule As DictionaryEntry In ChildrenModules
        '                'Make sure we don't recurse on ourselves
        '                If Not childmodule.Value.GUID = GUID Then
        '                    If DeleteObjectRecursivelyByGUID(childmodule.Value.GUID, childmodule.Value.Version, childmodule.Value.language, childmodule.Value.resolution, DeleteSubs) = False Then
        '                        modErrorHandler.Errors.PrintMessage(3, "Failed to delete descendant of: " + GUID, strModuleName + "RecursiveDeletion")
        '                        Return False
        '                    End If
        '                End If
        '            Next
        '        End If
        '        Return True
        '    End If
        'End Function

        Public Sub New()

        End Sub

    End Class
    Public Class IshBaseline
        Public Shared ISHBaselineObj As New Baseline20.BaseLine20
    End Class
    Public Class IshConditions
        Public Shared ISHCondObj As New Condition20.Condition20
    End Class
    Public Class IshMeta
        Public Shared ISHMetaObj As New MetaDataAssist20.MetaDataAssist20
        ''' <summary>
        ''' Determines if a user has priviledges of a particular role and belongs to a specified group.
        ''' </summary>
        ''' <param name="Username">Username in the CMS.</param>
        ''' <param name="Role">Role priviledge such as "Administrator", "Author", "Illustrator", etc.</param>
        ''' <param name="UserGroup">(OPTIONAL) The group to search within ("TCL", for instance).</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function UserHasRole(ByVal Username As String, ByVal Role As String, Optional ByVal UserGroup As String = "TCL") As Boolean
            Dim returneduserlist() As String
            Dim userlist As New ArrayList

            VMwareISHModulesNS.clsISHObj.IshMeta.ISHMetaObj.GetUsers(m_Context, Role, UserGroup, returneduserlist)

            For Each uname As String In returneduserlist
                userlist.Add(uname)
            Next
            If userlist.Contains(Username) Then
                Return True
            Else
                Return False
            End If


        End Function
    End Class
    Public Class IshOutput
        Public Shared ISHOutputObj As New OutputFormat20.OutputFormat20
    End Class
    Public Class IshPub
        Public Shared ISHPubObj As New Publication20.Publication20
    End Class
    Public Class IshPubOutput
        Public Shared ISHPubOutObj As New PublicationOutput20.PublicationOutput20
    End Class
    Public Class IshReports
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
        Public Shared ISHReportsObj As New Reports20.Reports20
        ''' <summary>
        ''' Given a commonly returned "ishobjects" XML Document returned from most CMS queries, 
        ''' this function converts each entry to Dictionary Entries in a hashtable.  Each entry 
        ''' uses a combined GUID+meta key that appears the same as the commonly used CMS filenames. 
        ''' The value is the ObjectData structure found in this class.
        ''' </summary>
        ''' <param name="XMLDoc">A standard "ishobjects" XML Document returned from a query to the CMS.</param>
        ''' <returns>A Hashtable containing all of the ishobjects and their metadata.</returns>
        Public Shared Function GetReportedObjects(ByVal XMLDoc As XmlDocument) As Hashtable
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
        ''' Finds all objects that refer to the specified ISHObject as parents.
        ''' </summary>
        ''' <param name="GUID"></param>
        ''' <param name="ReferencingModules">Returns a hashtable of all referencing objects.</param>
        ''' <param name="Version"></param>
        ''' <param name="Language"></param>
        ''' <param name="RequestedMetadata">You can limit the metadata to retrieve from the CMS using the common templating structure for CMS metadata requests.  Expected format is string of XML content.</param>        
        Public Shared Function GetReferencingModules(ByVal GUID As String, ByRef ReferencingModules As Hashtable, Optional ByVal Version As String = "1", Optional ByVal Language As String = "en", Optional ByVal RequestedMetadata As String = "") As Boolean
            Dim requestedobjects As String = ""
            Dim RealRequestedMeta As New StringBuilder
            If RequestedMetadata = "" Then
                RealRequestedMeta = BuildRequestedMetadata()
            Else
                RealRequestedMeta.Append(RequestedMetadata)
            End If

            'Returns a list of referencing objects to "requestedobjects" string as xml:
            ISHReportsObj.GetReferencedByDocObj(m_Context, GUID, Version, Language, False, RealRequestedMeta.ToString, requestedobjects)
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
        ''' <summary>
        ''' Finds all objects that are referred to by the specified ISHObject as children.
        ''' </summary>
        ''' <param name="GUID"></param>
        ''' <param name="ReferencedModules">Returns a hashtable of all referenced modules.</param>
        ''' <param name="Version"></param>
        ''' <param name="Language"></param>
        ''' <param name="RequestedMetadata">You can limit the metadata to retrieve from the CMS using the common templating structure for CMS metadata requests.  Expected format is string of XML content.</param>        
        Public Shared Function GetReferencedModules(ByVal GUID As String, ByRef ReferencedModules As Hashtable, Optional ByVal Version As String = "1", Optional ByVal Language As String = "en", Optional ByVal RequestedMetadata As String = "") As Boolean
            Dim requestedobjects As String = ""
            Dim RealRequestedMeta As New StringBuilder
            If RequestedMetadata = "" Then
                RealRequestedMeta = BuildRequestedMetadata()
            Else
                RealRequestedMeta.Append(RequestedMetadata)
            End If


            'Returns a list of referencing objects to "requestedobjects" string as xml:
            ISHReportsObj.GetReferencedDocObj(m_Context, GUID, Version, Language, False, RealRequestedMeta.ToString, requestedobjects)
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
    End Class
    Public Class IshSearch
        Public Shared ISHSearchObj As New Search20.Search20
    End Class
    Public Class IshWorkflow
        Public Shared ISHWorkflowObj As New Workflow20.WorkFlow20
    End Class
    Public Class IshFolder
        Public Shared ISHFolderObj As New Folder20.Folder20
        ''' <summary>
        ''' Creates a new folder as a sub-folder to the given ParentFolderID and returns the newly created ID.  If a folder of the same name already exists, the function returns the ID of that folder instead.
        ''' </summary>
        ''' <param name="ParentFolderID">Containing folder's ID.</param>
        ''' <param name="ISHType">Folder type: ISHNone, ISHMasterDoc, ISHModule, ISHIllustration, ISHReusedObj, ISHTemplate, and ISHLibrary</param>
        ''' <param name="FolderName">Name of folder to create/reuse.</param>
        ''' <param name="NewFolderID">New/reused folder's ID.</param>
        ''' <param name="OwnershipGroup">Ownership group based on groups defined in the CMS.</param>
        ''' <param name="ReadAccessList">Comma-separated string with user groups that have read access to this folder.</param>
        Public Shared Function CreateOrUseFolder(ByVal ParentFolderID As Long, ByVal ISHType As String, ByVal FolderName As String, ByRef NewFolderID As Long, Optional ByVal OwnershipGroup As String = "", Optional ByVal ReadAccess As String = "") As String
            Dim CreateFolderResult As String
            Dim RealIshType As Folder20.eISHType
            Select Case ISHType
                Case "ISHNone"
                    RealIshType = Folder20.eISHType.ISHNone
                Case "ISHMasterDoc"
                    RealIshType = Folder20.eISHType.ISHMasterDoc
                Case "ISHModule"
                    RealIshType = Folder20.eISHType.ISHModule
                Case "ISHIllustration"
                    RealIshType = Folder20.eISHType.ISHIllustration
                Case "ISHReusedObj"
                    RealIshType = Folder20.eISHType.ISHReusedObj
                Case "ISHTemplate"
                    RealIshType = Folder20.eISHType.ISHTemplate
                Case "ISHLibrary"
                    RealIshType = Folder20.eISHType.ISHLibrary
                Case "ISHPublication"
                    RealIshType = Folder20.eISHType.ISHPublication
            End Select
            If FolderExists(ParentFolderID, FolderName) Then
                NewFolderID = GetFolderIDByName(ParentFolderID, FolderName, OwnershipGroup, ReadAccess)
                CreateFolderResult = "Reused existing ID."
            Else
                CreateFolderResult = ISHFolderObj.Create(m_Context, ParentFolderID, RealIshType, FolderName, OwnershipGroup, NewFolderID, ReadAccess)
            End If

            Return CreateFolderResult
        End Function

        ''' <summary>
        ''' Function that returns the ID of a folder specified by a full CMS Path
        ''' Remember to include //Doc at the beginning.
        ''' </summary>
        ''' <param name="FullCMSPath"></param>
        Public Shared Function GetFolderIDByPath(ByVal FullCMSPath As String, Optional ByRef Result As String = "", Optional ByRef OwnershipGroup As String = "", Optional ByRef ReadAccess As String = "") As Long
            FullCMSPath = FullCMSPath.Trim()
            'First, make sure it's a properly formated CMS Path, starting with "//"
            If FullCMSPath.StartsWith("//") = False Then
                'invalid path given - must start with //
                Return -1
            Else
                'trim "//" off the front and any / at the end (if there is one)
                FullCMSPath = FullCMSPath.Replace("//", "")
                FullCMSPath = FullCMSPath.TrimEnd("/")
            End If
            ' Assume the root ID is 0 and drop the first folder that corresponds to it.
            Dim currentID As Long = 0
            'if we have more than just the root, we need to trim the root folder name off.
            If FullCMSPath.Contains("/") Then
                FullCMSPath = FullCMSPath.Remove(0, FullCMSPath.IndexOf("/") + 1)
            Else
                'if we JUST have the root folder, make sure the user called it the right name.
                Dim RealRootName As String = ""
                Dim rootishtype As VMwareISHModulesNS.Folder20.eISHFolderType
                Dim OutQuery As String = ""
                Dim OwnedBy As String = ""
                Dim rootfolderid As Long = GetFolderIDByName(0, FullCMSPath, OwnershipGroup, ReadAccess)
                If rootfolderid > -1 Then
                    IshFolder.ISHFolderObj.GetProperties(m_Context, rootfolderid, RealRootName, rootishtype, OutQuery, OwnershipGroup, ReadAccess)
                    'if it is the right name, return 0
                    If RealRootName = FullCMSPath Then
                        Return 0
                    Else
                        'otherwise, return -1
                        Result = "Invalid name for root folder specified."
                        Return -1
                    End If
                Else
                    Result = "Invalid name for root folder specified."
                    Return -1
                End If


            End If

            Dim subID As Long
            'Check to see if we have sub-folders...
            If FullCMSPath.Length > 0 Then
                'Loop through each foldername
                For Each foldername As String In FullCMSPath.Split("/")
                    subID = GetFolderIDByName(currentID, foldername, OwnershipGroup, ReadAccess)
                    If Not subID = -1 Then
                        currentID = subID
                    Else
                        ' We hit a directory in our path that doesn't exist in the CMS.  This will return a value of "0" to the calling function.  Root is invalid for a sub.
                        Result = "Path doesn't exist in the currently specified CMS."
                        Return -1
                    End If
                Next
            Else
                'no subfolders so just do the import at the root.
                currentID = 0
            End If

            Return currentID
        End Function
        Private Shared Function FolderContainsGUID(ByVal GUID As String, ByVal FolderID As Long) As Boolean
            If Not FolderID = 0 Then
                Dim requestedobjects As String = ""


                Dim RealRequestedMeta As New StringBuilder
                RealRequestedMeta = BuildRequestedMetadata()



                'Returns a list of referencing objects to "requestedobjects" string as xml:
                Try
                    ISHFolderObj.GetContents(m_Context, FolderID, "", RealRequestedMeta.ToString, requestedobjects)
                Catch ex As Exception
                    modErrorHandler.Errors.PrintMessage(2, "Unable to get contents of folder " + FolderID, strModuleName + "-FolderContainsGUID")
                End Try


                'Load the string as an xmldoc
                Dim doc As New XmlDocument
                doc.LoadXml(requestedobjects)
                'get our returned objects into a hashtable (key=GUID, Value=Information + XMLNode of returned data)
                Dim SubObjects As Hashtable = IshReports.GetReportedObjects(doc)
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
        Public Shared Function FindFolderIDforObjbyGUID(ByVal GUID As String, ByVal ParentFolderID As Long) As Long
            'Recursively searches for a GUID and returns the ID when found.  returns -1 if not found.

            'First, check to see if our GUID exists in this folder (only if not root folder):
            If IshFolder.FolderContainsGUID(GUID, ParentFolderID) Then
                'if we found it, return this as the valid parent!
                Return ParentFolderID
                Exit Function
            Else

                'if not, we need to dive deeper.
                'get a folderlist of all children of currentfolderid
                Dim subfolderlistXML As String = ""
                Dim CMSReply As String
                Try
                    CMSReply = ISHFolderObj.GetSubFolders(m_Context, ParentFolderID.ToString, 1, subfolderlistXML)
                Catch ex As Exception

                End Try
                'Load the subfolderlistXML into an xml document
                ' Create the reader from the string.
                Dim strReader As New StringReader(subfolderlistXML)

                Dim reader As XmlReader = XmlReader.Create(strReader)
                ' Create the new XMLdoc and load the content into it.
                Dim doc As New XmlDocument()
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


        ''' <summary>
        ''' Function that returns the ID of a subfolder specified by its name
        ''' </summary>
        ''' <param name="ParentFOlderID">ID of the Parent Folder</param>
        ''' <param name="SubFolderName">Name of the subfolder to find</param>
        Public Shared Function GetFolderIDByName(ByVal ParentFolderID As Long, ByVal SubFolderName As String, ByRef OwnershipGroup As String, ByRef ReadAccess As String) As Long
            'get a folderlist of all children of currentfolderid
            Dim subfolderlistXML As String = ""
            Dim CMSReply As String
            Try
                CMSReply = ISHFolderObj.GetSubFolders(m_Context, ParentFolderID.ToString, 1, subfolderlistXML)
            Catch ex As Exception

            End Try
            'Load the subfolderlistXML into an xml document
            ' Create the reader from the string.
            Dim strReader As New StringReader(subfolderlistXML)

            Dim reader As XmlReader = XmlReader.Create(strReader)
            ' Create the new XMLdoc and load the content into it.
            Dim doc As New XmlDocument()
            doc.Load(reader)

            'find the id of the one matching the name we're looking for 
            Dim subid As String = ""
            Dim foldername As String = ""
            Dim curishtype As VMwareISHModulesNS.Folder20.eISHFolderType
            Dim curoutquery As String = ""
            Try
                subid = doc.SelectSingleNode("//ishfolders/ishfolder[@name='" + SubFolderName + "']/@ishfolderref").Value.ToString
                IshFolder.ISHFolderObj.GetProperties(m_Context, Convert.ToInt64(subid), foldername, curishtype, curoutquery, OwnershipGroup, ReadAccess)
                Return Convert.ToInt64(subid)
            Catch ex As Exception
                ' Unable to find the node we were looking for in the returned XML.  
                Return -1
            End Try
        End Function
        ''' <summary>
        ''' Returns an array of long integers containing all the ids of child sub-folders of the given folder ID.
        ''' </summary>
        ''' <param name="FolderID">ID of folder to return sub-ids from.</param>
        ''' <returns>Array of Long Integers indicating children of specified folder ID.</returns>
        ''' <remarks>If the specified FolderID is invalid, the arraylist will be empty.  It is recommended to check to see if the folder ID exists in the CMS before running this method.</remarks>
        Public Shared Function GetSubFolderIDs(ByVal FolderID As Long) As ArrayList
            Dim subfolderIDs As New ArrayList
            'get a folderlist of all children of currentfolderid
            Dim subfolderlistXML As String = ""
            Dim CMSReply As String
            Try
                CMSReply = ISHFolderObj.GetSubFolders(m_Context, FolderID.ToString, 1, subfolderlistXML)
            Catch ex As Exception

            End Try
            'Load the subfolderlistXML into an xml document
            ' Create the reader from the string.
            Dim strReader As New StringReader(subfolderlistXML)

            Dim reader As XmlReader = XmlReader.Create(strReader)
            ' Create the new XMLdoc and load the content into it.
            Dim doc As New XmlDocument()
            doc.Load(reader)

            If doc.HasChildNodes Then
                For Each subid As XmlAttribute In doc.SelectNodes("//ishfolders/ishfolder/@ishfolderref")
                    subfolderIDs.Add(Convert.ToInt64(subid.Value.ToString))
                Next
            End If
            Return subfolderIDs
        End Function
        ''' <summary>
        ''' Checks to see if a subfolder of a given ID exists within a parentfolder specified by a separate ID.
        ''' </summary>
        Public Shared Function FolderExists(ByVal ParentFolderID As Long, ByVal SubFolderName As String) As Boolean
            Dim OwnershipGroup As String = ""
            Dim ReadAccess As String = ""
            Dim subID As Long = GetFolderIDByName(ParentFolderID, SubFolderName, OwnershipGroup, ReadAccess)

            If subID = -1 Then ' Subfolder not found
                Return False
            Else 'Subfolder found
                Return True
            End If
        End Function
        ''' <summary>
        ''' Checks to see if a given folder has sub-content (objects).
        ''' </summary>
        Public Shared Function FolderHasContents(ByVal FolderID As Long) As Boolean
            Dim foldercontentXML As String = ""
            Dim CMSReply As String
            Try
                CMSReply = ISHFolderObj.GetContents(m_Context, FolderID.ToString, "", "", foldercontentXML)
            Catch ex As Exception

            End Try
            'Load the subfolderlistXML into an xml document
            ' Create the reader from the string.
            Dim strReader As New StringReader(foldercontentXML)

            Dim reader As XmlReader = XmlReader.Create(strReader)
            ' Create the new XMLdoc and load the content into it.
            Dim doc As New XmlDocument()
            doc.Load(reader)

            Dim significantchildren As XmlNodeList = doc.SelectNodes("//ishobject")
            If significantchildren.Count > 0 Then 'children found
                Return True
            Else ' no children found
                Return False
            End If
        End Function
        ''' <summary>
        ''' Checks to see if a given folder has sub-folders.
        ''' </summary>
        Public Shared Function FolderHasSubFolders(ByVal FolderID As Long)
            'get a folderlist of all children of currentfolderid
            Dim subfolderlistXML As String = ""
            Dim CMSReply As String
            Try
                CMSReply = ISHFolderObj.GetSubFolders(m_Context, FolderID.ToString, 1, subfolderlistXML)
            Catch ex As Exception

            End Try
            'Load the subfolderlistXML into an xml document
            ' Create the reader from the string.
            Dim strReader As New StringReader(subfolderlistXML)

            Dim reader As XmlReader = XmlReader.Create(strReader)
            ' Create the new XMLdoc and load the content into it.
            Dim doc As New XmlDocument()
            doc.Load(reader)
            Dim subfolders As XmlNodeList = doc.SelectNodes("//ishfolder")
            If subfolders.Count = 0 Then
                Return False
            Else
                Return True
            End If
        End Function

        ''' <summary>
        ''' Attempts to delete all content (not folders) of a specified folder. Uses IshDocument.DeleteObjectRecursivelyByGUID with "DeleteSubs" set to False for each sub-object.
        ''' </summary>
        ''' <returns>Returns true if successful.</returns>
        Public Shared Function DeleteSubContent(ByVal FolderID As Long) As Boolean
            Dim RealRequestedMeta As New StringBuilder
            RealRequestedMeta = BuildRequestedMetadata()
            Dim requestedmeta As String = RealRequestedMeta.ToString
            Dim foldercontentXML As String = ""
            Dim CMSReply As String
            Try
                CMSReply = ISHFolderObj.GetContents(m_Context, FolderID.ToString, "", requestedmeta, foldercontentXML)
            Catch ex As Exception

            End Try
            'Load the subcontentlistXML into an xml document
            ' Create the reader from the string.
            Dim strReader As New StringReader(foldercontentXML)

            Dim reader As XmlReader = XmlReader.Create(strReader)
            ' Create the new XMLdoc and load the content into it.
            Dim doc As New XmlDocument()
            doc.Load(reader)

            Dim significantchildren As XmlNodeList = doc.SelectNodes("//ishobject")
            'First, collect the info from each xml node.  Use the structure found in IshDocument
            If significantchildren.Count > 0 Then
                Dim ContentHash As New Hashtable
                ContentHash = IshReports.GetReportedObjects(doc)
                For Each child As DictionaryEntry In ContentHash
                    If IshDocument.DeleteObjectRecursivelyByGUID(child.Value.GUID, child.Value.Version, child.Value.Language, child.Value.Resolution, False) Then
                        modErrorHandler.Errors.PrintMessage(1, "Deleting object in targeted folder. GUID: " + child.Key.ToString, strModuleName)
                        'Return True
                    Else
                        modErrorHandler.Errors.PrintMessage(2, "Unable to delete object in targeted folder. GUID: " + child.Key.ToString, strModuleName)
                        'Return False
                    End If
                Next
            Else
                ' no significant modules to delete, carry on.
                Return True
            End If
            Return True
        End Function
        ''' <summary>
        ''' Function that deletes specified folder and all empty sub-folders if they are empty. Returns true if successful, false if not. Optionally allows all sub-content to be deleted as well (using DeleteSubContent).
        ''' Remember to include // at the beginning.
        ''' </summary>
        ''' <param name="delresult">Stores first error message if anything goes wrong during delete.</param>
        Public Shared Function DeleteFolderRecursive(ByVal FolderID As Long, ByRef DelResult As String, Optional ByVal DeleteContentToo As Boolean = True) As Boolean
            Dim process_result As Boolean = False
            Dim hassubs As Boolean = FolderHasSubFolders(FolderID)
            If hassubs And FolderID > 0 Then
                'MsgBox("Has SubFolders")
                For Each subfolderID As Long In GetSubFolderIDs(FolderID)
                    DeleteFolderRecursive(subfolderID, DelResult, DeleteContentToo)
                    If DelResult.Length > 0 Then
                        'Return False
                    End If
                Next
            End If
            If DeleteContentToo = True Then
                If FolderHasContents(FolderID) Then
                    'delete subcontent
                    If DeleteSubContent(FolderID) Then
                        process_result = True
                    Else
                        process_result = False
                    End If
                    'delete folder
                    'Return DeleteFolder(FolderID, DelResult)
                End If
            End If


            'all subfolders and sub-conent have been deleted.  Should be able to delete unless content was left in folders by user... if that's the case, we should return false.
            Return DeleteFolder(FolderID, DelResult)



        End Function

        ''' <summary>
        ''' Deletes a specified empty folder.
        ''' </summary>
        ''' <param name="DelResult">Result of the attempt.</param>
        ''' <returns>Returns True if able to delete the folder, False if not.</returns>
        Public Shared Function DeleteFolder(ByVal folderID As Long, ByRef DelResult As String) As Boolean
            If FolderHasContents(folderID) = False Then
                Try
                    ISHFolderObj.Delete(m_Context, folderID)
                    DelResult = ""
                    Return True
                Catch ex As Exception
                    DelResult = "Error deleting FolderID: " + folderID.ToString + " Message: " + ex.Message
                    modErrorHandler.Errors.PrintMessage(3, DelResult, strModuleName)
                    Return False
                End Try


            Else
                DelResult = "Folder identified by ID#" + folderID.ToString + " has sub-content.  Unable to delete."
                modErrorHandler.Errors.PrintMessage(2, DelResult, strModuleName)
                Return False
            End If

        End Function

    End Class



    Private Class DITAResolver
        Inherits XmlUrlResolver
        Private ReadOnly Property strModuleName() As String
            Get
                strModuleName = "DITAResolver"
            End Get
        End Property
        Public Shared myHash As New Hashtable()

        Public Sub New()
        End Sub 'New

        Public Overrides Function GetEntity(ByVal absoluteUri As Uri, ByVal role As String, ByVal ofObjectToReturn As Type) As Object
            Dim pubid1slash As String
            Dim Pubid As String
            Pubid = "-//VMWARE//DTD DITA "
            pubid1slash = "-/VMWARE/DTD DITA "
            Dim mapid1slash As String
            Dim Mapid As String
            Mapid = "-//OASIS//DTD DITA "
            mapid1slash = "-/OASIS/DTD DITA "
            Dim absURI As String = absoluteUri.ToString
            If (absoluteUri.ToString().Contains(Pubid) Or absoluteUri.ToString().Contains(Mapid) Or absoluteUri.ToString().Contains(pubid1slash) Or absoluteUri.ToString().Contains(mapid1slash)) Then
                Dim DTDpath As String
                DTDpath = Path.GetTempPath() & "nbsp.dtd"
                Return New FileStream(DTDpath, FileMode.Open, FileAccess.Read, FileShare.Read)
                ' Return New FileStream(DTDTopic, FileMode.Open, FileAccess.Read, FileShare.Read)


            ElseIf myHash.ContainsKey(absoluteUri) Then
                modErrorHandler.Errors.PrintMessage(1, "Reading resource" + absoluteUri.ToString() + " from cached stream", strModuleName)
                'Returns the cached stream.
                Return New FileStream(CType(myHash(absoluteUri), [String]), FileMode.Open, FileAccess.Read, FileShare.Read)
            Else
                Return MyBase.GetEntity(absoluteUri, role, ofObjectToReturn)

            End If
        End Function 'GetEntity
    End Class 'CustomResolver

    Private Shared Sub CopyDTDFile()
        Dim DTDpath As String
        DTDpath = Path.GetTempPath() & "nbsp.dtd"

        Dim clsResourceStream As System.IO.Stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("VMwareISHModulesNS.nbsp.dtd")
        Dim bResource As Byte() = DirectCast(Array.CreateInstance(GetType(Byte), clsResourceStream.Length), Byte())
        clsResourceStream.Read(bResource, 0, bResource.Length)
        clsResourceStream.Close()
        clsResourceStream.Dispose()
        Dim sResource As String = System.Text.Encoding.ASCII.GetString(bResource)
        Dim b As System.IO.TextWriter = New System.IO.StreamWriter(DTDpath, False, System.Text.Encoding.Unicode)
        b.Write(sResource)
        ' Clear object 
        b.Flush()
        b.Close()
    End Sub

    Private Function SaveTextToFile(ByVal strData As String, ByVal FullPath As String, Optional ByVal ErrInfo As String = "") As Boolean

        Dim bAns As Boolean = False
        Dim objReader As StreamWriter
        Try


            objReader = New StreamWriter(FullPath, False, Encoding.Unicode)
            objReader.Write(strData)
            objReader.Close()
            bAns = True
        Catch Ex As Exception
            ErrInfo = Ex.Message

        End Try
        Return bAns
    End Function 'SaveTextToFile
    ''' <summary>
    ''' Loads any FilePath specified XML file into an XML Document.  Works on all valid XML regardless of entities or DTD declarations.
    ''' </summary>
    ''' <returns>Returns a fully-formed XML Document object.</returns>
    Public Shared Function LoadFileIntoXMLDocument(ByVal FilePath As String) As XmlDocument
        Dim settings As XmlReaderSettings
        Dim resolver As New DITAResolver()
        settings = New XmlReaderSettings()
        settings.ProhibitDtd = False
        settings.ValidationType = ValidationType.None
        settings.XmlResolver = resolver
        settings.CloseInput = True
        Try
            Dim reader As XmlReader = XmlReader.Create(FilePath, settings)
            Dim doc As New XmlDocument
            doc.Load(reader)
            reader.Close()
            Return doc
        Catch ex As Exception
            modErrorHandler.Errors.PrintMessage(3, "Failed to load document as xml: " + FilePath + " ErrorMessage: " + ex.Message, strModuleName)
            Return Nothing
        End Try

    End Function
    ''' <summary>
    ''' Builds a template of standard metadata to be requested from the CMS.
    ''' </summary>
    ''' <returns>Stringbuilder of metadata for the CMS. </returns>
    Private Shared Function BuildRequestedMetadata() As StringBuilder
        Dim requestedmeta As New StringBuilder
        requestedmeta.Append("<ishfields>")
        requestedmeta.Append("<ishfield name=""FTITLE"" level=""logical""/>")
        requestedmeta.Append("<ishfield name=""VERSION"" level=""version""/>")
        'If Resolution = "" Then
        requestedmeta.Append("<ishfield name=""FAUTHOR"" level=""lng""/>")
        'Else
        requestedmeta.Append("<ishfield name=""FILLUSTRATOR"" level=""lng""/>")
        requestedmeta.Append("<ishfield name=""FRESOLUTION"" level=""lng""/>")
        'End If
        requestedmeta.Append("<ishfield name=""FSTATUS"" level=""lng""/>")
        requestedmeta.Append("<ishfield name=""DOC-LANGUAGE"" level=""lng""/>")
        'requestedmeta.Append("<ishfield name=""EDT-FILE-EXTENSION"" level=""lng""/>")
        requestedmeta.Append("</ishfields>")
        Return requestedmeta
    End Function


End Class


