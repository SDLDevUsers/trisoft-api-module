Imports System.Xml
Imports ErrorHandlerNS
Imports System.IO
Imports System.Text
Imports VMwareISHModulesNS.clsCommonFuncs

Public Class IshDocument
    Inherits mainAPIclass
#Region "Private Members"
    Private ReadOnly strModuleName As String = "IshDocument"

    Private m_recursion_hash As New Hashtable
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
    
    Public Function MoveObject(ByVal GUID As String, ByVal ToFolderID As Long) As Boolean
        'TODO: Has to search through the entire CMS Structure just to figure out the current folder ID of the specified GUID...  Terribly inefficient!
        'The problem is that FolderID is not tracked with objects as part of its metadata.  Looking up an object gives you NO information about where it exists in the CMS.
        Try
            Dim CurrentFolder As Long = FindFolderIDforObjbyGUID(GUID, 0)
            If CurrentFolder > 0 Then
                oISHAPIObjs.ISHDocObj.Move(Context, GUID, CurrentFolder.ToString, ToFolderID.ToString)
            Else
                modErrorHandler.Errors.PrintMessage(3, "Unable to move object " + GUID + " to specified location!", strModuleName + "-MoveObject")
            End If
            Return True
        Catch ex As Exception
            modErrorHandler.Errors.PrintMessage(3, "Unable to move object " + GUID + " to specified location!", strModuleName + "-MoveObject")
            Return False
        End Try
    End Function

    Public Function ChangeState(ByVal strDesiredState As String, ByVal strGUID As String, ByVal strVer As String, ByVal strLanguage As String, ByVal strResolution As String) As Boolean
        Dim myCurrentState = GetCurrentState(strGUID, strVer, strResolution, strLanguage)
        Dim processingresult As Boolean = True
        Dim strMetaState As String = "<ishfields><ishfield name=""FSTATUS"" level=""lng"">" + strDesiredState + "</ishfield></ishfields>"
        ''Could be used to update person assigned to a specific role as well...
        'Dim strMetaRole As String = "<ishfields><ishfield name=""FEDITOR"" level=""lng"">" + strEditorName + "</ishfield></ishfields>"
        Dim result As Boolean = True

        ' Generic move status drive used to change the status of a topic and COULD be used to update the name associated with the status
        If (Context = "") Then
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

    Public Function CheckIn(ByVal PathToCheckInFile As String) As Boolean
        Dim checkinfile As New FileInfo(PathToCheckInFile)
        Dim CMSFilename, GUID, Version, Language, Resolution As New String("")
        oCommonFuncs.GetCommonMetaFromLocalFile(checkinfile.FullName, CMSFilename, GUID, Version, Language, Resolution)
        Dim checkinblob As Byte() = oCommonFuncs.GetIshBlobFromFile(checkinfile.FullName)
        Try
            oISHAPIObjs.ISHDocObj.CheckIn(Context, GUID, Version, Language, Resolution, "", oCommonFuncs.GetISHEdt(checkinfile.Extension), checkinblob)
            modErrorHandler.Errors.PrintMessage(1, "Checked in object " + checkinfile.FullName + ".", strModuleName + "-CheckIn")
            checkinfile.Attributes = FileAttributes.Normal
            checkinfile.Delete()
            Return True
        Catch ex As Exception
            modErrorHandler.Errors.PrintMessage(3, "Failed to check in object " + checkinfile.FullName + ". Message: " + ex.Message, strModuleName + "-CheckIn")
            Return False
        End Try
    End Function
    Public Function CheckOut(ByVal GUID As String, ByVal Version As String, ByVal Language As String, ByVal Resolution As String, ByVal LocalStorePath As String) As Boolean
        Dim CheckOutFile As New String("")
        'first, ensure it exists.
        If ObjectExists(GUID, Version, Language, Resolution) Then
            'Check out the object
            Try
                oISHAPIObjs.ISHDocObj.CheckOut(Context, GUID, Version, Language, Resolution, "", CheckOutFile)
                modErrorHandler.Errors.PrintMessage(1, "Checked out object " + GUID + "=" + Version + "=" + Language + "=" + Resolution + ".", strModuleName + "-CheckOut")
                CheckOutFile = ""
            Catch ex As Exception
                If GetCurrentState(GUID, Version, Resolution, Language) = "Released" Then
                    CreateNewVersion(GUID, Version, Language, Resolution)
                    'now try checking out the new version:
                    Try
                        oISHAPIObjs.ISHDocObj.CheckOut(Context, GUID, Version, Language, Resolution, "", CheckOutFile)
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
    Public Function CreateNewVersion(ByVal GUID As String, ByRef Version As String, ByVal Language As String, Optional ByVal Resolution As String = "") As Boolean
        'Specified version is released so we need to create a new version of the specified branch.
        'Get the existing content at the current version.
        Dim newverDoc As XmlDocument = GetObjByID(GUID, Version, Language, Resolution)
        'if the content is a map or topic, we need to remove the processing instruction before it can be added as a new version.
        Dim datablob As Byte()
        Dim newverIshType As String = oCommonFuncs.GetISHTypeFromMeta(newverDoc)
        If newverIshType = "ISHModule" Or newverIshType = "ISHMasterDoc" Then
            Dim MyNode As XmlNode = newverDoc.SelectSingleNode("//ishdata")
            'get the dita topic out of the CData:
            Dim DITATopic As XmlDocument = oCommonFuncs.GetXMLOut(MyNode)
            'drop the ISH version specific ProcInstr
            Dim ishnode As XmlNode = DITATopic.SelectSingleNode("/processing-instruction('ish')")
            DITATopic.RemoveChild(ishnode)
            DITATopic.Save("c:\temp\deletetopic.xml")
            'load the doc to a datablob:
            'Convert the doc to an ISH blob
            datablob = oCommonFuncs.GetIshBlobFromFile("c:\temp\deletetopic.xml")
            'delete local file
            File.Delete("c:\temp\deletetopic.xml")
        Else
            'get the blob (images only) needed to create the new version:
            datablob = oCommonFuncs.GetBinaryOut(newverDoc.SelectSingleNode("//ishdata"))
        End If

        'get the various required parameters needed to create the new version:
        Dim IshType As DocumentObj20.eISHType = StringToISHType(oCommonFuncs.GetISHTypeFromMeta(newverDoc))
        Dim basefolder As DocumentObj25.eBaseFolder
        Dim folderpath() As String
        Dim folderID() As Long
        oISHAPIObjs.ISHDocObj25.FolderLocation(Context, GUID, basefolder, folderpath, folderID)
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
                oISHAPIObjs.ISHDocObj.CreateOrUpdate(Context, folderID(folderID.Length - 1), IshType, GUID, Version, Language, Resolution, "Draft", XMLMetaData, XMLMetaData, psEDT, datablob)
            Catch ex2 As Exception
                modErrorHandler.Errors.PrintMessage(3, "Failed to create new version of object " + GUID + "=" + Version + "=" + Language + "=" + Resolution + ". Message: " + ex2.Message, strModuleName + "-CheckOut")
                Return False
            End Try
        Else
            Try
                'Create the new version AND language:
                oISHAPIObjs.ISHDocObj.CreateOrUpdate(Context, folderID(folderID.Length - 1), IshType, GUID, "new", Language, Resolution, "Draft", XMLMetaData, XMLMetaData, psEDT, datablob)
            Catch ex2 As Exception
                modErrorHandler.Errors.PrintMessage(3, "Failed to create new version of object " + GUID + "=" + Version + "=" + Language + "=" + Resolution + ". Message: " + ex2.Message, strModuleName + "-CheckOut")
                Return False
            End Try
        End If

    End Function
    Public Function GetLatestVersionNumber(ByVal GUID As String) As String
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
    ''' Imports specified file using the parameters provided.  Returns true if successful.
    ''' </summary>
    ''' <param name="FilePath">Path to the file to be imported.</param>
    ''' <param name="CMSFolderID">ID of the CMS folder to import to.</param>
    ''' <param name="Author">Author of the object being imported.</param>
    ''' <param name="ISHType">Type of content being imported.  Allowed values are: "ISHIllustration", "ISHLibrary", "ISHMasterDoc", "ISHModule", "ISHNone", "ISHPublication", "ISHReusedObj", and "ISHTemplate"</param>
    ''' <param name="ReturnedGUID">Returns GUID set by CMS upon successful import.</param>
    ''' <param name="CMSTitle">Title to be used for the object within the CMS. This is shown to users of the CMS and can be updated in the object Properties.</param>
    ''' <param name="ObjectMetaType">Specifies the value found in the LOV for a module or image (Graphic, Icon, Screenshot, Concept, Reference, Task, etc.).  Value is arbitrary but is only properly set if found in the CMS.</param>
    Public Function ImportObject(ByVal FilePath As String, ByVal CMSFolderID As Long, ByVal Author As String, ByVal ISHType As String, ByVal strState As String, ByRef ReturnedGUID As String, ByRef CMSTitle As String, Optional ByVal ObjectMetaType As String = "") As Boolean
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
            If oCommonFuncs.GetCommonMetaFromLocalFile(FilePath, CMSFileName, CMSTitle, GUID, Version, Language, Resolution) = False Then
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
            metaxml = oCommonFuncs.GetMetaDataXMLStucture(CMSTitle, Version, Author, strState, Resolution, Language, ObjectMetaType)
            If metaxml = "" Then
                modErrorHandler.Errors.PrintMessage(3, "Failed to generate the xml metadata needed to create the content in the CMS. Aborting import.", strModuleName + "importobj-getmeta")
                Return False
            End If
            'now that we have the meta, need to get the bytearray data blob
            Dim data As Byte()
            data = oCommonFuncs.GetIshBlobFromFile(FilePath)


            Dim result As String = ""
            ' Import the content if it doesn't already exist in the CMS
            If ObjectExists(GUID, Version, Language, Resolution) = False And Language = "en" Then
                Try
                    oISHAPIObjs.ISHDocObj.Create(Context, CMSFolderID.ToString, StringToISHType(ISHType), GUID, Version, Language, Resolution, metaxml, oCommonFuncs.GetISHEdt(Path.GetExtension(FilePath)), data)
                    ReturnedGUID = GUID
                    ''if objectmetatype is icon, also import thumbnail as new resolution
                    'If ObjectMetaType = "Icon" Then
                    '    ISHDocObj.Create(Context, CMSFolderID.ToString, StringToISHType(ISHType), ReturnedGUID, Version, Language, "Thumbnail", metaxml, GetISHEdt(Path.GetExtension(FilePath)), data)
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
                ReturnedGUID = GUID
                'May be a localization import.  Check the language code to see if it's 'en'
                If Language = "en" Then
                    'Is an EN object but it exists in the DB, so report it as already imported.
                    modErrorHandler.Errors.PrintMessage(2, "Object to be imported already exists in the CMS. Skipping import for " + FilePath, strModuleName + "-ImportObject")

                    Return True ' It's already imported, let the user know.
                Else
                    'This is a localized file that needs imported.
                    'Get the current state of the file being imported.
                    Dim currentstate As String = GetCurrentState(GUID, Version, Resolution, Language)
                    'if it is not "In Translation", change it so that it can be imported.
                    If Not currentstate = "In Translation" Then
                        Dim strMetaState As String = "<ishfields><ishfield name=""FSTATUS"" level=""lng"">In Translation</ishfield></ishfields>"
                        If SetMeta(strMetaState, GUID, Version, Resolution, Language) = False Then
                            modErrorHandler.Errors.PrintMessage(3, "Failed to change object state to 'In Translation' to allow reimport. File: " + FilePath, strModuleName + "importobj-updatel10n")
                            Return False
                        End If
                    End If
                    'Now we should be able to import it...

                    Dim currentmeta As String = ""
                    'Get the current metadata to allow the update
                    oISHAPIObjs.ISHDocObj.GetMetaData(Context, GUID, Version, Language, Resolution, "<ishfields><ishfield name=""FTITLE"" level=""logical""/><ishfield name=""FSTATUS"" level=""lng""/></ishfields>", currentmeta)
                    Try
                        'attempt to update the current content with the new content and change the state to "Translated":
                        oISHAPIObjs.ISHDocObj.Update(Context, GUID, Version, Language, Resolution, metaxml, currentmeta, oCommonFuncs.GetISHEdt(Path.GetExtension(FilePath)), data)
                        Return True
                    Catch ex As Exception
                        modErrorHandler.Errors.PrintMessage(3, "Failed to import a file to the CMS. File: " + FilePath + ". Error Message: " + ex.Message, strModuleName + "importobj-updatel10n")
                        Return False
                    End Try
                End If
            End If
        Else
            Return False 'File didn't exist locally...
        End If
    End Function


    ''' <summary>
    ''' Converts a string (ISHIllustration, ISHBaseline, etc.) to a valid ISHType object.
    ''' </summary>
    Public Function StringToISHType(ByVal IshType As String) As VMwareISHModulesNS.DocumentObj20.eISHType
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


    Public Function _ResetRecursionHash()
        m_recursion_hash.Clear()
        Return True
    End Function

    Public Function ChangeRoleRecursivelybyBaseline(ByVal PubGUID As String, ByVal PubVer As String, ByVal NewPerson As String, ByVal Role As String, Optional ByVal SubGUID As String = "", Optional ByVal Resolution As String = "") As Boolean
        'if we were provided a SubGUID, that should be our startingpoint.  otherwise, our starting point is the mastermap.
        Dim GUID As String = ""
        If SubGUID.Length > 0 Then
            GUID = SubGUID
        Else
            GUID = GetMasterMapGUID(PubGUID, PubVer)
        End If
        'Get the baseline dictionary
        Dim dictBaseLineInfo As Dictionary(Of String, CMSObject)
        dictBaseLineInfo = GetBaselineObjects(PubGUID, PubVer)
        ChangeRecursiveSubRoutine(dictBaseLineInfo, GUID, NewPerson, Role, Resolution)
        Return True
    End Function
    Private Function ChangeRecursiveSubRoutine(ByRef dictBaseLineInfo As Dictionary(Of String, CMSObject), ByVal GUID As String, ByVal NewAuthor As String, ByVal Role As String, Optional ByVal Resolution As String = "") As Boolean
        'Dim requestmetadata As StringBuilder = BuildRequestedMetadata()
        Dim RequestedXMLObject As String = ""
        Dim doc As New XmlDocument
        Dim docorig As New XmlDocument

        'Version is determined by the baseline.
        Dim version As String
        Dim objCMSObject As New CMSObject("", "", "")
        dictBaseLineInfo.TryGetValue(GUID, objCMSObject)
        If objCMSObject.Version.Length > 0 Then
            version = objCMSObject.Version
        Else
            'failed to find object in the baseline
            modErrorHandler.Errors.PrintMessage(2, "Failed to find object in baseline. Info: " + GUID + ".", strModuleName + "-ChangeRecursiveSubRoutine")
            Return False
        End If
        Try
            oISHAPIObjs.ISHDocObj.GetMetaData(Context, GUID, version, "en", Resolution, oCommonFuncs.BuildRequestedMetadata().ToString, RequestedXMLObject)
        Catch ex As Exception
            'failed to get object
            modErrorHandler.Errors.PrintMessage(2, "Failed to get object in DB. Info: " + GUID + "=" + version + ". Message: " + ex.Message.ToString, strModuleName + "-ChangeRecursiveSubRoutine")
            Return False
        End Try

        'Load the XML and get the metadata:
        doc.LoadXml(RequestedXMLObject)

        'keep the original for matching later.
        docorig.LoadXml(RequestedXMLObject)
        Dim IshType As String = oCommonFuncs.GetISHTypeFromMeta(doc)

        'Get the children and recurse if applicable (by type).
        Select Case IshType
            Case "ISHMasterDoc", "ISHModule"
                'if a map or topic, get children
                'Dim CurMeta As Object = IshReports.GetReportedObjects(doc)
                Dim children As New Hashtable
                GetReferencedModules(GUID, children, version, "en")
                For Each childmodule As DictionaryEntry In children
                    If Not childmodule.Value.GUID = GUID And Not m_recursion_hash.Contains(childmodule.Value.GUID) Then
                        'Track the GUID we're diving into so that we don't accidently try to process it again later (endless looping).
                        m_recursion_hash.Add(childmodule.Value.GUID, childmodule.Value.GUID)
                        ChangeRecursiveSubRoutine(dictBaseLineInfo, childmodule.Value.GUID, NewAuthor, Role, childmodule.Value.Resolution)
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
                        oISHAPIObjs.ISHDocObj.SetMetaData(Context, GUID, version, "en", "", ishfields.OuterXml.ToString, RequestedXMLObject)
                    Catch ex As Exception
                        modErrorHandler.Errors.PrintMessage(2, "Unable to change assignee on object " + GUID + "=" + version + ". Message: " + ex.Message.ToString, strModuleName + "-ChangeRecursiveSubRoutine")
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
                        oISHAPIObjs.ISHDocObj.SetMetaData(Context, GUID, version, "en", "High", ishfields.OuterXml.ToString, ishfieldsorig.OuterXml.ToString)
                    Catch ex As Exception

                    End Try
                    res.InnerText = "Low"
                    resOrig.InnerText = "Low"
                    Try
                        oISHAPIObjs.ISHDocObj.SetMetaData(Context, GUID, version, "en", "Low", ishfields.OuterXml.ToString, ishfieldsorig.OuterXml.ToString)
                    Catch ex As Exception

                    End Try
                    res.InnerText = "Thumbnail"
                    resOrig.InnerText = "Thumbnail"
                    Try
                        oISHAPIObjs.ISHDocObj.SetMetaData(Context, GUID, version, "en", "Thumbnail", ishfields.OuterXml.ToString, ishfieldsorig.OuterXml.ToString)
                    Catch ex As Exception

                    End Try
                    res.InnerText = "Source"
                    resOrig.InnerText = "Source"
                    Try
                        oISHAPIObjs.ISHDocObj.SetMetaData(Context, GUID, version, "en", "Source", ishfields.OuterXml.ToString, ishfieldsorig.OuterXml.ToString)
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
    Public Function ChangeAssigneeRecursively(ByVal GUID As String, ByVal Version As String, ByVal NewAuthor As String, ByVal Role As String, Optional ByVal Resolution As String = "") As Boolean
        'Dim requestmetadata As StringBuilder = BuildRequestedMetadata()
        Dim RequestedXMLObject As String = ""
        Dim doc As New XmlDocument
        Dim docorig As New XmlDocument
        Try
            oISHAPIObjs.ISHDocObj.GetMetaData(Context, GUID, Version, "en", Resolution, oCommonFuncs.BuildRequestedMetadata().ToString, RequestedXMLObject)
        Catch ex As Exception
            'failed to get object
            modErrorHandler.Errors.PrintMessage(2, "Failed to get object in DB. Info: " + GUID + "=" + Version + ". Message: " + ex.Message.ToString, strModuleName + "-ChangeAssigneeRecursively")
            Return False
        End Try

        'Load the XML and get the metadata:
        doc.LoadXml(RequestedXMLObject)

        'keep the original for matching later.
        docorig.LoadXml(RequestedXMLObject)
        Dim IshType As String = oCommonFuncs.GetISHTypeFromMeta(doc)

        'Get the children and recurse if applicable (by type).
        Select Case IshType
            Case "ISHMasterDoc", "ISHModule"
                'if a map or topic, get children
                'Dim CurMeta As Object = IshReports.GetReportedObjects(doc)
                Dim children As New Hashtable
                GetReferencedModules(GUID, children, Version, "en")
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
                        oISHAPIObjs.ISHDocObj.SetMetaData(Context, GUID, Version, "en", "", ishfields.OuterXml.ToString, RequestedXMLObject)
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
                        oISHAPIObjs.ISHDocObj.SetMetaData(Context, GUID, Version, "en", "High", ishfields.OuterXml.ToString, ishfieldsorig.OuterXml.ToString)
                    Catch ex As Exception

                    End Try
                    res.InnerText = "Low"
                    resOrig.InnerText = "Low"
                    Try
                        oISHAPIObjs.ISHDocObj.SetMetaData(Context, GUID, Version, "en", "Low", ishfields.OuterXml.ToString, ishfieldsorig.OuterXml.ToString)
                    Catch ex As Exception

                    End Try
                    res.InnerText = "Thumbnail"
                    resOrig.InnerText = "Thumbnail"
                    Try
                        oISHAPIObjs.ISHDocObj.SetMetaData(Context, GUID, Version, "en", "Thumbnail", ishfields.OuterXml.ToString, ishfieldsorig.OuterXml.ToString)
                    Catch ex As Exception

                    End Try
                    res.InnerText = "Source"
                    resOrig.InnerText = "Source"
                    Try
                        oISHAPIObjs.ISHDocObj.SetMetaData(Context, GUID, Version, "en", "Source", ishfields.OuterXml.ToString, ishfieldsorig.OuterXml.ToString)
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

    Public Function CanMoveToState(ByVal strState As String, ByVal strGUID As String, ByVal strVersion As String, ByVal strResolution As String, Optional ByVal strLanguage As String = "en") As Boolean

        Dim OutStates As String()
        Try
            ' Declare variable for the Application service
            'Dim DocService As ISDoc.DocumentObj20 = New ISDoc.DocumentObj20()

            ' Clear variable for the result
            oISHAPIObjs.ISHDocObj.GetPossibleTransitionStates(Context, strGUID, strVersion, strLanguage, strResolution, OutStates)

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

    
    Public Function GetCurrentState(ByVal GUID As String, ByVal Version As String, ByVal strResolution As String, Optional ByVal Language As String = "en") As String

        Dim state As String = "nothing"
        Dim OutXML As String = ""
        Try
            '' Declare variable for the Application service
            'Dim DocService As IshDocument.DocumentObj20 = New ISDoc.DocumentObj20()

            Dim strMeta As String = "<ishfields><ishfield name=""FSTATUS"" level=""lng""/></ishfields>"

            oISHAPIObjs.ISHDocObj.GetMetaData(Context, GUID, Version, Language, strResolution, strMeta, OutXML)

        Catch ex As Exception
            modErrorHandler.Errors.PrintMessage(2, "Error getting current state for " + GUID + "" + Version + "" + Language + ". Message: " + ex.Message.ToString, strModuleName + "-GetCurrentState")
            Return False
        End Try

        Dim strFind As String = "<ishfield name=""FSTATUS"" level=""lng"">"
        state = OutXML.Substring(OutXML.LastIndexOf(strFind) + strFind.Length)
        state = state.Remove(state.LastIndexOf("</ishfield>"))
        Return state
    End Function
#End Region
End Class
