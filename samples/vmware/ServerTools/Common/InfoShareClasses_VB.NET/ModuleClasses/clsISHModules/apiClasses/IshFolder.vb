Imports System.Xml
Imports ErrorHandlerNS
Imports System.Text
Imports System.IO

Public Class IshFolder
    Inherits mainAPIclass
#Region "Private Members"
    Private ReadOnly strModuleName As String = "ISHDocument"
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
    ''' Creates a new folder as a sub-folder to the given ParentFolderID and returns the newly created ID.  If a folder of the same name already exists, the function returns the ID of that folder instead.
    ''' </summary>
    ''' <param name="ParentFolderID">Containing folder's ID.</param>
    ''' <param name="ISHType">Folder type: ISHNone, ISHMasterDoc, ISHModule, ISHIllustration, ISHReusedObj, ISHTemplate, and ISHLibrary</param>
    ''' <param name="FolderName">Name of folder to create/reuse.</param>
    ''' <param name="NewFolderID">New/reused folder's ID.</param>
    ''' <param name="OwnershipGroup">Ownership group based on groups defined in the CMS.</param>
    ''' <param name="ReadAccessList">Comma-separated string with user groups that have read access to this folder.</param>
    Public Function CreateOrUseFolder(ByVal ParentFolderID As Long, ByVal ISHType As String, ByVal FolderName As String, ByRef NewFolderID As Long, Optional ByVal OwnershipGroup As String = "", Optional ByVal ReadAccess As String = "") As String
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
            CreateFolderResult = oISHAPIObjs.ISHFolderObj.Create(Context, ParentFolderID, RealIshType, FolderName, OwnershipGroup, NewFolderID, ReadAccess)
        End If

        Return CreateFolderResult
    End Function

    ''' <summary>
    ''' Function that returns the ID of a folder specified by a full CMS Path
    ''' Remember to include //Doc at the beginning.
    ''' </summary>
    ''' <param name="FullCMSPath"></param>
    Public Function GetFolderIDByPath(ByVal FullCMSPath As String, Optional ByRef Result As String = "", Optional ByRef OwnershipGroup As String = "", Optional ByRef ReadAccess As String = "") As Long
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
                oISHAPIObjs.ISHFolderObj.GetProperties(Context, rootfolderid, RealRootName, rootishtype, OutQuery, OwnershipGroup, ReadAccess)
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
    
    


    ''' <summary>
    ''' Function that returns the ID of a subfolder specified by its name
    ''' </summary>
    ''' <param name="ParentFOlderID">ID of the Parent Folder</param>
    ''' <param name="SubFolderName">Name of the subfolder to find</param>
    Public Function GetFolderIDByName(ByVal ParentFolderID As Long, ByVal SubFolderName As String, ByRef OwnershipGroup As String, ByRef ReadAccess As String) As Long
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

        'find the id of the one matching the name we're looking for 
        Dim subid As String = ""
        Dim foldername As String = ""
        Dim curishtype As VMwareISHModulesNS.Folder20.eISHFolderType
        Dim curoutquery As String = ""
        Try
            subid = doc.SelectSingleNode("//ishfolders/ishfolder[@name='" + SubFolderName + "']/@ishfolderref").Value.ToString
            oISHAPIObjs.ISHFolderObj.GetProperties(Context, Convert.ToInt64(subid), foldername, curishtype, curoutquery, OwnershipGroup, ReadAccess)
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
    Public Function GetSubFolderIDs(ByVal FolderID As Long) As ArrayList
        Dim subfolderIDs As New ArrayList
        'get a folderlist of all children of currentfolderid
        Dim subfolderlistXML As String = ""
        Dim CMSReply As String
        Try
            CMSReply = oISHAPIObjs.ISHFolderObj.GetSubFolders(Context, FolderID.ToString, 1, subfolderlistXML)
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
    Public Function FolderExists(ByVal ParentFolderID As Long, ByVal SubFolderName As String) As Boolean
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
    Public Function FolderHasContents(ByVal FolderID As Long) As Boolean
        Dim foldercontentXML As String = ""
        Dim CMSReply As String
        Try
            CMSReply = oISHAPIObjs.ISHFolderObj.GetContents(Context, FolderID.ToString, "", "", foldercontentXML)
        Catch ex As Exception

        End Try
        'Load the subfolderlistXML into an xml document
        ' Create the reader from the string.
        Dim strReader As New StringReader(foldercontentXML)

        Dim reader As XmlReader = XmlReader.Create(strReader)
        ' Create the new XMLdoc and load the content into it.
        Dim doc As New XmlDocument()
        doc.PreserveWhitespace = True
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
    Public Function FolderHasSubFolders(ByVal FolderID As Long)
        'get a folderlist of all children of currentfolderid
        Dim subfolderlistXML As String = ""
        Dim CMSReply As String
        Try
            CMSReply = oISHAPIObjs.ISHFolderObj.GetSubFolders(Context, FolderID.ToString, 1, subfolderlistXML)
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
    Public Function DeleteSubContent(ByVal FolderID As Long) As Boolean
        Dim RealRequestedMeta As New StringBuilder
        RealRequestedMeta = oCommonFuncs.BuildRequestedMetadata()
        Dim requestedmeta As String = RealRequestedMeta.ToString
        Dim foldercontentXML As String = ""
        Dim CMSReply As String
        Try
            CMSReply = oISHAPIObjs.ISHFolderObj.GetContents(Context, FolderID.ToString, "", requestedmeta, foldercontentXML)
        Catch ex As Exception

        End Try
        'Load the subcontentlistXML into an xml document
        ' Create the reader from the string.
        Dim strReader As New StringReader(foldercontentXML)

        Dim reader As XmlReader = XmlReader.Create(strReader)
        ' Create the new XMLdoc and load the content into it.
        Dim doc As New XmlDocument()
        doc.PreserveWhitespace = True
        doc.Load(reader)

        Dim significantchildren As XmlNodeList = doc.SelectNodes("//ishobject")
        'First, collect the info from each xml node.  Use the structure found in IshDocument
        If significantchildren.Count > 0 Then
            Dim ContentHash As New Hashtable
            ContentHash = GetReportedObjects(doc)
            For Each child As DictionaryEntry In ContentHash
                If DeleteObjectRecursivelyByGUID(child.Value.GUID, child.Value.Version, child.Value.Language, child.Value.Resolution, False) Then
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
    Public Function DeleteFolderRecursive(ByVal FolderID As Long, ByRef DelResult As String, Optional ByVal DeleteContentToo As Boolean = True) As Boolean
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
    Public Function DeleteFolder(ByVal folderID As Long, ByRef DelResult As String) As Boolean
        If FolderHasContents(folderID) = False Then
            Try
                oISHAPIObjs.ISHFolderObj.Delete(Context, folderID)
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

#End Region



End Class
