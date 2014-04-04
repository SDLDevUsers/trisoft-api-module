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
    'Public Sub New()
    '    CopyDTDFile()
    '    'Load up our XML Templates for each type:
    '    oDocument.LoadXMLTemplates()

    'End Sub
    ''' <summary>
    ''' Creates the main ISH object to be used while interfacing with the LiveContent Architect system. 
    ''' TODO: Need to enable dynamic application switching between servers...
    ''' </summary>
    ''' <param name="Username"></param>
    ''' <param name="Password"></param>
    ''' <param name="ServerURL">Server URL.</param>
    ''' <remarks>Server URL should include FQDN including "WS". For instance: https://trisoft03.sdlproducts.com/InfoShareWS2 </remarks>
    Public Sub New(ByVal Username As String, ByVal Password As String, ByVal ServerURL As String)

        SetContext(Username, Password, ServerURL)
        oDocument.oCommonFuncs.LoadXMLTemplates()
        oDocument.oCommonFuncs.CopyDTDFile()
    End Sub






    Public oApplication As IshApplication
    Public oDocument As IshDocument
    Public oBaseline As IshBaseline
    Public oConditions As IshConditions
    Public oListOfValues As IshListOfValues
    Public oMeta As IshMeta
    Public oOutput As IshOutput
    Public oPub As IshPub
    Public oPubOutput As IshPubOutput
    Public oReports As IshReports
    Public oSearch As IshSearch
    Public oWorkflow As IshWorkflow
    Public oFolder As IshFolder

    Public oIshObjs As ISHObjs

    Public CMSServerURL As New String("")
    'Private Shared m_Context As New String("")


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


    Private ReadOnly Property ISHApp() As String
        Get
            ISHApp = "InfoShareAuthor"
        End Get
    End Property
    Private ReadOnly Property strModuleName() As String
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
            Context = oApplication.Context
        End Get
        Set(ByVal value As String)
            oApplication.Context = value
            oDocument.Context = value
            oBaseline.Context = value
            oConditions.Context = value
            oMeta.Context = value
            oOutput.Context = value
            oPub.Context = value
            oPubOutput.Context = value
            oReports.Context = value
            oSearch.context = value
            oWorkflow.Context = value
            oFolder.Context = value
        End Set
    End Property

    ''' <summary>
    ''' Used to set the context of the current user's credentials on the specified CMS URL.
    ''' </summary>
    Public Function SetContext(ByVal uname As String, ByVal passwd As String, ByVal RepositoryURL As String) As Boolean
        Try
        'First, test the App web reference with the passed URL.
        oApplication = New IshApplication(uname, passwd, RepositoryURL)
            CMSServerURL = RepositoryURL
        Catch ex As Exception
            modErrorHandler.Errors.PrintMessage(3, "Login failed: " + ex.Message.ToString, strModuleName)
            Return False
        End Try


            'Create all the new objects as previously defined but with the proper credentials:
            oApplication = New IshApplication(uname, passwd, CMSServerURL)
            oDocument = New IshDocument(uname, passwd, CMSServerURL)
            oBaseline = New IshBaseline(uname, passwd, CMSServerURL)
            oConditions = New IshConditions(uname, passwd, CMSServerURL)
        oListOfValues = New IshListOfValues(uname, passwd, CMSServerURL)
            oMeta = New IshMeta(uname, passwd, CMSServerURL)
            oOutput = New IshOutput(uname, passwd, CMSServerURL)
            oPub = New IshPub(uname, passwd, CMSServerURL)
            oPubOutput = New IshPubOutput(uname, passwd, CMSServerURL)
            oReports = New IshReports(uname, passwd, CMSServerURL)
            oSearch = New IshSearch(uname, passwd, CMSServerURL)
            oWorkflow = New IshWorkflow(uname, passwd, CMSServerURL)
            oFolder = New IshFolder(uname, passwd, CMSServerURL)


        ''Set all the ISHObj web refs to use this passed URL...
        'Try
        '    'Sets the context property which results in all the oISH objects' contexts getting set to the same context.
        '    Context = m_context
        'Catch ex As Exception
        '    modErrorHandler.Errors.PrintMessage(1, "Unable to set Web Reference URL.  Message: " + ex.Message, strModuleName)
        '    Return False
        'End Try




        Return True
    End Function
End Class



