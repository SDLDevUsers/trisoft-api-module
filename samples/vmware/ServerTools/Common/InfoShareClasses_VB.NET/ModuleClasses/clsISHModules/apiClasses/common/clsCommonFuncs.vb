Imports System.Xml
Imports System.Text
Imports ErrorHandlerNS
Imports System.IO


Public Class clsCommonFuncs
    Private ReadOnly strModuleName As String = "CommonFuncs"
    Public Structure XMLTemplateStruct
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
        Dim library As XmlDocument
        Dim libraryblob() As Byte
    End Structure
    Public XMLTemplates As New XMLTemplateStruct

    ''' <summary>
    ''' Given an XML Document type, returns a base-64 encoded blob that can be fed directly to the CMS.
    ''' </summary>
    Public Function GetIshBlobFromXMLDoc(ByVal Doc As XmlDocument) As Byte()
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
            modErrorHandler.Errors.PrintMessage(3, "Failed to convert content to Base64 blob: " + ex.Message, strModuleName + "-GetIshBlobFromXMLDoc")
            Return Nothing
        End Try
    End Function
    Public Sub LoadXMLTemplates()
        Dim TemplateHash As New Hashtable
        TemplateHash.Add("map.ditamap", My.Application.Info.DirectoryPath + "\templateModules\map.ditamap")
        TemplateHash.Add("concept.xml", My.Application.Info.DirectoryPath + "\templateModules\concept.xml")
        TemplateHash.Add("task.xml", My.Application.Info.DirectoryPath + "\templateModules\task.xml")
        TemplateHash.Add("reference.xml", My.Application.Info.DirectoryPath + "\templateModules\reference.xml")
        TemplateHash.Add("troubleshooting.xml", My.Application.Info.DirectoryPath + "\templateModules\troubleshooting.xml")
        TemplateHash.Add("library-template.xml", My.Application.Info.DirectoryPath + "\templateModules\library-template.xml")
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
                        Case "library-template.xml"
                            .library = doc
                            .libraryblob = GetIshBlobFromXMLDoc(doc)
                    End Select
                End With
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(3, "Unable to load template files into memory! Check that they exist in " + My.Application.Info.DirectoryPath + "\templateModules" + ". Message: " + ex.Message, strModuleName + "-LoadXMLTemplates")
            End Try
        Next

    End Sub
    Public Function SetGUIDinTemplates(ByVal GUID As String) As Boolean
        Try
            With XMLTemplates
                .map.DocumentElement.Attributes.GetNamedItem("id").Value = GUID
                .concept.DocumentElement.Attributes.GetNamedItem("id").Value = GUID
                .task.DocumentElement.Attributes.GetNamedItem("id").Value = GUID
                .reference.DocumentElement.Attributes.GetNamedItem("id").Value = GUID
                .troubleshooting.DocumentElement.Attributes.GetNamedItem("id").Value = GUID
                .library.DocumentElement.Attributes.GetNamedItem("id").Value = GUID
            End With
        Catch ex As Exception
            modErrorHandler.Errors.PrintMessage(2, "Unable to set GUID in XML templates. Message: " + ex.Message, strModuleName)
        End Try

    End Function

    Public Function GetISHEdt(ByVal FileExtension As String) As String
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
    Public Function GetTopicTypeFromMeta(ByVal doc As XmlDocument) As String
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

    Public Function GetBinaryOut(ByVal DITANode As XmlNode) As Byte()
        Dim CDataNode As XmlNode
        Dim CData As String = ""
        Dim decodedBytes As Byte()

        Dim settings As XmlReaderSettings
        Dim resolver As New DITAResolver()
        settings = New XmlReaderSettings()
        settings.DtdProcessing = DtdProcessing.Parse 'False
        settings.ValidationType = ValidationType.None
        settings.XmlResolver = resolver
        settings.CloseInput = True
        settings.IgnoreWhitespace = False


        CDataNode = DITANode.FirstChild
        CData = CDataNode.InnerText
        decodedBytes = Convert.FromBase64String(CData)
        Return decodedBytes
    End Function
    Public Function GetISHTypeFromMeta(ByVal doc As XmlDocument) As String
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
    Public Function GetXMLOut(ByVal DITANode As XmlNode, Optional ByVal KeepBom As Boolean = False) As XmlDocument
        Dim CDataNode As XmlNode
        Dim CData As String = ""
        Dim decodedBytes As Byte()
        Dim decodedText As String

        Dim settings As XmlReaderSettings
        Dim resolver As New DITAResolver()
        settings = New XmlReaderSettings()
        settings.DtdProcessing = DtdProcessing.Parse 'False
        settings.ValidationType = ValidationType.None
        settings.XmlResolver = resolver
        settings.CloseInput = True
        settings.IgnoreWhitespace = False


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
        doc.PreserveWhitespace = True
        doc.Load(reader)

        Return doc
    End Function

    ''' <summary>
    ''' Given a path to a file of any type, returns a base-64 encoded blob that can be fed directly to the CMS.
    ''' </summary>
    Public Function GetIshBlobFromFile(ByVal FilePath As String) As Byte()
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
            modErrorHandler.Errors.PrintMessage(3, FilePath + ": Failed to convert content to Base64 blob. Message: " + ex.Message, strModuleName + "-GetIshBlobFromFile")
            Return Nothing
        End Try
    End Function
    Public Function GetFilenameFromIshMeta(ByVal IshMetaData As XmlDocument) As String
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

    Public Function GetMetaDataXMLStucture(ByVal CMSTitle As String, ByVal Version As String, ByVal Author As String, ByVal Status As String, ByVal Resolution As String, ByVal Language As String, Optional ByVal ModuleType As String = "", Optional ByVal Illustrator As String = "mmatus") As String
        'If CMSTitle = "" Or Version = "" Or Author = "" Or Status = "" Then
        '    'if one or more required fields are blank, abort opperation!
        '    modErrorHandler.Errors.PrintMessage(3, "One or more required Metadata fields are blank. Check the Author, Status, CMSTitle, and Version values.", strModuleName + "-GetMetadataXMLStructure")
        '    Return ""
        'End If
        Dim XMLString As New StringBuilder
        XMLString.Append("<ishfields>")
        If CMSTitle.Length > 0 And Language = "en" Then
            XMLString.Append("<ishfield name=""FTITLE"" level=""logical"">")
            XMLString.Append(CMSTitle)
            XMLString.Append("</ishfield>")
        End If
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
            XMLString.Append("<ishfield type=""hidden"" name=""FNOTRANSLATIONMGMT"" level=""logical"" label=""Disable translation management"">No</ishfield>")
        End If
        If Author.Length > 0 Then
            XMLString.Append("<ishfield name=""FAUTHOR"" level=""lng"">")
            XMLString.Append(Author)
            XMLString.Append("</ishfield>")
        End If
        If Status.Length > 0 Then
            XMLString.Append("<ishfield name=""FSTATUS"" level=""lng"">")
            XMLString.Append(Status)
            XMLString.Append("</ishfield>")
        End If
        XMLString.Append("</ishfields>")
        Return XMLString.ToString
    End Function
    ''' <summary>
    ''' Retrieves CMS metadata from a local file including CMSFilename for XML files.  File must be exported from the CMS or preprocessed for the CMS.
    ''' </summary>
    Public Function GetCommonMetaFromLocalFile(ByVal LocalFilePath As String, ByRef CMSFileName As String, ByRef GUID As String, ByRef Version As String, ByRef Language As String, ByRef Resolution As String) As Boolean
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
    Public Function GetCommonMetaFromLocalFile(ByVal LocalFilePath As String, ByRef CMSFileName As String, ByRef CMSTitle As String, ByRef GUID As String, ByRef Version As String, ByRef Language As String, ByRef Resolution As String) As Boolean
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
                    CMSTitle = CMSTitle.Replace("\", "")
                    CMSTitle = CMSTitle.Replace("/", "")
                    CMSTitle = CMSTitle.Replace(":", "")
                    CMSTitle = CMSTitle.Replace("*", "")
                    CMSTitle = CMSTitle.Replace("?", "")
                    CMSTitle = CMSTitle.Replace("""", "")
                    CMSTitle = CMSTitle.Replace("<", "")
                    CMSTitle = CMSTitle.Replace(">", "")
                    CMSTitle = CMSTitle.Replace("|", "")
                    CMSTitle = CMSTitle.Replace(vbCrLf, " ")
                    CMSTitle = CMSTitle.Replace(vbCr, " ")
                    CMSTitle = CMSTitle.Replace(vbLf, " ")
                    CMSTitle = CMSTitle.Replace("  ", " ")
                    CMSTitle = CMSTitle.Replace("  ", " ")
                    CMSTitle = CMSTitle.Replace("  ", " ")
                    CMSTitle = CMSTitle.Replace("  ", " ")
                Case Else 'just assign the title as the filename portion.
                    CMSTitle = CMSFileName
            End Select

            Return True
        Catch ex As Exception
            modErrorHandler.Errors.PrintMessage(3, "Failed get metadata based on filename and/or content for: " + LocalFilePath + " Error: " + ex.Message, strModuleName)
            Return False
        End Try


    End Function
    Public Sub CopyDTDFile()
        Try
            Dim DTDpath As String
            DTDpath = Path.GetTempPath() & "nbsp.dtd"

            Dim clsResourceStream As System.IO.Stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("IshModulesNS.nbsp.dtd")
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
        Catch ex As Exception

        End Try

    End Sub

    Public Function RemoveWindowsIllegalChars(ByVal strFileorFolder As String) As String
        strFileorFolder = strFileorFolder.Replace("‚", "") ' Special 'lower, single quotation mark'.
        strFileorFolder = strFileorFolder.Replace(",", "")
        strFileorFolder = strFileorFolder.Replace("\", "")
        strFileorFolder = strFileorFolder.Replace("/", "")
        strFileorFolder = Replace(strFileorFolder, "*", "")
        strFileorFolder = Replace(strFileorFolder, "?", "")
        strFileorFolder = Replace(strFileorFolder, ">", "")
        strFileorFolder = Replace(strFileorFolder, "<", "")
        strFileorFolder = Replace(strFileorFolder, ":", "")
        strFileorFolder = Replace(strFileorFolder, "|", "")
        strFileorFolder = Replace(strFileorFolder, "#", "")
        strFileorFolder = Replace(strFileorFolder, "!", "")
        strFileorFolder = Replace(strFileorFolder, """", "")
        Return strFileorFolder
    End Function


    Public Function SaveTextToFile(ByVal strData As String, ByVal FullPath As String, Optional ByVal ErrInfo As String = "") As Boolean

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
    Public Function LoadFileIntoXMLDocument(ByVal FilePath As String) As XmlDocument
        Dim settings As XmlReaderSettings
        Dim resolver As New DITAResolver()
        settings = New XmlReaderSettings()
        settings.DtdProcessing = DtdProcessing.Parse 'False
        settings.ValidationType = ValidationType.None
        settings.XmlResolver = resolver
        settings.CloseInput = True
        settings.IgnoreWhitespace = False
        Try
            Dim reader As XmlReader = XmlReader.Create(FilePath, settings)
            Dim doc As New XmlDocument
            doc.PreserveWhitespace = True
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
    Public Function BuildRequestedMetadata() As StringBuilder
        Dim requestedmeta As New StringBuilder
        requestedmeta.Append("<ishfields>")
        requestedmeta.Append("<ishfield name=""FTITLE"" level=""logical""/>")
        requestedmeta.Append("<ishfield name=""VERSION"" level=""version""/>")
        'If Resolution = "" Then
        requestedmeta.Append("<ishfield name=""FAUTHOR"" level=""lng""/>")
        'Else
        requestedmeta.Append("<ishfield name=""FRESOLUTION"" level=""lng""/>")
        'End If
        requestedmeta.Append("<ishfield name=""FSTATUS"" level=""lng""/>")
        requestedmeta.Append("<ishfield name=""DOC-LANGUAGE"" level=""lng""/>")
        'requestedmeta.Append("<ishfield name=""EDT-FILE-EXTENSION"" level=""lng""/>")
        requestedmeta.Append("</ishfields>")
        Return requestedmeta
    End Function
    Public Function BuildFullPubMetadata() As StringBuilder
        Dim requestedmeta As New StringBuilder
        requestedmeta.Append("<ishfields>")
        requestedmeta.Append("<ishfield name=""VERSION"" level=""version""/>")
        requestedmeta.Append("<ishfield name=""FISHPUBSTATUS"" level=""lng""/>")
        requestedmeta.Append("<ishfield name=""FISHISRELEASED"" level=""version""/>")
        requestedmeta.Append("<ishfield name=""DOC-LANGUAGE"" level=""lng""/>")
        requestedmeta.Append("<ishfield name=""FISHOUTPUTFORMATREF"" level=""lng""/>")
        requestedmeta.Append("<ishfield name=""FTITLE"" level=""logical""/>")
        requestedmeta.Append("</ishfields>")
        Return requestedmeta
    End Function
    Public Function BuildMinPubMetadata() As StringBuilder
        Dim requestedmeta As New StringBuilder
        requestedmeta.Append("<ishfields>")
        requestedmeta.Append("<ishfield name=""VERSION"" level=""version""/>")
        requestedmeta.Append("<ishfield name=""FISHISRELEASED"" level=""version""/>")
        requestedmeta.Append("<ishfield name=""FTITLE"" level=""logical""/>")
        requestedmeta.Append("<ishfield name=""FISHMASTERREF"" level=""version""/>")
        requestedmeta.Append("<ishfield name=""FISHBASELINE"" level=""version""/>")
        requestedmeta.Append("</ishfields>")
        Return requestedmeta
    End Function
    Public Function BuildPubMetaDataFilter(ByVal GUID As String, Optional ByVal Version As String = "latest", Optional ByVal Language As String = "", Optional ByVal OutputType As String = "all") As StringBuilder
        Dim requestedmeta As New StringBuilder

        requestedmeta.Append("<ishfields>")


        If Version = "latest" Then
            'Get all versions. Returns multiple ishobjects.
        Else
            'Get the specified version only.
            requestedmeta.Append("<ishfield name=""FMAPID"" level=""lng"">" + GUID + "</ishfield>")
        End If

        If OutputType = "all" Then
            'Get all versions. By not including the filter, all outputs will be returned.
        Else
            'Get the specified output type only
            requestedmeta.Append("<ishfield name=""FISHOUTPUTFORMATREF"" level=""lng"">" + OutputType + "</ishfield>")
        End If

        If Language.Length > 0 Then
            'A specific language is wanted. Return only outputs that are of that language.  otherwise, all languages are returned.
            requestedmeta.Append("<ishfield name=""DOC-LANGUAGE"" level=""lng"">" + Language + "</ishfield>")
        End If
        requestedmeta.Append("<ishfield name=""VERSION"" level=""version"">" + Version + "</ishfield>")
        requestedmeta.Append("</ishfields>")
        Return requestedmeta
    End Function

    Public Class DITAResolver
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
            Pubid = "-//MYCOMPANY//DTD DITA "
            pubid1slash = "-/MYCOMPANY/DTD DITA "
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
End Class
