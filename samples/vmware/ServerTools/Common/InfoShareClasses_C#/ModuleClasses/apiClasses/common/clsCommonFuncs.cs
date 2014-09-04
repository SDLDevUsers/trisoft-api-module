using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Xml;
using System.Text;
using ErrorHandlerNS;
using System.IO;
namespace ISHModulesNS
{


	public class clsCommonFuncs
	{
		private readonly string strModuleName = "CommonFuncs";
		public struct XMLTemplateStruct
		{
			public XmlDocument map;
			public byte[] mapblob;
			public XmlDocument concept;
			public byte[] conceptblob;
			public XmlDocument task;
			public byte[] taskblob;
			public XmlDocument reference;
			public byte[] referenceblob;
			public XmlDocument troubleshooting;
			public byte[] troubleshootingblob;
			public XmlDocument library;
			public byte[] libraryblob;
		}

		public XMLTemplateStruct XMLTemplates = new XMLTemplateStruct();
		/// <summary>
		/// Given an XML Document type, returns a base-64 encoded blob that can be fed directly to the CMS.
		/// </summary>
		public byte[] GetIshBlobFromXMLDoc(XmlDocument Doc)
		{
			byte[] Data = null;
			try {
				long numBytes = Doc.OuterXml.Length;
				System.IO.MemoryStream myxmlstream = null;

				//Get the encoding of the XML Document so we know how to read it into a binary stream:
				XmlDeclaration decl = null;
				//assume UTF-8 encoding
				string xmlencoding = null;
				//find the real encoding:
				try {
					decl = (XmlDeclaration)Doc.FirstChild;
					//grab the encoding from the declaration
					xmlencoding = decl.Encoding;
				} catch (Exception fnf) {
					// If an error is caught, the encoding is not set.  This means we just go ahead with the UTF-8 encoding.
					// document might be missing an xml declaration
					xmlencoding = "utf-8";
					modErrorHandler.Errors.PrintMessage(2, "The XMLDocument threw an error when getting the encoding.  This xml document could be missing an XML declaration. Assuming UTF-8 to try anyway." + fnf.Message, strModuleName + "-GetIshBlobFromXMLDoc");
				}

				switch (xmlencoding.ToLower()) {
					case "utf-8":
						myxmlstream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(Doc.OuterXml));
						break;
					case "utf-16":
					case "Unicode":
					case "unicode":
						myxmlstream = new System.IO.MemoryStream(System.Text.Encoding.Unicode.GetBytes(Doc.OuterXml));
						break;
					case "utf-32":
						myxmlstream = new System.IO.MemoryStream(System.Text.Encoding.UTF32.GetBytes(Doc.OuterXml));
						break;
					default:
						//um. huh?  What encoding is this?
						return null;
				}
				//Read the stream into our binary reader (byte array) and return it.
				BinaryReader br = new BinaryReader(myxmlstream);
				Data = br.ReadBytes(Convert.ToInt32(numBytes));
				br.Close();
				myxmlstream.Close();
				return Data;
			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(3, "Failed to convert content to Base64 blob: " + ex.Message, strModuleName + "-GetIshBlobFromXMLDoc");
				return null;
			}
		}
		public void LoadXMLTemplates()
		{
			Hashtable TemplateHash = new Hashtable();
			TemplateHash.Add("map.ditamap", ISHModulesNS.My.MyProject.Application.Info.DirectoryPath + "\\templateModules\\map.ditamap");
			TemplateHash.Add("concept.xml", ISHModulesNS.My.MyProject.Application.Info.DirectoryPath + "\\templateModules\\concept.xml");
			TemplateHash.Add("task.xml", ISHModulesNS.My.MyProject.Application.Info.DirectoryPath + "\\templateModules\\task.xml");
			TemplateHash.Add("reference.xml", ISHModulesNS.My.MyProject.Application.Info.DirectoryPath + "\\templateModules\\reference.xml");
			TemplateHash.Add("troubleshooting.xml", ISHModulesNS.My.MyProject.Application.Info.DirectoryPath + "\\templateModules\\troubleshooting.xml");
			TemplateHash.Add("library-template.xml", ISHModulesNS.My.MyProject.Application.Info.DirectoryPath + "\\templateModules\\library-template.xml");
			foreach (DictionaryEntry templatefile in TemplateHash) {
				try {
					XmlDocument doc = LoadFileIntoXMLDocument(templatefile.Value);
					var _with1 = XMLTemplates;
					switch (templatefile.Key) {
						case "map.ditamap":
							_with1.map = doc;
							_with1.mapblob = GetIshBlobFromXMLDoc(doc);
							break;
						case "concept.xml":
							_with1.concept = doc;
							_with1.conceptblob = GetIshBlobFromXMLDoc(doc);
							break;
						case "reference.xml":
							_with1.reference = doc;
							_with1.referenceblob = GetIshBlobFromXMLDoc(doc);
							break;
						case "task.xml":
							_with1.task = doc;
							_with1.taskblob = GetIshBlobFromXMLDoc(doc);
							break;
						case "troubleshooting.xml":
							_with1.troubleshooting = doc;
							_with1.troubleshootingblob = GetIshBlobFromXMLDoc(doc);
							break;
						case "library-template.xml":
							_with1.library = doc;
							_with1.libraryblob = GetIshBlobFromXMLDoc(doc);
							break;
					}
				} catch (Exception ex) {
					modErrorHandler.Errors.PrintMessage(3, "Unable to load template files into memory! Check that they exist in " + ISHModulesNS.My.MyProject.Application.Info.DirectoryPath + "\\templateModules" + ". Message: " + ex.Message, strModuleName + "-LoadXMLTemplates");
				}
			}

		}
		public bool SetGUIDinTemplates(string GUID)
		{
			try {
				var _with2 = XMLTemplates;
				_with2.map.DocumentElement.Attributes.GetNamedItem("id").Value = GUID;
				_with2.concept.DocumentElement.Attributes.GetNamedItem("id").Value = GUID;
				_with2.task.DocumentElement.Attributes.GetNamedItem("id").Value = GUID;
				_with2.reference.DocumentElement.Attributes.GetNamedItem("id").Value = GUID;
				_with2.troubleshooting.DocumentElement.Attributes.GetNamedItem("id").Value = GUID;
				_with2.library.DocumentElement.Attributes.GetNamedItem("id").Value = GUID;
			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(2, "Unable to set GUID in XML templates. Message: " + ex.Message, strModuleName);
			}

		}

		public string GetISHEdt(string FileExtension)
		{
			//Valid EDT Values:
			//EDT-PDF
			//EDT-WORD
			//EDT-EXCEL
			//EDT-PPT
			//EDT-TEXT
			//EDT-HTML
			//EDT-REP-S3
			//EDT-TIFF
			//EDT-TRIDOC
			//EDTXML
			//EDTFM
			//EDTJPEG
			//EDTGIF
			//EDTUNDEFINED
			//EDTFLASH
			//EDTCGM
			//EDTBMP
			//EDTMPEG
			//EDTEPS
			//EDTAVI
			//EDTZIP
			//EDTPNG
			//EDTSVG
			//EDTSVGZ
			//EDTWMF
			//EDTRAR
			//EDTTAR
			//EDTHLP
			//EDTRTF
			//EDTHTM
			//EDTCSS
			//EDTPDF
			//EDTDOC
			//EDTXLS
			//EDTCHM
			//EDTVSD
			//EDTPSD
			//EDTPSP
			//EDTEMF
			//EDTAI

			switch (FileExtension.Replace(".", "").ToLower()) {
				case "xml":
				case "dita":
				case "ditamap":
					return "EDTXML";
				case "eps":
					return "EDTEPS";
				case "jpg":
				case "jpeg":
					return "EDTJPEG";
				case "fm":
				case "book":
				case "mif":
					return "EDTFM";
				case "xls":
					return "EDTXLS";
				case "chm":
					return "EDTCHM";
				case "VSD":
					return "EDTVSD";
				case "psd":
					return "EDTPSD";
				case "emf":
					return "EDTEMF";
				case "ai":
					return "EDTAI";
				case "tif":
				case "tiff":
					return "EDT-TIFF";
				case "doc":
					return "EDTDOC";
				case "pdf":
					return "EDTPDF";
				case "css":
					return "EDTCSS";
				case "htm":
				case "html":
					return "EDTHTM";
				case "RTF":
					return "EDTRTF";
				case "hlp":
					return "EDTHLP";
				case "tar":
					return "EDTTAR";
				case "rar":
					return "EDTRAR";
				case "wmf":
					return "EDTWMF";
				case "svgz":
					return "EDTSVGZ";
				case "svg":
					return "EDTSVG";
				case "png":
					return "EDTPNG";
				case "zip":
					return "EDTZIP";
				case "avi":
					return "EDTAVI";
				case "mpg":
				case "mpeg":
					return "EDTMPEG";
				case "gif":
					return "EDTGIF";
				case "fla":
				case "swf":
					return "EDTFLASH";
				case "bmp":
					return "EDTBMP";
				case "fla":
					return "EDTFLASH";
				case "pdf":
					return "EDTPDF";
				default:
					return "EDTUNDEFINED";
			}
		}
		public string GetTopicTypeFromMeta(XmlDocument doc)
		{
			XmlNode MyNode = null;
			XmlDocument DITATopic = new XmlDocument();
			string topictype = "";
			MyNode = doc.SelectSingleNode("//ishdata");
			//get the dita topic out of the CData:
			DITATopic = GetXMLOut(MyNode);
			topictype = DITATopic.DocumentElement.Name;
			//topictype = DITATopic.
			return topictype;
		}

		public byte[] GetBinaryOut(XmlNode DITANode)
		{
			XmlNode CDataNode = null;
			string CData = "";
			byte[] decodedBytes = null;

			XmlReaderSettings settings = null;
			DITAResolver resolver = new DITAResolver();
			settings = new XmlReaderSettings();
			settings.DtdProcessing = DtdProcessing.Parse;
			//False
			settings.ValidationType = ValidationType.None;
			settings.XmlResolver = resolver;
			settings.CloseInput = true;
			settings.IgnoreWhitespace = false;


			CDataNode = DITANode.FirstChild;
			CData = CDataNode.InnerText;
			decodedBytes = Convert.FromBase64String(CData);
			return decodedBytes;
		}
		public string GetISHTypeFromMeta(XmlDocument doc)
		{
			XmlNode MyNode = null;
			string ishtype = "";
			MyNode = doc.SelectSingleNode("//ishobject");
			//Get the ISHType:
			foreach (XmlAttribute ishattrib in MyNode.Attributes) {
				if (ishattrib.Name == "ishtype") {
					ishtype = ishattrib.Value;
				}
			}
			return ishtype;
		}
		public XmlDocument GetXMLOut(XmlNode DITANode, bool KeepBom = false)
		{
			XmlNode CDataNode = null;
			string CData = "";
			byte[] decodedBytes = null;
			string decodedText = null;

			XmlReaderSettings settings = null;
			DITAResolver resolver = new DITAResolver();
			settings = new XmlReaderSettings();
			settings.DtdProcessing = DtdProcessing.Parse;
			//False
			settings.ValidationType = ValidationType.None;
			settings.XmlResolver = resolver;
			settings.CloseInput = true;
			settings.IgnoreWhitespace = false;


			CDataNode = DITANode.FirstChild;
			CData = CDataNode.InnerText;
			decodedBytes = Convert.FromBase64String(CData);
			decodedText = Encoding.Unicode.GetString(decodedBytes);
			string strStripBOM = "";
			//We're stripping the BOM off here; Disable via optional parameter if you need to keep it.
			strStripBOM = decodedText.Substring(1);



			// Try creating the reader from the string
			StringReader strReader = new StringReader(strStripBOM);

			XmlReader reader = XmlReader.Create(strReader, settings);
			// Create the new XMLdoc and load the content into it.
			XmlDocument doc = new XmlDocument();
			doc.PreserveWhitespace = true;
			doc.Load(reader);

			return doc;
		}

		/// <summary>
		/// Given a path to a file of any type, returns a base-64 encoded blob that can be fed directly to the CMS.
		/// </summary>
		public byte[] GetIshBlobFromFile(string FilePath)
		{
			byte[] Data = null;
			try {
				FileInfo fInfo = new FileInfo(FilePath);
				long numBytes = fInfo.Length;
				FileStream fStream = new FileStream(FilePath, FileMode.Open, FileAccess.Read);
				BinaryReader br = new BinaryReader(fStream);
				Data = br.ReadBytes(Convert.ToInt32(numBytes));
				br.Close();
				fStream.Close();
				return Data;
			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(3, FilePath + ": Failed to convert content to Base64 blob. Message: " + ex.Message, strModuleName + "-GetIshBlobFromFile");
				return null;
			}
		}
		public string GetFilenameFromIshMeta(XmlDocument IshMetaData)
		{
			StringBuilder filename = new StringBuilder();
			try {
				//first, check to make sure we actually have attributes we need:
				if (IshMetaData.ChildNodes.Count == 0) {
					return "";
				}
				//Next, return the basic data that all objects have:
				filename.Append(IshMetaData.SelectSingleNode("/ishobjects/ishobject/ishfields/ishfield[contains(@name, 'FTITLE')]").InnerText + "=");
				filename.Append(IshMetaData.SelectSingleNode("/ishobjects/ishobject/@ishref").Value + "=");
				// get guid
				filename.Append(IshMetaData.SelectSingleNode("/ishobjects/ishobject/ishfields/ishfield[contains(@name, 'VERSION')]").InnerText + "=");
				filename.Append(IshMetaData.SelectSingleNode("/ishobjects/ishobject/ishfields/ishfield[contains(@name, 'DOC-LANGUAGE')]").InnerText + "=");

				//last, check to see if we have an image.  If we do, we need to get the data and return it as part of the string too.  Otherwise, we just return ""
				string resolution = "";
				if ((IshMetaData.SelectSingleNode("/ishobjects/ishobject/ishfields/ishfield[contains(@name, 'FRESOLUTION')]") != null)) {
					//there is text here... let's capture it:
					resolution = IshMetaData.SelectSingleNode("/ishobjects/ishobject/ishfields/ishfield[contains(@name, 'FRESOLUTION')]").InnerText;
				}
				filename.Append(resolution);

				return filename.ToString();
			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(3, "One or more metadata was not found while trying to build a filename from a located GUID. " + ex.Message, strModuleName);
				return "";
			}
			return true;
		}

		public string GetMetaDataXMLStucture(string CMSTitle, string Version, string Author, string Status, string Resolution, string Language, string ModuleType = "", string Illustrator = "mmatus")
		{
			//If CMSTitle = "" Or Version = "" Or Author = "" Or Status = "" Then
			//    'if one or more required fields are blank, abort opperation!
			//    modErrorHandler.Errors.PrintMessage(3, "One or more required Metadata fields are blank. Check the Author, Status, CMSTitle, and Version values.", strModuleName + "-GetMetadataXMLStructure")
			//    Return ""
			//End If
			StringBuilder XMLString = new StringBuilder();
			XMLString.Append("<ishfields>");
			if (CMSTitle.Length > 0 & Language == "en") {
				XMLString.Append("<ishfield name=\"FTITLE\" level=\"logical\">");
				XMLString.Append(CMSTitle);
				XMLString.Append("</ishfield>");
			}
			//If we have a ModuleType, we need to set the metadata appropriately for the LOV.
			if (!string.IsNullOrEmpty(ModuleType)) {
				switch (ModuleType.ToLower()) {
					case "concept":
					case "task":
					case "reference":
					case "topic":
					case "dita":
					case "troubleshooting":
					case "glossary":
					case "glossary term":
						XMLString.Append("<ishfield name=\"FMODULETYPE\" level=\"logical\">");
						break;
					case "map":
					case "submap":
						XMLString.Append("<ishfield name=\"FMASTERTYPE\" level=\"logical\">");
						break;
					case "graphic":
					case "icon":
					case "screenshot":
						XMLString.Append("<ishfield name=\"FILLUSTRATIONTYPE\" level=\"logical\">");

						break;
				}
				XMLString.Append(ModuleType);
				XMLString.Append("</ishfield>");
			}
			//if we have an image, need to set the default illustrator for it.
			if (Resolution.Length > 0) {
				XMLString.Append("<ishfield type=\"hidden\" name=\"FNOTRANSLATIONMGMT\" level=\"logical\" label=\"Disable translation management\">No</ishfield>");
			}
			if (Author.Length > 0) {
				XMLString.Append("<ishfield name=\"FAUTHOR\" level=\"lng\">");
				XMLString.Append(Author);
				XMLString.Append("</ishfield>");
			}
			if (Status.Length > 0) {
				XMLString.Append("<ishfield name=\"FSTATUS\" level=\"lng\">");
				XMLString.Append(Status);
				XMLString.Append("</ishfield>");
			}
			XMLString.Append("</ishfields>");
			return XMLString.ToString();
		}
		/// <summary>
		/// Retrieves CMS metadata from a local file including CMSFilename for XML files.  File must be exported from the CMS or preprocessed for the CMS.
		/// </summary>
		public bool GetCommonMetaFromLocalFile(string LocalFilePath, ref string CMSFileName, ref string GUID, ref string Version, ref string Language, ref string Resolution)
		{
			//most commonly used to get parameters for deleting an object in the CMS...
			FileInfo myfile = new FileInfo(LocalFilePath);
			ArrayList aryMeta = new ArrayList();
			try {
				foreach (string metapiece in myfile.Name.Replace(myfile.Extension, "").Split("=")) {
					aryMeta.Add(metapiece);
				}
				CMSFileName = aryMeta[0];
				GUID = aryMeta[1];
				Version = aryMeta[2];
				Language = aryMeta[3];
				if (aryMeta.Count > 4) {
					Resolution = aryMeta[4];
				} else {
					Resolution = "";
				}
				//Extension = myfile.Extension
				return true;
			} catch (Exception ex) {
				return false;
			}
		}
		/// <summary>
		/// Retrieves CMS metadata from a local file excluding CMSFilename for non-XML files.  File must be exported from the CMS or preprocessed for the CMS.
		/// </summary>
		public bool GetCommonMetaFromLocalFile(string LocalFilePath, ref string CMSFileName, ref string CMSTitle, ref string GUID, ref string Version, ref string Language, ref string Resolution)
		{
			//used to get metadata used for creating or modifying content in the CMS.
			FileInfo myfile = new FileInfo(LocalFilePath);
			ArrayList aryMeta = new ArrayList();
			try {
				foreach (string metapiece in myfile.Name.Replace(myfile.Extension, "").Split("=")) {
					aryMeta.Add(metapiece);
				}
				CMSFileName = aryMeta[0];
				GUID = aryMeta[1];
				Version = aryMeta[2];
				Language = aryMeta[3];
				if (aryMeta.Count > 4) {
					Resolution = aryMeta[4];
				} else {
					Resolution = "";
				}


				switch (myfile.Extension) {
					case ".xml":
					case ".ditamap":
					case ".dita":
						// read the xml file into an xmldoc 
						XmlDocument doc = new XmlDocument();
						doc = LoadFileIntoXMLDocument(LocalFilePath);
						//Get the title info from the doc:
						if ((doc.SelectSingleNode("//title[1]") != null)) {
							CMSTitle = doc.SelectSingleNode("//title[1]").InnerText;
						} else if ((doc.DocumentElement.Attributes.GetNamedItem("title") != null)) {
							CMSTitle = doc.DocumentElement.Attributes.GetNamedItem("title").InnerText;
						} else {
							CMSTitle = CMSFileName;
						}
						CMSTitle = CMSTitle.Replace("&", "&amp;");
						CMSTitle = CMSTitle.Replace("<", "&lt;");
						CMSTitle = CMSTitle.Replace(">", "&gt;");
						CMSTitle = CMSTitle.Replace("\\", "");
						CMSTitle = CMSTitle.Replace("/", "");
						CMSTitle = CMSTitle.Replace(":", "");
						CMSTitle = CMSTitle.Replace("*", "");
						CMSTitle = CMSTitle.Replace("?", "");
						CMSTitle = CMSTitle.Replace("\"", "");
						CMSTitle = CMSTitle.Replace("<", "");
						CMSTitle = CMSTitle.Replace(">", "");
						CMSTitle = CMSTitle.Replace("|", "");
						CMSTitle = CMSTitle.Replace(Constants.vbCrLf, " ");
						CMSTitle = CMSTitle.Replace(Constants.vbCr, " ");
						CMSTitle = CMSTitle.Replace(Constants.vbLf, " ");
						CMSTitle = CMSTitle.Replace("  ", " ");
						CMSTitle = CMSTitle.Replace("  ", " ");
						CMSTitle = CMSTitle.Replace("  ", " ");
						CMSTitle = CMSTitle.Replace("  ", " ");
						break;
					default:
						//just assign the title as the filename portion.
						CMSTitle = CMSFileName;
						break;
				}

				return true;
			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(3, "Failed get metadata based on filename and/or content for: " + LocalFilePath + " Error: " + ex.Message, strModuleName);
				return false;
			}


		}
		public void CopyDTDFile()
		{
			try {
				string DTDpath = null;
				DTDpath = Path.GetTempPath() + "nbsp.dtd";

				System.IO.Stream clsResourceStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("IshModulesNS.nbsp.dtd");
				byte[] bResource = (byte[])Array.CreateInstance(typeof(byte), clsResourceStream.Length);
				clsResourceStream.Read(bResource, 0, bResource.Length);
				clsResourceStream.Close();
				clsResourceStream.Dispose();
				string sResource = System.Text.Encoding.ASCII.GetString(bResource);
				System.IO.TextWriter b = new System.IO.StreamWriter(DTDpath, false, System.Text.Encoding.Unicode);
				b.Write(sResource);
				// Clear object 
				b.Flush();
				b.Close();

			} catch (Exception ex) {
			}

		}

		public string RemoveWindowsIllegalChars(string strFileorFolder)
		{
			strFileorFolder = strFileorFolder.Replace("‚", "");
			// Special 'lower, single quotation mark'.
			strFileorFolder = strFileorFolder.Replace(",", "");
			strFileorFolder = strFileorFolder.Replace("\\", "");
			strFileorFolder = strFileorFolder.Replace("/", "");
			strFileorFolder = Strings.Replace(strFileorFolder, "*", "");
			strFileorFolder = Strings.Replace(strFileorFolder, "?", "");
			strFileorFolder = Strings.Replace(strFileorFolder, ">", "");
			strFileorFolder = Strings.Replace(strFileorFolder, "<", "");
			strFileorFolder = Strings.Replace(strFileorFolder, ":", "");
			strFileorFolder = Strings.Replace(strFileorFolder, "|", "");
			strFileorFolder = Strings.Replace(strFileorFolder, "#", "");
			strFileorFolder = Strings.Replace(strFileorFolder, "!", "");
			strFileorFolder = Strings.Replace(strFileorFolder, "\"", "");
			return strFileorFolder;
		}


		public bool SaveTextToFile(string strData, string FullPath, string ErrInfo = "")
		{

			bool bAns = false;
			StreamWriter objReader = null;

			try {

				objReader = new StreamWriter(FullPath, false, Encoding.Unicode);
				objReader.Write(strData);
				objReader.Close();
				bAns = true;
			} catch (Exception Ex) {
				ErrInfo = Ex.Message;

			}
			return bAns;
		}
		//SaveTextToFile
		/// <summary>
		/// Loads any FilePath specified XML file into an XML Document.  Works on all valid XML regardless of entities or DTD declarations.
		/// </summary>
		/// <returns>Returns a fully-formed XML Document object.</returns>
		public XmlDocument LoadFileIntoXMLDocument(string FilePath)
		{
			XmlReaderSettings settings = null;
			DITAResolver resolver = new DITAResolver();
			settings = new XmlReaderSettings();
			settings.DtdProcessing = DtdProcessing.Parse;
			//False
			settings.ValidationType = ValidationType.None;
			settings.XmlResolver = resolver;
			settings.CloseInput = true;
			settings.IgnoreWhitespace = false;
			try {
				XmlReader reader = XmlReader.Create(FilePath, settings);
				XmlDocument doc = new XmlDocument();
				doc.PreserveWhitespace = true;
				doc.Load(reader);
				reader.Close();
				return doc;
			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(3, "Failed to load document as xml: " + FilePath + " ErrorMessage: " + ex.Message, strModuleName);
				return null;
			}

		}
		/// <summary>
		/// Builds a template of standard metadata to be requested from the CMS.
		/// </summary>
		/// <returns>Stringbuilder of metadata for the CMS. </returns>
		public StringBuilder BuildRequestedMetadata()
		{
			StringBuilder requestedmeta = new StringBuilder();
			requestedmeta.Append("<ishfields>");
			requestedmeta.Append("<ishfield name=\"FTITLE\" level=\"logical\"/>");
			requestedmeta.Append("<ishfield name=\"VERSION\" level=\"version\"/>");
			//If Resolution = "" Then
			requestedmeta.Append("<ishfield name=\"FAUTHOR\" level=\"lng\"/>");
			//Else
			requestedmeta.Append("<ishfield name=\"FRESOLUTION\" level=\"lng\"/>");
			//End If
			requestedmeta.Append("<ishfield name=\"FSTATUS\" level=\"lng\"/>");
			requestedmeta.Append("<ishfield name=\"DOC-LANGUAGE\" level=\"lng\"/>");
			//requestedmeta.Append("<ishfield name=""EDT-FILE-EXTENSION"" level=""lng""/>")
			requestedmeta.Append("</ishfields>");
			return requestedmeta;
		}
		public StringBuilder BuildFullPubMetadata()
		{
			StringBuilder requestedmeta = new StringBuilder();
			requestedmeta.Append("<ishfields>");
			requestedmeta.Append("<ishfield name=\"VERSION\" level=\"version\"/>");
			requestedmeta.Append("<ishfield name=\"FISHPUBSTATUS\" level=\"lng\"/>");
			requestedmeta.Append("<ishfield name=\"FISHISRELEASED\" level=\"version\"/>");
			requestedmeta.Append("<ishfield name=\"DOC-LANGUAGE\" level=\"lng\"/>");
			requestedmeta.Append("<ishfield name=\"FISHOUTPUTFORMATREF\" level=\"lng\"/>");
			requestedmeta.Append("<ishfield name=\"FTITLE\" level=\"logical\"/>");
			requestedmeta.Append("</ishfields>");
			return requestedmeta;
		}
		public StringBuilder BuildMinPubMetadata()
		{
			StringBuilder requestedmeta = new StringBuilder();
			requestedmeta.Append("<ishfields>");
			requestedmeta.Append("<ishfield name=\"VERSION\" level=\"version\"/>");
			requestedmeta.Append("<ishfield name=\"FISHISRELEASED\" level=\"version\"/>");
			requestedmeta.Append("<ishfield name=\"FTITLE\" level=\"logical\"/>");
			requestedmeta.Append("<ishfield name=\"FISHMASTERREF\" level=\"version\"/>");
			requestedmeta.Append("<ishfield name=\"FISHBASELINE\" level=\"version\"/>");
			requestedmeta.Append("</ishfields>");
			return requestedmeta;
		}
		public StringBuilder BuildPubMetaDataFilter(string GUID, string Version = "latest", string Language = "", string OutputType = "all")
		{
			StringBuilder requestedmeta = new StringBuilder();

			requestedmeta.Append("<ishfields>");


			if (Version == "latest") {
				//Get all versions. Returns multiple ishobjects.
			} else {
				//Get the specified version only.
				requestedmeta.Append("<ishfield name=\"FMAPID\" level=\"lng\">" + GUID + "</ishfield>");
			}

			if (OutputType == "all") {
				//Get all versions. By not including the filter, all outputs will be returned.
			} else {
				//Get the specified output type only
				requestedmeta.Append("<ishfield name=\"FISHOUTPUTFORMATREF\" level=\"lng\">" + OutputType + "</ishfield>");
			}

			if (Language.Length > 0) {
				//A specific language is wanted. Return only outputs that are of that language.  otherwise, all languages are returned.
				requestedmeta.Append("<ishfield name=\"DOC-LANGUAGE\" level=\"lng\">" + Language + "</ishfield>");
			}
			requestedmeta.Append("<ishfield name=\"VERSION\" level=\"version\">" + Version + "</ishfield>");
			requestedmeta.Append("</ishfields>");
			return requestedmeta;
		}

		public class DITAResolver : XmlUrlResolver
		{
			private string strModuleName {
				get { return "DITAResolver"; }
			}

			public static Hashtable myHash = new Hashtable();
			public DITAResolver()
			{
			}
			//New

			public override object GetEntity(Uri absoluteUri, string role, Type ofObjectToReturn)
			{
				string pubid1slash = null;
				string Pubid = null;
				Pubid = "-//MYCOMPANY//DTD DITA ";
				pubid1slash = "-/MYCOMPANY/DTD DITA ";
				string mapid1slash = null;
				string Mapid = null;
				Mapid = "-//OASIS//DTD DITA ";
				mapid1slash = "-/OASIS/DTD DITA ";
				string absURI = absoluteUri.ToString();
				if ((absoluteUri.ToString().Contains(Pubid) | absoluteUri.ToString().Contains(Mapid) | absoluteUri.ToString().Contains(pubid1slash) | absoluteUri.ToString().Contains(mapid1slash))) {
					string DTDpath = null;
					DTDpath = Path.GetTempPath() + "nbsp.dtd";
					return new FileStream(DTDpath, FileMode.Open, FileAccess.Read, FileShare.Read);
					// Return New FileStream(DTDTopic, FileMode.Open, FileAccess.Read, FileShare.Read)


				} else if (myHash.ContainsKey(absoluteUri)) {
					modErrorHandler.Errors.PrintMessage(1, "Reading resource" + absoluteUri.ToString() + " from cached stream", strModuleName);
					//Returns the cached stream.
					return new FileStream((String)myHash[absoluteUri], FileMode.Open, FileAccess.Read, FileShare.Read);
				} else {
					return base.GetEntity(absoluteUri, role, ofObjectToReturn);

				}
			}
			//GetEntity
		}
		//CustomResolver
	}
}
