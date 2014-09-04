using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Xml;
using ErrorHandlerNS;
using System.IO;
using System.Text;
using ISHModulesNS.clsCommonFuncs;
namespace ISHModulesNS
{

	public class IshDocument : mainAPIclass
	{
		#region "Private Members"

		private readonly string strModuleName = "IshDocument";
			#endregion
		private Hashtable m_recursion_hash = new Hashtable();
		#region "Constructors"
		public IshDocument(string Username, string Password, string ServerURL)
		{
			//Note that you should use the full URL up to the WS address (https://serverURL/InfoShareWS)
			oISHAPIObjs = new ISHObjs(Username, Password, ServerURL);
			//oISHAPIObjs.ISHAppObj.Login("InfoShareAuthor", Username, Password, Context)
		}
		#endregion
		#region "Properties"

		#endregion
		#region "Methods"

		public bool MoveObject(string GUID, long ToFolderID)
		{
			//TODO: Has to search through the entire CMS Structure just to figure out the current folder ID of the specified GUID...  Terribly inefficient!
			//The problem is that FolderID is not tracked with objects as part of its metadata.  Looking up an object gives you NO information about where it exists in the CMS.
			try {
				long CurrentFolder = FindFolderIDforObjbyGUID(GUID, 0);
				if (CurrentFolder > 0) {
					oISHAPIObjs.ISHDocObj.Move(GUID, CurrentFolder.ToString(), ToFolderID.ToString());
				} else {
					modErrorHandler.Errors.PrintMessage(3, "Unable to move object " + GUID + " to specified location!", strModuleName + "-MoveObject");
				}
				return true;
			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(3, "Unable to move object " + GUID + " to specified location!", strModuleName + "-MoveObject");
				return false;
			}
		}

		public bool ChangeState(string strDesiredState, string strGUID, string strVer, string strLanguage, string strResolution)
		{
			dynamic myCurrentState = GetCurrentState(strGUID, strVer, strResolution, strLanguage);
			bool processingresult = true;
			string strMetaState = "<ishfields><ishfield name=\"FSTATUS\" level=\"lng\">" + strDesiredState + "</ishfield></ishfields>";
			//'Could be used to update person assigned to a specific role as well...
			//Dim strMetaRole As String = "<ishfields><ishfield name=""FEDITOR"" level=""lng"">" + strEditorName + "</ishfield></ishfields>"
			bool result = true;

			// Generic move status drive used to change the status of a topic and COULD be used to update the name associated with the status
			if ((string.IsNullOrEmpty(Context))) {
				modErrorHandler.Errors.PrintMessage(3, "Context is not set. Unable to continue.", strModuleName + "-ChangeState");
				return false;
			}

			if ((CanMoveToState(strDesiredState, strGUID, strVer, strResolution, strLanguage))) {
				// Change the state
				if ((SetMeta(strMetaState, strGUID, strVer, strResolution, strLanguage))) {
					modErrorHandler.Errors.PrintMessage(1, "Changed state for " + strGUID + "=" + strVer + "=" + strResolution + "=" + strLanguage + ".", strModuleName + "-ChangeState");
					//' Check to see if we'r esupposed to assign the editor name as well...

					//If (GetCurrentState(strGUID, strVer, strResolution, strLanguage) = "Editing") And strEditorName.Length > 0 Then
					//    ' Now set the metadata on the role to the current user's name
					//    ' For example, set "lgalindo" as FEDITOR
					//    If SetMeta(strMetaRole, strGUID, strVer, strResolution, strLanguage) Then
					//        modErrorHandler.Errors.PrintMessage(1, "Changed editor for " + strGUID + "=" + strVer + "=" + strResolution + "=" + strLanguage + " to " + strEditorName + ".", strModuleName + "-ChangeState")
					//    Else
					//        'Something happened when changing the Editor!
					//        modErrorHandler.Errors.PrintMessage(2, "Failed to change editor for " + strGUID + "=" + strVer + "=" + strResolution + "=" + strLanguage + " to " + strEditorName + ".", strModuleName + "-ChangeState")
					//        processingresult = False
					//    End If

					//End If



				} else {
					//Something happened when changing the state!
					modErrorHandler.Errors.PrintMessage(3, "Failed to change state for " + strGUID + "=" + strVer + "=" + strResolution + "=" + strLanguage + ". Current state is: " + myCurrentState, strModuleName + "-ChangeState");
					processingresult = false;
				}
			} else {
				//' If the current state is editing, just make sure that the editor name is updated
				//If (GetCurrentState(strGUID, strVer, strResolution) = "Editing") And strEditorName.Length > 0 Then
				//    ' Now set the metadata on the role to the current user's name
				//    ' For example, set "lgalindo" as FEDITOR
				//    If SetMeta(strMetaRole, strGUID, strVer, strResolution) Then
				//        modErrorHandler.Errors.PrintMessage(1, "Changed editor for " + strGUID + "=" + strVer + "=" + strResolution + " to " + strEditorName + ".", strModuleName + "-ChangeState")
				//    Else
				//        'Something happened when changing the Editor!
				//        modErrorHandler.Errors.PrintMessage(2, "Failed to change editor for " + strGUID + "=" + strVer + "=" + strResolution + " to " + strEditorName + ".", strModuleName + "-ChangeState")
				//        processingresult = False
				//    End If
				//End If
				if (myCurrentState == strDesiredState) {
					modErrorHandler.Errors.PrintMessage(1, "State already set as requested for " + strGUID + "=" + strVer + "=" + strResolution + "=" + strLanguage + ".", strModuleName + "-ChangeState");
					processingresult = true;
				} else {
					modErrorHandler.Errors.PrintMessage(3, "Failed to change state for " + strGUID + "=" + strVer + "=" + strResolution + "=" + strLanguage + ". Current state is: " + myCurrentState, strModuleName + "-ChangeState");
					processingresult = false;
				}
			}



			return processingresult;
		}

		public bool CheckIn(string PathToCheckInFile)
		{
			FileInfo checkinfile = new FileInfo(PathToCheckInFile);
			dynamic CMSFilename = null;
			dynamic GUID = null;
			dynamic Version = null;
			dynamic Language = null;
			string Resolution = new string("");
			oCommonFuncs.GetCommonMetaFromLocalFile(checkinfile.FullName, ref CMSFilename, ref GUID, ref Version, ref Language, ref Resolution);
			byte[] checkinblob = oCommonFuncs.GetIshBlobFromFile(checkinfile.FullName);
			try {
				oISHAPIObjs.ISHDocObj.CheckIn(GUID, Version, Language, Resolution, "", oCommonFuncs.GetISHEdt(checkinfile.Extension), checkinblob);
				modErrorHandler.Errors.PrintMessage(1, "Checked in object " + checkinfile.FullName + ".", strModuleName + "-CheckIn");
				checkinfile.Attributes = FileAttributes.Normal;
				checkinfile.Delete();
				return true;
			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(3, "Failed to check in object " + checkinfile.FullName + ". Message: " + ex.Message, strModuleName + "-CheckIn");
				return false;
			}
		}
		public bool CheckOut(string GUID, string Version, string Language, string Resolution, string LocalStorePath)
		{
			string CheckOutFile = new string("");
			//first, ensure it exists.
			if (ObjectExists(GUID, Version, Language, Resolution)) {
				//Check out the object
				try {
					oISHAPIObjs.ISHDocObj.CheckOut(GUID, ref Version, Language, Resolution, "", ref CheckOutFile);
					modErrorHandler.Errors.PrintMessage(1, "Checked out object " + GUID + "=" + Version + "=" + Language + "=" + Resolution + ".", strModuleName + "-CheckOut");
					CheckOutFile = "";
				} catch (Exception ex) {
					if (GetCurrentState(GUID, Version, Resolution, Language) == "Released") {
						CreateNewVersion(GUID, ref Version, Language, Resolution);
						//now try checking out the new version:
						try {
							oISHAPIObjs.ISHDocObj.CheckOut(GUID, ref Version, Language, Resolution, "", ref CheckOutFile);
						} catch (Exception ex3) {
							modErrorHandler.Errors.PrintMessage(3, "Failed to check out object " + GUID + "=" + Version + "=" + Language + "=" + Resolution + ". Message: " + ex3.Message, strModuleName + "-CheckOut");
							return false;
						}
					}
				}
				//Now try fetching the object to our local system. Don't do this for objects that have a "High" resolution if the current one being processed is Low.
				try {
					if (ObjectExists(GUID, Version, Language, "High") & Resolution == "Low") {
						//Skip it.
					} else {
						//Grab it down locally.
						GetObjByID(GUID, Version, Language, Resolution, LocalStorePath);
						modErrorHandler.Errors.PrintMessage(1, "Downloaded object " + GUID + "=" + Version + "=" + Language + "=" + Resolution + ".", strModuleName + "-CheckOut");
					}
				} catch (Exception ex) {
					modErrorHandler.Errors.PrintMessage(3, "Failed to download checked-out object " + GUID + "=" + Version + "=" + Language + "=" + Resolution + ". Message: " + ex.Message, strModuleName + "-CheckOut");
				}
				return true;
			} else {
				return false;
			}

		}
		/// <summary>
		/// Given the current informaiton of a given CMS Object, creates a new version on the same branch containing the 
		/// content from the previous version in a Draft state.
		/// </summary>
		/// <param name="GUID">Unique object ID</param>
		/// <param name="Version">Passes in the current version of the object to be versioned, returns the new version number</param>
		/// <param name="Language">Language</param>
		/// <param name="Resolution">(Optional) Resolution</param>
		/// <returns>True if successful, False if failed</returns>
		/// <remarks></remarks>
		public bool CreateNewVersion(string GUID, ref string Version, string Language, string Resolution = "")
		{
			//Specified version is released so we need to create a new version of the specified branch.
			//Get the existing content at the current version.
			XmlDocument newverDoc = GetObjByID(GUID, Version, Language, Resolution);
			//if the content is a map or topic, we need to remove the processing instruction before it can be added as a new version.
			byte[] datablob = null;
			string newverIshType = oCommonFuncs.GetISHTypeFromMeta(newverDoc);
			if (newverIshType == "ISHModule" | newverIshType == "ISHMasterDoc") {
				XmlNode MyNode = newverDoc.SelectSingleNode("//ishdata");
				//get the dita topic out of the CData:
				XmlDocument DITATopic = oCommonFuncs.GetXMLOut(MyNode);
				//drop the ISH version specific ProcInstr
				XmlNode ishnode = DITATopic.SelectSingleNode("/processing-instruction('ish')");
				DITATopic.RemoveChild(ishnode);
				DITATopic.Save("c:\\temp\\deletetopic.xml");
				//load the doc to a datablob:
				//Convert the doc to an ISH blob
				datablob = oCommonFuncs.GetIshBlobFromFile("c:\\temp\\deletetopic.xml");
				//delete local file
				File.Delete("c:\\temp\\deletetopic.xml");
			} else {
				//get the blob (images only) needed to create the new version:
				datablob = oCommonFuncs.GetBinaryOut(newverDoc.SelectSingleNode("//ishdata"));
			}

			//get the various required parameters needed to create the new version:
			DocumentObj20ServiceReference.ISHType IshType = StringToISHType(oCommonFuncs.GetISHTypeFromMeta(newverDoc));
			//Dim basefolder As DocumentObj25ServiceReference.BaseFolder
			string[] folderpath = {
				
			};
			long[] folderID = {
				
			};

			oISHAPIObjs.ISHDocObj25.FolderLocation(ref folderpath, ref folderID, GUID);
			//Need to first collect the current version info and then drop it.
			XmlNode ishfields = newverDoc.SelectSingleNode("//ishfields");
			//TODO: This doesn't currently handle branched objects... would need to pull the final value (after the last '.') in the string, increment it, then replace the final value.
			XmlNode VersionNode = ishfields.SelectSingleNode("ishfield[@name='VERSION']");
			Version = VersionNode.InnerText;
			Version = Strings.Trim(Conversion.Str(Conversion.Int(Version) + 1));
			XmlNode delfield = ishfields.SelectSingleNode("ishfield[@name='VERSION']");
			ishfields.RemoveChild(delfield);
			string XMLMetaData = ishfields.OuterXml;
			string psEDT = newverDoc.SelectSingleNode("//ishdata").Attributes.GetNamedItem("edt").Value;
			//if the container version exists, just create the resolution within it.  otherwise, create the new version too.
			if (ObjectExists(GUID, Version, Language)) {
				try {
					//Create the new language content on the existing new version (created previously, but not populated with this resolution's content):
					oISHAPIObjs.ISHDocObj.CreateOrUpdate(folderID[folderID.Length - 1], IshType, ref GUID, ref Version, Language, Resolution, "Draft", XMLMetaData, XMLMetaData, psEDT,
					datablob);
				} catch (Exception ex2) {
					modErrorHandler.Errors.PrintMessage(3, "Failed to create new version of object " + GUID + "=" + Version + "=" + Language + "=" + Resolution + ". Message: " + ex2.Message, strModuleName + "-CheckOut");
					return false;
				}
			} else {
				try {
					//Create the new version AND language:
					oISHAPIObjs.ISHDocObj.CreateOrUpdate(folderID[folderID.Length - 1], IshType, ref GUID, ref "new", Language, Resolution, "Draft", XMLMetaData, XMLMetaData, psEDT,
					datablob);
				} catch (Exception ex2) {
					modErrorHandler.Errors.PrintMessage(3, "Failed to create new version of object " + GUID + "=" + Version + "=" + Language + "=" + Resolution + ". Message: " + ex2.Message, strModuleName + "-CheckOut");
					return false;
				}
			}

		}
		public string GetLatestVersionNumber(string GUID)
		{
			//Get the existing content at the current version.
			XmlDocument VerDoc = GetObjByID(GUID, "latest", "en", "");
			if (VerDoc == null) {
				VerDoc = GetObjByID(GUID, "latest", "en", "Low");
			}
			XmlNode ishfields = VerDoc.SelectSingleNode("//ishfields");
			XmlNode VersionNode = ishfields.SelectSingleNode("ishfield[@name='VERSION']");
			return VersionNode.InnerText;
		}




		/// <summary>
		/// Imports specified file using the parameters provided.  Returns true if successful.
		/// </summary>
		/// <param name="FilePath">Path to the file to be imported.</param>
		/// <param name="CMSFolderID">ID of the CMS folder to import to.</param>
		/// <param name="Author">Author of the object being imported.</param>
		/// <param name="ISHType">Type of content being imported.  Allowed values are: "ISHIllustration", "ISHLibrary", "ISHMasterDoc", "ISHModule", "ISHNone", "ISHPublication", "ISHReusedObj", and "ISHTemplate"</param>
		/// <param name="ReturnedGUID">Returns GUID set by CMS upon successful import.</param>
		/// <param name="CMSTitle">Title to be used for the object within the CMS. This is shown to users of the CMS and can be updated in the object Properties.</param>
		/// <param name="ObjectMetaType">Specifies the value found in the LOV for a module or image (Graphic, Icon, Screenshot, Concept, Reference, Task, etc.).  Value is arbitrary but is only properly set if found in the CMS.</param>
		public bool ImportObject(string FilePath, long CMSFolderID, string Author, string ISHType, string strState, ref string ReturnedGUID, ref string CMSTitle, string ObjectMetaType = "")
		{
			if (File.Exists(FilePath)) {
				//Dim CDataNode As XmlNode
				//Dim CData As String = ""
				//Dim decodedBytes As Byte()
				//Dim decodedText As String
				string CMSFileName = "";
				string GUID = "";
				string Version = "";
				string Language = "";
				string Resolution = "";
				//Get our meta from the file:
				if (oCommonFuncs.GetCommonMetaFromLocalFile(FilePath, ref CMSFileName, ref CMSTitle, ref GUID, ref Version, ref Language, ref Resolution) == false) {
					return false;
				}
				//check to see if user has commas in their titles, report if they do and fail the import:
				if (CMSTitle.Contains(",")) {
					modErrorHandler.Errors.PrintMessage(2, "Object's title contains character(s) not allowed by the CMS. Replacing with equivalent Unicode character '‚' (Single Low Quotation Mark). Title: '" + CMSTitle + "'. File: " + FilePath + ".", strModuleName + "importobj-getmeta");
					CMSTitle = CMSTitle.Replace(",", "‚");
					//Return False
				}
				//Create the MetaXML string based on our metadata:
				string metaxml = null;
				metaxml = oCommonFuncs.GetMetaDataXMLStucture(CMSTitle, Version, Author, strState, Resolution, Language, ObjectMetaType);
				if (string.IsNullOrEmpty(metaxml)) {
					modErrorHandler.Errors.PrintMessage(3, "Failed to generate the xml metadata needed to create the content in the CMS. Aborting import.", strModuleName + "importobj-getmeta");
					return false;
				}
				//now that we have the meta, need to get the bytearray data blob
				byte[] data = null;
				data = oCommonFuncs.GetIshBlobFromFile(FilePath);


				string result = "";
				// Import the content if it doesn't already exist in the CMS
				if (ObjectExists(GUID, Version, Language, Resolution) == false & Language == "en") {
					try {
						oISHAPIObjs.ISHDocObj.Create(CMSFolderID.ToString(), StringToISHType(ISHType), ref GUID, ref Version, Language, Resolution, metaxml, oCommonFuncs.GetISHEdt(Path.GetExtension(FilePath)), data);
						ReturnedGUID = GUID;
						//'if objectmetatype is icon, also import thumbnail as new resolution
						//If ObjectMetaType = "Icon" Then
						//    ISHDocObj.Create(CMSFolderID.ToString, StringToISHType(ISHType), ReturnedGUID, Version, Language, "Thumbnail", metaxml, GetISHEdt(Path.GetExtension(FilePath)), data)
						//End If
						return true;
					} catch (Exception ex) {
						if (ex.Message.ToString().Contains("-227") & ex.Message.ToString().Contains("Check that the id in the document is the same as the DocId provided via metadata")) {
							modErrorHandler.Errors.PrintMessage(3, GUID + " contains an invalid ID.  Most likely, the Public ID in the DOCTYPE declaration is not recognized by the catalog lookup/DTDs. Message from CMS: " + ex.Message, strModuleName + "-importobj-CMSCreate");
						} else {
							modErrorHandler.Errors.PrintMessage(3, "Failed to import " + GUID + " to the CMS: " + ex.Message, strModuleName + "-importobj-CMSCreate");
						}



						return false;
					}
				} else {
					ReturnedGUID = GUID;
					//May be a localization import.  Check the language code to see if it's 'en'
					if (Language == "en") {
						//Is an EN object but it exists in the DB, so report it as already imported.
						modErrorHandler.Errors.PrintMessage(2, "Object to be imported already exists in the CMS. Skipping import for " + FilePath, strModuleName + "-ImportObject");

						return true;
						// It's already imported, let the user know.
					} else {
						//This is a localized file that needs imported.
						//Get the current state of the file being imported.
						string currentstate = GetCurrentState(GUID, Version, Resolution, Language);
						//if it is not "In Translation", change it so that it can be imported.
						if (!(currentstate == "In Translation")) {
							string strMetaState = "<ishfields><ishfield name=\"FSTATUS\" level=\"lng\">In Translation</ishfield></ishfields>";
							if (SetMeta(strMetaState, GUID, Version, Resolution, Language) == false) {
								modErrorHandler.Errors.PrintMessage(3, "Failed to change object state to 'In Translation' to allow reimport. File: " + FilePath, strModuleName + "importobj-updatel10n");
								return false;
							}
						}
						//Now we should be able to import it...

						string currentmeta = "";
						//Get the current metadata to allow the update
						oISHAPIObjs.ISHDocObj.GetMetaData(GUID, ref Version, Language, Resolution, "<ishfields><ishfield name=\"FTITLE\" level=\"logical\"/><ishfield name=\"FSTATUS\" level=\"lng\"/></ishfields>", ref currentmeta);
						try {
							//attempt to update the current content with the new content and change the state to "Translated":
							oISHAPIObjs.ISHDocObj.Update(GUID, ref Version, Language, Resolution, metaxml, currentmeta, oCommonFuncs.GetISHEdt(Path.GetExtension(FilePath)), data);
							return true;
						} catch (Exception ex) {
							modErrorHandler.Errors.PrintMessage(3, "Failed to import a file to the CMS. File: " + FilePath + ". Error Message: " + ex.Message, strModuleName + "importobj-updatel10n");
							return false;
						}
					}
				}
			} else {
				return false;
				//File didn't exist locally...
			}
		}


		/// <summary>
		/// Converts a string (ISHIllustration, ISHBaseline, etc.) to a valid ISHType object.
		/// </summary>
		public DocumentObj20ServiceReference.ISHType StringToISHType(string IshType)
		{
			switch (IshType) {
				case "ISHBaseline":
					return ISHModulesNS.DocumentObj20ServiceReference.ISHType.ISHBaseline;
				case "ISHIllustration":
					return ISHModulesNS.DocumentObj20ServiceReference.ISHType.ISHIllustration;
				case "ISHLibrary":
					return ISHModulesNS.DocumentObj20ServiceReference.ISHType.ISHLibrary;
				case "ISHMasterDoc":
					return ISHModulesNS.DocumentObj20ServiceReference.ISHType.ISHMasterDoc;
				case "ISHModule":
					return ISHModulesNS.DocumentObj20ServiceReference.ISHType.ISHModule;
				case "ISHNone":
					return ISHModulesNS.DocumentObj20ServiceReference.ISHType.ISHNone;
				case "ISHPublication":
					return ISHModulesNS.DocumentObj20ServiceReference.ISHType.ISHPublication;
				case "ISHReusedObj":
					return ISHModulesNS.DocumentObj20ServiceReference.ISHType.ISHReusedObj;
				case "ISHTemplate":
					return ISHModulesNS.DocumentObj20ServiceReference.ISHType.ISHTemplate;
				default:
					return ISHModulesNS.DocumentObj20ServiceReference.ISHType.ISHNone;
			}

		}


		public object _ResetRecursionHash()
		{
			m_recursion_hash.Clear();
			return true;
		}

		public bool ChangeRoleRecursivelybyBaseline(string PubGUID, string PubVer, string NewPerson, string Role, string SubGUID = "", string Resolution = "")
		{
			//if we were provided a SubGUID, that should be our startingpoint.  otherwise, our starting point is the mastermap.
			string GUID = "";
			if (SubGUID.Length > 0) {
				GUID = SubGUID;
			} else {
				GUID = GetMasterMapGUID(PubGUID, PubVer);
			}
			//Get the baseline dictionary
			Dictionary<string, CMSObject> dictBaseLineInfo = null;
			dictBaseLineInfo = GetBaselineObjects(PubGUID, PubVer);
			ChangeRecursiveSubRoutine(ref dictBaseLineInfo, GUID, NewPerson, Role, Resolution);
			return true;
		}
		private bool ChangeRecursiveSubRoutine(ref Dictionary<string, CMSObject> dictBaseLineInfo, string GUID, string NewAuthor, string Role, string Resolution = "")
		{
			//Dim requestmetadata As StringBuilder = BuildRequestedMetadata()
			string RequestedXMLObject = "";
			XmlDocument doc = new XmlDocument();
			XmlDocument docorig = new XmlDocument();

			//Version is determined by the baseline.
			string version = null;
			CMSObject objCMSObject = new CMSObject("", "", "");
			dictBaseLineInfo.TryGetValue(GUID, out objCMSObject);
			if (objCMSObject.Version.Length > 0) {
				version = objCMSObject.Version;
			} else {
				//failed to find object in the baseline
				modErrorHandler.Errors.PrintMessage(2, "Failed to find object in baseline. Info: " + GUID + ".", strModuleName + "-ChangeRecursiveSubRoutine");
				return false;
			}
			try {
				oISHAPIObjs.ISHDocObj.GetMetaData(GUID, ref version, "en", Resolution, oCommonFuncs.BuildRequestedMetadata().ToString(), ref RequestedXMLObject);
			} catch (Exception ex) {
				//failed to get object
				modErrorHandler.Errors.PrintMessage(2, "Failed to get object in DB. Info: " + GUID + "=" + version + ". Message: " + ex.Message.ToString(), strModuleName + "-ChangeRecursiveSubRoutine");
				return false;
			}

			//Load the XML and get the metadata:
			doc.LoadXml(RequestedXMLObject);

			//keep the original for matching later.
			docorig.LoadXml(RequestedXMLObject);
			string IshType = oCommonFuncs.GetISHTypeFromMeta(doc);

			//Get the children and recurse if applicable (by type).
			switch (IshType) {
				case "ISHMasterDoc":
				case "ISHModule":
					//if a map or topic, get children
					//Dim CurMeta As Object = IshReports.GetReportedObjects(doc)
					Hashtable children = new Hashtable();
					GetReferencedModules(GUID, children, version, "en");
					foreach (DictionaryEntry childmodule in children) {
						if (!(childmodule.Value.GUID == GUID) & !m_recursion_hash.Contains(childmodule.Value.GUID)) {
							//Track the GUID we're diving into so that we don't accidently try to process it again later (endless looping).
							m_recursion_hash.Add(childmodule.Value.GUID, childmodule.Value.GUID);
							ChangeRecursiveSubRoutine(ref dictBaseLineInfo, childmodule.Value.GUID, NewAuthor, Role, childmodule.Value.Resolution);
						}
					}

					break;
				default:
					break;
				//Returned something that we don't need to parse for children...
				//Return False
			}

			//change owner by GUID Stuff
			XmlNode ishfields = doc.SelectSingleNode("//ishfields");
			XmlNode ishfieldsorig = docorig.SelectSingleNode("//ishfields");
			XmlNode funame = ishfields.SelectSingleNode("ishfield[@name='" + Role + "']");
			//If there's no currently assigned name to the role specified, we need to insert it.
			if (funame == null) {
				XmlDocument funamedoc = new XmlDocument();
				funamedoc.LoadXml("<funame><ishfield name='" + Role + "' level='lng'>" + NewAuthor + "</ishfield></funame>");
				funame = funamedoc.FirstChild;
				ishfields.AppendChild(doc.ImportNode(funame.FirstChild, true));
				funame = ishfields.LastChild;
			}
			funame.InnerText = NewAuthor;
			XmlNode ishver = ishfields.SelectSingleNode("ishfield[@name='VERSION']");
			ishfields.RemoveChild(ishver);
			//change the owner of the GUID at the specified ver
			switch (IshType) {
				case "ISHMasterDoc":
				case "ISHModule":
					//if a map or topic, update the guid with simple command
					try {
						oISHAPIObjs.ISHDocObj.SetMetaData(GUID, ref version, "en", "", ishfields.OuterXml.ToString(), RequestedXMLObject);
					} catch (Exception ex) {
						modErrorHandler.Errors.PrintMessage(2, "Unable to change assignee on object " + GUID + "=" + version + ". Message: " + ex.Message.ToString(), strModuleName + "-ChangeRecursiveSubRoutine");
						return false;
					}
					break;
				case "ISHIllustration":
					if (!(Role == "FEDITOR") & !(Role == "FCODEREVIEWER")) {
						//If illustration, need to update all possible resolutions.
						//For Images, we need to replace ALL instances which means we need to trick it into thinking that we've returned the matching resolution for each type.

						//Start by making the original res match the High resolution.
						XmlNode res = doc.SelectSingleNode("//ishfields/ishfield[@name='FRESOLUTION']");
						XmlNode resOrig = docorig.SelectSingleNode("//ishfields/ishfield[@name='FRESOLUTION']");
						res.InnerText = "High";
						resOrig.InnerText = "High";


						try {
							oISHAPIObjs.ISHDocObj.SetMetaData(GUID, ref version, "en", "High", ishfields.OuterXml.ToString(), ishfieldsorig.OuterXml.ToString());

						} catch (Exception ex) {
						}
						res.InnerText = "Low";
						resOrig.InnerText = "Low";
						try {
							oISHAPIObjs.ISHDocObj.SetMetaData(GUID, ref version, "en", "Low", ishfields.OuterXml.ToString(), ishfieldsorig.OuterXml.ToString());

						} catch (Exception ex) {
						}
						res.InnerText = "Thumbnail";
						resOrig.InnerText = "Thumbnail";
						try {
							oISHAPIObjs.ISHDocObj.SetMetaData(GUID, ref version, "en", "Thumbnail", ishfields.OuterXml.ToString(), ishfieldsorig.OuterXml.ToString());

						} catch (Exception ex) {
						}
						res.InnerText = "Source";
						resOrig.InnerText = "Source";
						try {
							oISHAPIObjs.ISHDocObj.SetMetaData(GUID, ref version, "en", "Source", ishfields.OuterXml.ToString(), ishfieldsorig.OuterXml.ToString());

						} catch (Exception ex) {
						}
					}
					break;
				default:
					break;
				//Something else altogether... Not sure what to do here.
			}
			//Track the GUID we're diving into so that we don't accidently try to process it again later (endless looping).
			if (!m_recursion_hash.Contains(GUID)) {
				m_recursion_hash.Add(GUID, GUID);
			}
			return true;
		}
		public bool ChangeAssigneeRecursively(string GUID, string Version, string NewAuthor, string Role, string Resolution = "")
		{
			//Dim requestmetadata As StringBuilder = BuildRequestedMetadata()
			string RequestedXMLObject = "";
			XmlDocument doc = new XmlDocument();
			XmlDocument docorig = new XmlDocument();
			try {
				oISHAPIObjs.ISHDocObj.GetMetaData(GUID, ref Version, "en", Resolution, oCommonFuncs.BuildRequestedMetadata().ToString(), ref RequestedXMLObject);
			} catch (Exception ex) {
				//failed to get object
				modErrorHandler.Errors.PrintMessage(2, "Failed to get object in DB. Info: " + GUID + "=" + Version + ". Message: " + ex.Message.ToString(), strModuleName + "-ChangeAssigneeRecursively");
				return false;
			}

			//Load the XML and get the metadata:
			doc.LoadXml(RequestedXMLObject);

			//keep the original for matching later.
			docorig.LoadXml(RequestedXMLObject);
			string IshType = oCommonFuncs.GetISHTypeFromMeta(doc);

			//Get the children and recurse if applicable (by type).
			switch (IshType) {
				case "ISHMasterDoc":
				case "ISHModule":
					//if a map or topic, get children
					//Dim CurMeta As Object = IshReports.GetReportedObjects(doc)
					Hashtable children = new Hashtable();
					GetReferencedModules(GUID, children, Version, "en");
					foreach (DictionaryEntry childmodule in children) {
						if (!(childmodule.Value.GUID == GUID) & !m_recursion_hash.Contains(childmodule.Value.GUID)) {
							//Track the GUID we're diving into so that we don't accidently try to process it again later (endless looping).
							m_recursion_hash.Add(childmodule.Value.GUID, childmodule.Value.GUID);
							ChangeAssigneeRecursively(childmodule.Value.GUID, childmodule.Value.version, NewAuthor, Role, childmodule.Value.Resolution);
						}
					}

					break;
				default:
					break;
				//Returned something that we don't need to parse for children...
				//Return False
			}

			//change owner by GUID Stuff
			XmlNode ishfields = doc.SelectSingleNode("//ishfields");
			XmlNode ishfieldsorig = docorig.SelectSingleNode("//ishfields");
			XmlNode funame = ishfields.SelectSingleNode("ishfield[@name='" + Role + "']");
			//If there's no currently assigned name to the role specified, we need to insert it.
			if (funame == null) {
				XmlDocument funamedoc = new XmlDocument();
				funamedoc.LoadXml("<funame><ishfield name='" + Role + "' level='lng'>" + NewAuthor + "</ishfield></funame>");
				funame = funamedoc.FirstChild;
				ishfields.AppendChild(doc.ImportNode(funame.FirstChild, true));
				funame = ishfields.LastChild;
			}
			funame.InnerText = NewAuthor;
			XmlNode ishver = ishfields.SelectSingleNode("ishfield[@name='VERSION']");
			ishfields.RemoveChild(ishver);
			//change the owner of the GUID at the specified ver
			switch (IshType) {
				case "ISHMasterDoc":
				case "ISHModule":
					if (!(Role == "FILLUSTRATOR")) {
						//if a map or topic, update the guid with simple command
						try {
							oISHAPIObjs.ISHDocObj.SetMetaData(GUID, ref Version, "en", "", ishfields.OuterXml.ToString(), RequestedXMLObject);
						} catch (Exception ex) {
							modErrorHandler.Errors.PrintMessage(2, "Unable to change assignee on object " + GUID + "=" + Version + ". Message: " + ex.Message.ToString(), strModuleName + "-ChangeAssigneeRecursively");
							return false;
						}
					}
					break;
				case "ISHIllustration":
					if (!(Role == "FEDITOR") & !(Role == "FCODEREVIEWER")) {
						//If illustration, need to update all possible resolutions.
						//For Images, we need to replace ALL instances which means we need to trick it into thinking that we've returned the matching resolution for each type.

						//Start by making the original res match the High resolution.
						XmlNode res = doc.SelectSingleNode("//ishfields/ishfield[@name='FRESOLUTION']");
						XmlNode resOrig = docorig.SelectSingleNode("//ishfields/ishfield[@name='FRESOLUTION']");
						res.InnerText = "High";
						resOrig.InnerText = "High";


						try {
							oISHAPIObjs.ISHDocObj.SetMetaData(GUID, ref Version, "en", "High", ishfields.OuterXml.ToString(), ishfieldsorig.OuterXml.ToString());

						} catch (Exception ex) {
						}
						res.InnerText = "Low";
						resOrig.InnerText = "Low";
						try {
							oISHAPIObjs.ISHDocObj.SetMetaData(GUID, ref Version, "en", "Low", ishfields.OuterXml.ToString(), ishfieldsorig.OuterXml.ToString());

						} catch (Exception ex) {
						}
						res.InnerText = "Thumbnail";
						resOrig.InnerText = "Thumbnail";
						try {
							oISHAPIObjs.ISHDocObj.SetMetaData(GUID, ref Version, "en", "Thumbnail", ishfields.OuterXml.ToString(), ishfieldsorig.OuterXml.ToString());

						} catch (Exception ex) {
						}
						res.InnerText = "Source";
						resOrig.InnerText = "Source";
						try {
							oISHAPIObjs.ISHDocObj.SetMetaData(GUID, ref Version, "en", "Source", ishfields.OuterXml.ToString(), ishfieldsorig.OuterXml.ToString());

						} catch (Exception ex) {
						}
					}
					break;
				default:
					break;
				//Something else altogether... Not sure what to do here.
			}
			//Track the GUID we're diving into so that we don't accidently try to process it again later (endless looping).
			if (!m_recursion_hash.Contains(GUID)) {
				m_recursion_hash.Add(GUID, GUID);
			}
			return true;
		}

		public bool CanMoveToState(string strState, string strGUID, string strVersion, string strResolution, string strLanguage = "en")
		{
			bool functionReturnValue = false;

			string[] OutStates = {
				
			};
			try {
				// Declare variable for the Application service
				//Dim DocService As ISDoc.DocumentObj20 = New ISDoc.DocumentObj20()

				// Clear variable for the result
				oISHAPIObjs.ISHDocObj.GetPossibleTransitionStates(strGUID, ref strVersion, strLanguage, strResolution, ref OutStates);

			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(2, "Error getting possible transition states for " + strGUID + "" + strVersion + "" + strLanguage + ". Message: " + ex.Message.ToString(), strModuleName + "-CanMoveToState");
				return false;
			}

			// Now check to see if we can move to desired state:
			if ((OutStates.Length > 0)) {
				string s = null;
				foreach (string s_loopVariable in OutStates) {
					s = s_loopVariable;
					if ((s == strState)) {
						functionReturnValue = true;
						return functionReturnValue;
					}
				}
			}
			return false;
			return functionReturnValue;
		}


		public string GetCurrentState(string GUID, string Version, string strResolution, string Language = "en")
		{

			string state = "nothing";
			string OutXML = "";
			try {
				//' Declare variable for the Application service
				//Dim DocService As IshDocument.DocumentObj20 = New ISDoc.DocumentObj20()

				string strMeta = "<ishfields><ishfield name=\"FSTATUS\" level=\"lng\"/></ishfields>";

				oISHAPIObjs.ISHDocObj.GetMetaData(GUID, ref Version, Language, strResolution, strMeta, ref OutXML);

			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(2, "Error getting current state for " + GUID + "" + Version + "" + Language + ". Message: " + ex.Message.ToString(), strModuleName + "-GetCurrentState");
				return false;
			}

			string strFind = "<ishfield name=\"FSTATUS\" level=\"lng\">";
			state = OutXML.Substring(OutXML.LastIndexOf(strFind) + strFind.Length);
			state = state.Remove(state.LastIndexOf("</ishfield>"));
			return state;
		}
		#endregion
	}
}
