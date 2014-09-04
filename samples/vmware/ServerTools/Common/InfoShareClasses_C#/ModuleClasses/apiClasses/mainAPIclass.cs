using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Xml;
using System.IO;
using System.Xml.XmlReader;
using System.Text;

using Microsoft.VisualBasic.ControlChars;
using Microsoft.VisualBasic.FileIO;
using Microsoft.VisualBasic.FileIO.FileSystem;
using System.Convert;
using System.Text.RegularExpressions;
using ErrorHandlerNS;
namespace ISHModulesNS
{

	public class mainAPIclass
	{
		//This class contains:
		//1. An explicit object that hooks up all of the default APIs as objects
		//2. Shared methods used across multiple subclasses.
		//All sub API classes inherit this class to include the mutually used above two pieces.
		//Use of the subclasses must be done explictly.  All universally needed functionality
		//should be moved out to the instantiated class and shared there, not in these subs.

		#region "Private Members"
			#endregion
		private readonly string strModuleName = "CustomCMSFuncs";
		#region "Constructors"
		public mainAPIclass()
		{
			//Can be overridden by subs. Should report an error if ever called.
			//Do nothing.
		}

		#endregion
		#region "Properties"
		public ISHObjs oISHAPIObjs;
		public string Context = "";
		public clsCommonFuncs oCommonFuncs = new clsCommonFuncs();
		public ArrayList DeletedGUIDs = new ArrayList();

		public ArrayList DeleteFailedGUIDs = new ArrayList();
		/// <summary>
		/// Structure used to store common metadata from returned CMS objects.
		/// </summary>
		public struct ObjectData
		{
			public string GUID;
			public string IshType;
			public string Title;
			public string Version;
			public string Language;
			public string Status;
			public string Resolution;
			public XmlNode MetaData;
		}

		#endregion
		#region "Shared Methods"
		#region "Baseline Funcs"
		public Dictionary<string, CMSObject> GetBaselineObjects(string GUID, string Version, string Language = "en")
		{
			//Get the pub's baseline ID from the pub object
			string outObjectList = "";
			string[] GUIDs = new string[1];
			GUIDs[0] = GUID;
			string[] Languages = new string[1];
			Languages[0] = Language;
			string[] Resolutions = null;
			//[UPGRADE to 2013SP1] Changed the result to return the result instead of just some arbitrary response.
			Resolutions = oISHAPIObjs.ISHMetaObj.GetLOVValues("DRESOLUTION");

			//Get the existing publication content at the specified version.
			outObjectList = oISHAPIObjs.ISHPubOutObj25.RetrieveVersionMetadata(GUIDs, Version, oCommonFuncs.BuildMinPubMetadata().ToString());

			XmlDocument VerDoc = new XmlDocument();
			VerDoc.LoadXml(outObjectList);
			if (VerDoc == null | VerDoc.HasChildNodes == false) {
				modErrorHandler.Errors.PrintMessage(3, "Unable to find publication for specified GUID: " + GUID, strModuleName + "-GetBaselineObjects");
				return null;
			}
			//Get the Baseline ID from the publication:
			string baselineID = "";
			string baselinename = null;
			XmlNode ishfields = VerDoc.SelectSingleNode("//ishfields");
			baselinename = ishfields.SelectSingleNode("ishfield[@name='FISHBASELINE']").InnerText;
			//Pull the baseline info
			string myBaseline = "";
			//[UPGRADE to 2013SP1] Changed the result to return the result instead of just some arbitrary response.
			baselineID = oISHAPIObjs.ISHBaselineObj.GetBaselineId(baselinename);
			//[UPGRADE to 2013SP1] Changed the result to return the result instead of just some arbitrary response.
			outObjectList = oISHAPIObjs.ISHBaselineObj.GetReport(baselineID, null, Languages, Languages, Languages, Resolutions);
			//Load the resulting baseline string as an xml document
			XmlDocument baselineDoc = new XmlDocument();
			baselineDoc.LoadXml(outObjectList);
			Dictionary<string, CMSObject> dictBaselineObjects = new Dictionary<string, CMSObject>();
			//for each object referenced, store the various info in an object and then store them in the hashtable. GUIDs are the keys.
			foreach (XmlNode baselineObject in baselineDoc.SelectNodes("/baseline/objects/object")) {
				//create a new CMSObject storage container
				string refGuid = baselineObject.Attributes.GetNamedItem("ref").Value;
				string ishtype = baselineObject.Attributes.GetNamedItem("type").Value;
				if (ishtype == "ISHNone") {
					continue;
				}
				string refver = baselineObject.Attributes.GetNamedItem("versionnumber").Value;
				string reportitems = baselineObject.SelectSingleNode("reportitems").OuterXml.ToString();
				CMSObject CMSObject = new CMSObject(refGuid, refver, ishtype, reportitems);
				//save the object to the hash using the GUID as the key.
				dictBaselineObjects.Add(refGuid, CMSObject);
			}
			return dictBaselineObjects;
		}
		#endregion

		#region "Condition Functions"


		#endregion

		#region "Document Functions"
		public bool UpdateTitleProperty(string GUID, string Version)
		{

			//Get object the XML
			XmlDocument ObjectXML = null;
			ObjectXML = GetObjByID(GUID, Version, "en", "");
			//get the topic xml out of the CDATA from the object
			XmlDocument topicXML = new XmlDocument();
			topicXML = oCommonFuncs.GetXMLOut(ObjectXML.SelectSingleNode("//ishdata"));
			string ENTitle = topicXML.SelectSingleNode("/*/title").InnerText;
			ENTitle = ENTitle.Replace(",", "‚");
			ENTitle = ENTitle.Replace("\\", "/");
			ENTitle = Strings.Replace(ENTitle, "*", "");
			ENTitle = Strings.Replace(ENTitle, "?", "");
			ENTitle = Strings.Replace(ENTitle, ">", "");
			ENTitle = Strings.Replace(ENTitle, "<", "");
			ENTitle = Strings.Replace(ENTitle, ":", "");
			ENTitle = Strings.Replace(ENTitle, "|", "");
			ENTitle = Strings.Replace(ENTitle, "#", "");
			ENTitle = Strings.Replace(ENTitle, "!", "");
			//Get the title out of the XML
			string strMetaTitle = "<ishfields><ishfield name=\"FTITLE\" level=\"logical\">" + ENTitle + "</ishfield></ishfields>";
			//Set the title on the object
			if (SetMeta(strMetaTitle, GUID, Version, "")) {
				return true;
			} else {
				return false;
			}

		}

		public bool SetMeta(string strMeta, string GUID, string Version, string strResolution, string Language = "en")
		{
			try {
				// Clear variable for the result
				oISHAPIObjs.ISHDocObj.SetMetaData(GUID, ref Version, Language, strResolution, strMeta, "");
			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(2, "Error setting meta for " + GUID + "" + Version + "" + Language + ". Message: " + ex.Message.ToString(), strModuleName + "-SetMeta");
				return false;
			}
			return true;
		}

		/// <summary>
		/// Pulls the specified object from the CMS and saves it at the specified location.  Returns true if successful and file has been saved.
		/// </summary>
		public bool GetObjByID(string GUID, string Version, string Language, string Resolution, string SavePath)
		{
			XmlNode MyNode = null;
			XmlDocument MyDoc = new XmlDocument();
			XmlDocument MyMeta = new XmlDocument();
			string XMLString = "";
			string ISHMeta = "";
			string ISHResult = "";
			string filename = "BROKEN_FILENAME";
			string extension = "FIX";
			StringBuilder requestedmeta = oCommonFuncs.BuildRequestedMetadata();
			//Call the CMS to get our content!
			try {
				ISHResult = oISHAPIObjs.ISHDocObj.GetDocObj(GUID, ref Version, Language, Resolution, "", "", requestedmeta.ToString(), ref XMLString);
			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(3, "Failed to retrieve XML from CMS server: " + ex.Message, strModuleName + "-GetObjByID");
				return false;
			}

			//Load the XML and get the metadata:
			try {
				MyDoc.LoadXml(XMLString);
				filename = oCommonFuncs.GetFilenameFromIshMeta(MyDoc);
				//Remove any characters not allowed by windows operating system on filenames.
				filename = filename.Replace("\\", "");
				filename = filename.Replace("/", "");
				filename = filename.Replace(":", "");
				filename = filename.Replace("*", "");
				filename = filename.Replace("?", "");
				filename = filename.Replace("\"", "");
				filename = filename.Replace("<", "");
				filename = filename.Replace(">", "");
				filename = filename.Replace("|", "");
				MyNode = MyDoc.SelectSingleNode("//ishdata");
				//Get the extension:
				foreach (XmlAttribute ishattrib in MyNode.Attributes) {
					if (ishattrib.Name == "fileextension") {
						extension = ishattrib.Value;
					}
				}
			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(3, "Failed to retrieve object metadata from returned XML: " + ex.Message, strModuleName);
				return false;
			}
			//check to see if it's already been exported first, exit if it has...
			if (File.Exists(SavePath + "\\" + filename + "." + extension.ToLower())) {
				modErrorHandler.Errors.PrintMessage(2, "File already exists. Skipping: " + SavePath + "\\" + filename + "." + extension.ToLower(), strModuleName);
				return true;
			}
			//Convert the CDATA to byte array
			byte[] finalfile = null;
			try {
				//Convert CDATA Blob to Byte array
				finalfile = oCommonFuncs.GetBinaryOut(MyNode);
			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(3, "Failed to convert CDATA Blob to binary stream - no content returned from CMS: " + ex.Message, strModuleName);
				return false;
			}
			//Save the content out to a file:
			try {
				//Create the save path if it doesn't exist:
				if (Directory.Exists(SavePath) == false) {
					Directory.CreateDirectory(SavePath);
				}
				//write to filename, bytes we extracted, don't append

				ISHModulesNS.My.MyProject.Computer.FileSystem.WriteAllBytes(SavePath + "\\" + filename + "." + extension.ToLower(), finalfile, false);
				return true;
			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(3, "Failed to save returned object to a file: " + ex.Message, strModuleName);
				return false;
			}
		}
		/// <summary>
		/// Pulls the specified object from the CMS and saves it at the specified location.  Returns true if successful and file has been saved.
		/// </summary>
		public XmlDocument GetObjByID(string GUID, string Version, string Language, string Resolution)
		{
			XmlNode MyNode = null;
			XmlDocument MyDoc = new XmlDocument();
			XmlDocument MyMeta = new XmlDocument();
			string XMLString = "";
			string ISHMeta = "";
			string ISHResult = "";
			string filename = "BROKEN_FILENAME";
			string extension = "FIX";
			StringBuilder requestedmeta = oCommonFuncs.BuildRequestedMetadata();
			//Call the CMS to get our content!
			try {
				ISHResult = oISHAPIObjs.ISHDocObj.GetDocObj(GUID, ref Version, Language, Resolution, "", "", requestedmeta.ToString(), ref XMLString);
			} catch (Exception ex) {
				//modErrorHandler.Errors.PrintMessage(3, "Failed to retrieve XML from CMS server: " + ex.Message, strModuleName)
				return null;
			}

			//Load the XML and get the metadata:
			try {
				MyDoc.LoadXml(XMLString);
				return MyDoc;
				//filename = GetFilenameFromIshMeta(MyDoc)
				//MyNode = MyDoc.SelectSingleNode("//ishdata")
				//'Get the extension:
				//For Each ishattrib As XmlAttribute In MyNode.Attributes
				//    If ishattrib.Name = "fileextension" Then
				//        extension = ishattrib.Value
				//    End If
				//Next
			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(3, "Failed to retrieve object metadata from returned XML: " + ex.Message, strModuleName);
				return null;
			}
		}
		private bool ObliterateGUID(string GUID, string Version, string Language, string Resolution = "High")
		{
			try {
				if (Resolution.Length > 0) {
					try {
						oISHAPIObjs.ISHDocObj.Delete(GUID, Version, Language, "Low");
						modErrorHandler.Errors.PrintMessage(1, "Deleting low resolution for: " + GUID, strModuleName + "-ObliterateGUID");


					} catch (Exception ex) {
					}
					if (Resolution == "High") {
						try {
							oISHAPIObjs.ISHDocObj.Delete(GUID, Version, Language, "High");
							modErrorHandler.Errors.PrintMessage(1, "Deleting high resolution for: " + GUID, strModuleName + "-ObliterateGUID");


						} catch (Exception ex) {
						}
					}

					try {
						oISHAPIObjs.ISHDocObj.Delete(GUID, Version, Language, "Thumbnail");
						modErrorHandler.Errors.PrintMessage(1, "Deleting thumbnail for: " + GUID, strModuleName + "-ObliterateGUID");


					} catch (Exception ex) {
					}

					try {
						oISHAPIObjs.ISHDocObj.Delete(GUID, Version, Language, "Source");
						modErrorHandler.Errors.PrintMessage(1, "Deleting source resolution for: " + GUID, strModuleName + "-ObliterateGUID");


					} catch (Exception ex) {
					}

				}
				try {
					oISHAPIObjs.ISHDocObj.Delete(GUID, Version, "", "");
					modErrorHandler.Errors.PrintMessage(1, "Deleting language level for: " + GUID, strModuleName + "-ObliterateGUID");


				} catch (Exception ex) {
				}
				try {
					oISHAPIObjs.ISHDocObj.Delete(GUID, "", "", "");
					modErrorHandler.Errors.PrintMessage(1, "Deleting version level for: " + GUID, strModuleName + "-ObliterateGUID");


				} catch (Exception ex) {
				}
				//Deleted everything we could!  Let's see if we did it.
				if (ObjectExists(GUID, Version, Language, Resolution)) {
					modErrorHandler.Errors.PrintMessage(2, "Deletions performed on GUID failed. GUID: " + GUID + " still exists. Delete manually.", strModuleName + "-ObliterateGUID");
					DeleteFailedGUIDs.Add(GUID);

					return false;
				}
				DeletedGUIDs.Add(GUID);
				return true;
			} catch (Exception ex) {
				if (ex.Message.Contains("-115")) {
					//The module is referenced by some other module...
					modErrorHandler.Errors.PrintMessage(3, "Unable to delete module: " + GUID + " Referenced by other module(s).", strModuleName);
					return false;
				} else if (ex.Message.Contains("-102")) {
					oISHAPIObjs.ISHDocObj.Delete(GUID, "", "", "");
				} else {
					return false;
				}

			}

		}
		/// <summary>
		/// Check to see if an object with the specified parameters exists in the CMS.  Returns true if exists.
		/// </summary>
		public bool ObjectExists(string GUID, string Version, string Language, string Resolution = "")
		{
			XmlNode MyNode = null;
			XmlDocument MyDoc = new XmlDocument();
			XmlDocument MyMeta = new XmlDocument();
			string XMLString = "";
			string ISHMeta = "";
			string ISHResult = "";
			string filename = "BROKEN_FILENAME";
			string extension = "FIX";
			StringBuilder requestedmeta = oCommonFuncs.BuildRequestedMetadata();

			//Call the CMS to get our content!
			try {
				ISHResult = oISHAPIObjs.ISHDocObj.GetDocObj(GUID, ref Version, Language, Resolution, "", "", requestedmeta.ToString(), ref XMLString);
				return true;
			} catch (Exception ex) {
				//modErrorHandler.Errors.PrintMessage(3, "Failed to retrieve XML from CMS server: " + ex.Message, strModuleName)
				return false;
			}

		}

		/// <summary>
		/// Replaces a specified module with templated content.  
		/// Most commonly used before attempting to recursively delete referencing modules to prevent circular references.
		/// </summary>
		public bool ReplaceWithTemplatedContent(string GUID, string Version, string Language)
		{
			//Set the GUIDs in our templates.
			oCommonFuncs.SetGUIDinTemplates(GUID);

			//Start the replacement process:
			StringBuilder requestedmetadata = oCommonFuncs.BuildRequestedMetadata();
			string RequestedXMLObject = "";
			XmlDocument doc = new XmlDocument();
			string IshType = null;
			string TopicType = null;
			byte[] Data = null;

			try {
				//check out the module (must be map or topic)
				oISHAPIObjs.ISHDocObj.CheckOut(GUID, ref Version, Language, "", requestedmetadata.ToString(), ref RequestedXMLObject);
			} catch (Exception ex) {
				//If, for some reason, we already have an object checked out, great.  otherwise, we can't check it out for some reason.
				//Exit Code for already checking an object out is -132
				if (ex.Message.Contains("-132")) {
					//we have it checked out already, but we still need to get the object CData:
					oISHAPIObjs.ISHDocObj.GetDocObj(GUID, ref Version, Language, "", "", "", requestedmetadata.ToString(), ref RequestedXMLObject);
				} else {
					modErrorHandler.Errors.PrintMessage(3, "Unable to checkout GUID: " + GUID + " Error: " + ex.Message, strModuleName + "-ReplaceWithTemplatedContent");
					return false;
				}
			}

			//Load the XML and get the metadata:
			doc.LoadXml(RequestedXMLObject);
			//get the ISHType from the meta
			IshType = oCommonFuncs.GetISHTypeFromMeta(doc);
			switch (IshType) {
				case "ISHMasterDoc":
					//if a map, replace the content with our template content
					Data = oCommonFuncs.XMLTemplates.mapblob;

					break;
				case "ISHModule":
					//if a topic, find out what kind
					TopicType = oCommonFuncs.GetTopicTypeFromMeta(doc);
					switch (TopicType) {
						case "task":
							Data = oCommonFuncs.XMLTemplates.taskblob;
							break;
						case "concept":
							Data = oCommonFuncs.XMLTemplates.conceptblob;
							break;
						case "reference":
							Data = oCommonFuncs.XMLTemplates.referenceblob;
							break;
						case "troubleshooting":
							Data = oCommonFuncs.XMLTemplates.troubleshootingblob;
							break;
						default:
							modErrorHandler.Errors.PrintMessage(2, "Unexpected TopicType used for GUID: " + GUID, strModuleName + "-ReplaceWithTemplatedContent");
							return false;
					}
					break;
				//replace the content with our template content
				default:
					//Returned an unexpected type...
					//modErrorHandler.Errors.PrintMessage(2, "Unable to determine ISHType for GUID: " + GUID, strModuleName + "-ReplaceWithTemplatedContent")
					return false;
			}
			if ((Data == null) == false) {
				try {
					oISHAPIObjs.ISHDocObj.CheckIn(GUID, Version, Language, "", "", "EDTXML", Data);
					return true;
				} catch (Exception ex) {
					modErrorHandler.Errors.PrintMessage(3, "Unable to checkin after replacing content. GUID: " + GUID + " Object is still checked out to user!", strModuleName + "-ReplaceWithTemplatedContent");
					return false;
				}
			} else {
				modErrorHandler.Errors.PrintMessage(3, "No Data blob was created from the template file.  Can't replace content in CMS with nothing.  Failed to replace GUID:" + GUID, strModuleName + "-ReplaceWithTemplatedContent");
				return false;
			}
		}

		/// <summary>
		/// Deletes a given object and referencing parents (if possible).  Optionally allows deleting all referenced children (does not apply to parents).
		/// </summary>
		public bool DeleteObjectRecursivelyByGUID(string GUID, string Version = "1", string Language = "en", string Resolution = "", bool DeleteSubs = true)
		{
			Hashtable ParentModules = new Hashtable();
			Hashtable ChildrenModules = new Hashtable();
			StringBuilder requestedmetadata = oCommonFuncs.BuildRequestedMetadata();
			//first, find out if the GUID exists in the CMS:

			if (ObjectExists(GUID, Version, Language, Resolution)) {
				//We're going to delete it anyway.  Use template to replace the contents completely (if not an image).
				if (string.IsNullOrEmpty(Resolution)) {
					if (ReplaceWithTemplatedContent(GUID, Version, Language) == false) {
						//modErrorHandler.Errors.PrintMessage(2, "Unable to replace content in GUID: " + GUID + " with default, template content. May not be able to delete due if the topic contains circular references to referencing modules.", strModuleName + "-RecursiveDeletion")
					}
				}


				//if the guid has owners, use the list to recurse into them
				if (GetReferencingModules(GUID, ref ParentModules, Version, Language, requestedmetadata.ToString())) {
					foreach (DictionaryEntry parentmodule in ParentModules) {
						// Make sure we don't recurse into ourselves here...
						if (!(parentmodule.Value.GUID == GUID)) {
							bool result = DeleteObjectRecursivelyByGUID(parentmodule.Value.GUID, parentmodule.Value.Version, parentmodule.Value.Language, parentmodule.Value.Resolution, false);
							if (result == false) {
								//There was a problem deleting the parent!
								modErrorHandler.Errors.PrintMessage(3, "Unable to delete parent of GUID: " + GUID + " Parent that returned an error was GUID: " + parentmodule.Value.GUID, strModuleName + "-RecursiveDeletion");
								return false;
							}
						}
					}
				} else {
					//otherwise, if it doesn't, get the children
					//if true, has children
					if (GetReferencedModules(GUID, ref ChildrenModules, Version, Language, requestedmetadata.ToString())) {
						//then, delete the current GUID (if it can be deleted), 
						if (CanBeDeleted(GUID, Version, Language, Resolution)) {
							ObliterateGUID(GUID, Version, Language, Resolution);
						} else {
							//Can't be deleted for some reason...
							modErrorHandler.Errors.PrintMessage(3, "Unable to delete GUID: " + GUID, strModuleName + "RecursiveDeletion");
							return false;
						}
						//check to see if we should also delete children of the current module
						if (DeleteSubs == true) {
							//recursedelete into each of the children
							foreach (DictionaryEntry childmodule in ChildrenModules) {
								//Make sure we don't recurse on ourselves
								if (!(childmodule.Value.GUID == GUID)) {
									if (DeleteObjectRecursivelyByGUID(childmodule.Value.GUID, childmodule.Value.Version, childmodule.Value.language, childmodule.Value.resolution, DeleteSubs) == false) {
										modErrorHandler.Errors.PrintMessage(3, "Failed to delete descendant of: " + GUID, strModuleName + "RecursiveDeletion");
										return false;
									}
								}
							}
						}
						return true;
					} else {
						//no children?  Just delete the current GUID
						if (CanBeDeleted(GUID, Version, Language)) {
							ObliterateGUID(GUID, Version, Language, Resolution);
							return true;
						} else {
							//Can't be deleted for some reason...
							modErrorHandler.Errors.PrintMessage(3, "Unable to delete GUID: " + GUID, strModuleName + "RecursiveDeletion");
							return false;
						}
					}
				}
				//Object existed but had parents before and needed children trimmed.  That's done.
				//Shouldn't have any parents at this point, so we can likely delete it.  let's try:
				if (CanBeDeleted(GUID, Version, Language)) {
					ObliterateGUID(GUID, Version, Language, Resolution);
					return true;
				} else {
					//Can't be deleted for some reason...
					modErrorHandler.Errors.PrintMessage(3, "Unable to delete GUID: " + GUID, strModuleName + "RecursiveDeletion");
					return false;
				}
				//made it this far - everything deleted successfully without returning False.
				return true;
			//End "if GUID exists"
			} else {
				//GUID doesn't exist.  Mission accomplished - it is already deleted!
				modErrorHandler.Errors.PrintMessage(1, "GUID doesn't exist in the CMS - no need to delete.  GUID: " + GUID, strModuleName + "-DeleteObjectRecursivelyByGUID");
				return true;
			}

		}
		/// <summary>
		/// Checks to see if a specified module has no referencing modules and is not in a released state. Returns true if both conditions are met.
		/// </summary>
		/// <param name="GUID">GUID of object in CMS</param>
		/// <param name="Version">Version of object in CMS</param>
		/// <param name="Language">Language of object in CMS</param>
		/// <param name="Resolution">Resolution of object in CMS</param>
		/// <returns>Boolean</returns>
		public bool CanBeDeleted(string GUID, string Version = "1", string Language = "en", string Resolution = "")
		{
			Hashtable ModuleHash = new Hashtable();

			if (GetReferencingModules(GUID, ModuleHash, Version, Language) == true) {
				//Has parents, can't be deleted
				return false;
			} else {
				//This object has no parents.  first check passed.  Continue

			}
			string SearchResult = "";
			oISHAPIObjs.ISHDocObj.GetMetaData(GUID, ref Version, Language, Resolution, "<ishfields><ishfield name=\"FSTATUS\" level=\"lng\"/></ishfields>", ref SearchResult);
			if (SearchResult.Contains("\"FSTATUS\" level=\"lng\">Released")) {
				//Status is released.  Can't delete
				return false;
			} else {
				//Status is something else, can delete
				return true;
			}




		}
		#endregion

		#region "Folder Functions"
		private bool FolderContainsGUID(string GUID, long FolderID)
		{
			if (!(FolderID == 0)) {
				string requestedobjects = "";


				StringBuilder RealRequestedMeta = new StringBuilder();
				RealRequestedMeta = oCommonFuncs.BuildRequestedMetadata();



				//Returns a list of referencing objects to "requestedobjects" string as xml:
				try {
					oISHAPIObjs.ISHFolderObj.GetContents(FolderID, "", RealRequestedMeta.ToString(), ref requestedobjects);
				} catch (Exception ex) {
					modErrorHandler.Errors.PrintMessage(2, "Unable to get contents of folder " + FolderID, strModuleName + "-FolderContainsGUID");
				}


				//Load the string as an xmldoc
				XmlDocument doc = new XmlDocument();
				doc.LoadXml(requestedobjects);
				//get our returned objects into a hashtable (key=GUID, Value=Information + XMLNode of returned data)
				Hashtable SubObjects = GetReportedObjects(doc);
				//Key is a mashup of the GUID, version, and other info so it's not going to match against our GUID... We'll need to spin through the hash to find what we want.
				foreach (DictionaryEntry myentry in SubObjects) {
					if (myentry.Value.GUID == GUID) {
						return true;
					}
				}
				//No matches found, exit.
				return false;
			} else {
				//can't do searches for content in root folder.
				return false;
			}


		}

		public long FindFolderIDforObjbyGUID(string GUID, long ParentFolderID)
		{
			long functionReturnValue = 0;
			//Recursively searches for a GUID and returns the ID when found.  returns -1 if not found.

			//First, check to see if our GUID exists in this folder (only if not root folder):
			if (FolderContainsGUID(GUID, ParentFolderID)) {
				//if we found it, return this as the valid parent!
				return ParentFolderID;
				return functionReturnValue;

			} else {
				//if not, we need to dive deeper.
				//get a folderlist of all children of currentfolderid
				string subfolderlistXML = "";
				string CMSReply = null;
				try {
					CMSReply = oISHAPIObjs.ISHFolderObj.GetSubFolders(ParentFolderID.ToString(), 1, ref subfolderlistXML);

				} catch (Exception ex) {
				}
				//Load the subfolderlistXML into an xml document
				// Create the reader from the string.
				StringReader strReader = new StringReader(subfolderlistXML);

				XmlReader reader = XmlReader.Create(strReader);
				// Create the new XMLdoc and load the content into it.
				XmlDocument doc = new XmlDocument();
				doc.PreserveWhitespace = true;
				doc.Load(reader);

				//recurse into each subid
				long ReturnedFolderID = 0;
				try {
					foreach (XmlNode subid in doc.SelectNodes("//ishfolders/ishfolder/@ishfolderref")) {
						ReturnedFolderID = FindFolderIDforObjbyGUID(GUID, Convert.ToInt64(subid.Value.ToString()));
						if (ReturnedFolderID > 0) {
							break; // TODO: might not be correct. Was : Exit For
						}
					}
					return ReturnedFolderID;
				} catch (Exception ex) {
					// No Sub Folders.  
					return -1;
				}
			}
			return functionReturnValue;
		}
		#endregion

		#region "Meta Functions"

		#endregion

		#region "Output Functions"

		#endregion

		#region "Pub Functions"
		public string GetLatestPubVersionNumber(string GUID)
		{
			string outObjectList = "";
			string[] GUIDs = new string[1];
			GUIDs[0] = GUID;

			//Get the existing content at the latest version.
			//[Upgraded to LCA2013]
			outObjectList = oISHAPIObjs.ISHPubOutObj25.RetrieveVersionMetadata(GUIDs, "latest", oCommonFuncs.BuildMinPubMetadata().ToString());
			XmlDocument VerDoc = new XmlDocument();
			VerDoc.LoadXml(outObjectList);
			if (VerDoc == null | VerDoc.HasChildNodes == false) {
				modErrorHandler.Errors.PrintMessage(3, "Unable to find publication for specified GUID: " + GUID, strModuleName + "-GetLatestPubVersionNumber");
				return null;
			}
			//Iterate through each returned obj and figure out the highest number value for 'Version returned'
			XmlNode ishfields = VerDoc.SelectSingleNode("//ishfields");
			XmlNode VersionNode = ishfields.SelectSingleNode("ishfield[@name='VERSION']");
			return VersionNode.InnerText;
		}

		public string GetMasterMapGUID(string PubGUID, string PubVer)
		{
			//Get the pub's mastermap GUID from the pub object
			string outObjectList = "";
			string[] GUIDs = new string[1];
			GUIDs[0] = PubGUID;
			string[] Languages = new string[1];
			Languages[0] = "en";
			string[] Resolutions = null;
			//[UPGRADE to 2013SP1] Changed the result to return the result instead of just some arbitrary response.
			Resolutions = oISHAPIObjs.ISHMetaObj.GetLOVValues("DRESOLUTION");

			//Get the existing publication content at the specified version.
			//[UPGRADE to 2013SP1] Changed the result to return the result instead of just some arbitrary response.
			outObjectList = oISHAPIObjs.ISHPubOutObj25.RetrieveVersionMetadata(GUIDs, PubVer, oCommonFuncs.BuildMinPubMetadata().ToString());

			XmlDocument VerDoc = new XmlDocument();
			VerDoc.LoadXml(outObjectList);
			if (VerDoc == null | VerDoc.HasChildNodes == false) {
				modErrorHandler.Errors.PrintMessage(3, "Unable to find publication for specified GUID: " + PubGUID, strModuleName + "-GetMasterMapGUID");
				return null;
			}
			//Get the Master Map GUID from the publication:
			XmlNode ishfields = VerDoc.SelectSingleNode("//ishfields");
			return ishfields.SelectSingleNode("ishfield[@name='FISHMASTERREF']").InnerText;
		}
		#endregion

		#region "PubOutput Functions"

		#endregion

		#region "Reports Functions"


		/// <summary>
		/// Given a commonly returned "ishobjects" XML Document returned from most CMS queries, 
		/// this function converts each entry to Dictionary Entries in a hashtable.  Each entry 
		/// uses a combined GUID+meta key that appears the same as the commonly used CMS filenames. 
		/// The value is the ObjectData structure found in this class.
		/// </summary>
		/// <param name="XMLDoc">A standard "ishobjects" XML Document returned from a query to the CMS.</param>
		/// <returns>A Hashtable containing all of the ishobjects and their metadata.</returns>
		public Hashtable GetReportedObjects(XmlDocument XMLDoc)
		{
			//Grab each returned ishobject and build our structure out of it.  Add it to the hash we'll be returning.
			Hashtable ReturnHash = new Hashtable();
			foreach (XmlNode CMSModule in XMLDoc.SelectNodes("/ishobjects/ishobject")) {
				ObjectData returnobj = new ObjectData();
				var _with1 = returnobj;
				_with1.GUID = CMSModule.Attributes.GetNamedItem("ishref").Value.ToString();
				_with1.IshType = CMSModule.Attributes.GetNamedItem("ishtype").Value.ToString();
				if (_with1.IshType == "ISHNotFound") {
					//Ack, it's a broken link ishobject! skip it here.  
					//TODO: would need to rewrite this if we want to track the broken links for anything.
					continue;
				}
				XmlNode ishfields = CMSModule.SelectSingleNode("ishfields");

				_with1.Title = ishfields.SelectSingleNode("ishfield[contains(@name, 'FTITLE')]").InnerText;
				_with1.Version = ishfields.SelectSingleNode("ishfield[contains(@name, 'VERSION')]").InnerText;
				_with1.Language = ishfields.SelectSingleNode("ishfield[contains(@name, 'DOC-LANGUAGE')]").InnerText;
				_with1.Status = ishfields.SelectSingleNode("ishfield[contains(@name, 'FSTATUS')]").InnerText;
				if (_with1.IshType == "ISHIllustration") {
					_with1.Resolution = ishfields.SelectSingleNode("ishfield[contains(@name, 'FRESOLUTION')]").InnerText;
				} else {
					_with1.Resolution = "";
				}
				_with1.MetaData = CMSModule;
				if (ReturnHash.Contains(returnobj.Title + "=" + returnobj.GUID + "=" + returnobj.Version + "=" + returnobj.Language + "=" + returnobj.Resolution) == false) {
					ReturnHash.Add(returnobj.Title + "=" + returnobj.GUID + "=" + returnobj.Version + "=" + returnobj.Language + "=" + returnobj.Resolution, returnobj);
				} else {
					//MsgBox("Hit a Duplicate entry! Impossible")
				}
			}
			return ReturnHash;
		}

		/// <summary>
		/// Finds all objects that are referred to by the specified ISHObject as children.
		/// </summary>
		/// <param name="GUID"></param>
		/// <param name="ReferencedModules">Returns a hashtable of all referenced modules.</param>
		/// <param name="Version"></param>
		/// <param name="Language"></param>
		/// <param name="RequestedMetadata">You can limit the metadata to retrieve from the CMS using the common templating structure for CMS metadata requests.  Expected format is string of XML content.</param>        
		public bool GetReferencedModules(string GUID, ref Hashtable ReferencedModules, string Version = "1", string Language = "en", string RequestedMetadata = "")
		{
			string requestedobjects = "";
			StringBuilder RealRequestedMeta = new StringBuilder();
			if (string.IsNullOrEmpty(RequestedMetadata)) {
				RealRequestedMeta = oCommonFuncs.BuildRequestedMetadata();
			} else {
				RealRequestedMeta.Append(RequestedMetadata);
			}


			//Returns a list of referencing objects to "requestedobjects" string as xml:
			oISHAPIObjs.ISHReportsObj.GetReferencedDocObj(GUID, Version, Language, false, RealRequestedMeta.ToString(), ref requestedobjects);
			//Load the string as an xmldoc
			XmlDocument doc = new XmlDocument();
			doc.LoadXml(requestedobjects);
			//get our returned objects into a hashtable (key=GUID, Value=Information + XMLNode of returned data)
			ReferencedModules = GetReportedObjects(doc);
			Hashtable TrimmedReferencedModules = new Hashtable(ReferencedModules);
			//drop any objects that match our current object's GUID (could be more than one entry, depending on whether or not it's an image)
			foreach (DictionaryEntry obtainedmodule in ReferencedModules) {
				//if we find an entry that has the same GUID as our currently examined module's GUID, let's drop it.
				if (obtainedmodule.Value.guid == GUID) {
					TrimmedReferencedModules.Remove(obtainedmodule.Key);
				}
			}
			ReferencedModules = TrimmedReferencedModules;
			//we've removed all the reflective references.  Should be greater than 0 if we have any matches.

			if (ReferencedModules.Count > 0) {
				return true;
			} else {
				//only the requested GUID was found so there aren't any referenced children...
				return false;
			}
		}

		/// <summary>
		/// Finds all objects that refer to the specified ISHObject as parents.
		/// </summary>
		/// <param name="GUID"></param>
		/// <param name="ReferencingModules">Returns a hashtable of all referencing objects.</param>
		/// <param name="Version"></param>
		/// <param name="Language"></param>
		/// <param name="RequestedMetadata">You can limit the metadata to retrieve from the CMS using the common templating structure for CMS metadata requests.  Expected format is string of XML content.</param>        
		public bool GetReferencingModules(string GUID, ref Hashtable ReferencingModules, string Version = "1", string Language = "en", string RequestedMetadata = "")
		{
			string requestedobjects = "";
			StringBuilder RealRequestedMeta = new StringBuilder();
			if (string.IsNullOrEmpty(RequestedMetadata)) {
				RealRequestedMeta = oCommonFuncs.BuildRequestedMetadata();
			} else {
				RealRequestedMeta.Append(RequestedMetadata);
			}

			//Returns a list of referencing objects to "requestedobjects" string as xml:
			oISHAPIObjs.ISHReportsObj.GetReferencedByDocObj(GUID, Version, Language, false, RealRequestedMeta.ToString(), ref requestedobjects);
			//Load the string as an xmldoc
			XmlDocument doc = new XmlDocument();
			doc.LoadXml(requestedobjects);
			//get our returned objects into a hashtable (key=GUID, Value=Information + XMLNode of returned data)
			ReferencingModules = GetReportedObjects(doc);
			Hashtable TrimmedReferencingModules = new Hashtable(ReferencingModules);
			//drop any objects that match our current object's GUID (could be more than one entry, depending on whether or not it's an image)
			foreach (DictionaryEntry obtainedmodule in ReferencingModules) {
				//if we find an entry that has the same GUID as our currently examined module's GUID, let's drop it.
				if (obtainedmodule.Value.guid == GUID) {
					TrimmedReferencingModules.Remove(obtainedmodule.Key);
				}
			}
			ReferencingModules = TrimmedReferencingModules;
			//We've removed all reflective references from the list.
			if (ReferencingModules.Count > 0) {
				return true;
			} else {
				return false;
			}
		}
		#endregion

		#region "Search Functions"

		#endregion

		#region "Workflow Functions"

		#endregion

		#endregion




	}
}
