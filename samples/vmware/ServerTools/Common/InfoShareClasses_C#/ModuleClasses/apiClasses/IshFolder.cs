using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Xml;
using ErrorHandlerNS;
using System.Text;
using System.IO;
namespace ISHModulesNS
{

	public class IshFolder : mainAPIclass
	{
		#region "Private Members"
			#endregion
		private readonly string strModuleName = "ISHDocument";
		#region "Constructors"
		public IshFolder(string Username, string Password, string ServerURL)
		{
			//Make sure to use the FQDN up to the "WS" portion of your URL: "https://yourserver/InfoShareWS"
			oISHAPIObjs = new ISHObjs(Username, Password, ServerURL);
			//oISHAPIObjs.ISHAppObj.Login("InfoShareAuthor", Username, Password, Context)
		}
		#endregion
		#region "Properties"

		#endregion
		#region "Methods"
		/// <summary>
		/// Creates a new folder as a sub-folder to the given ParentFolderID and returns the newly created ID.  If a folder of the same name already exists, the function returns the ID of that folder instead.
		/// </summary>
		/// <param name="ParentFolderID">Containing folder's ID.</param>
		/// <param name="ISHType">Folder type: ISHNone, ISHMasterDoc, ISHModule, ISHIllustration, ISHReusedObj, ISHTemplate, and ISHLibrary</param>
		/// <param name="FolderName">Name of folder to create/reuse.</param>
		/// <param name="NewFolderID">New/reused folder's ID.</param>
		/// <param name="OwnershipGroup">Ownership group based on groups defined in the CMS.</param>
		public string CreateOrUseFolder(long ParentFolderID, string ISHType, string FolderName, ref long NewFolderID, string OwnershipGroup = "", string ReadAccess = "")
		{
			string CreateFolderResult = null;
			Folder20ServiceReference.ISHFolderType RealIshType = default(Folder20ServiceReference.ISHFolderType);
			switch (ISHType) {
				case "ISHNone":
					RealIshType = ISHModulesNS.Folder20ServiceReference.ISHFolderType.ISHNone;
					break;
				case "ISHMasterDoc":
					RealIshType = ISHModulesNS.Folder20ServiceReference.ISHFolderType.ISHMasterDoc;
					break;
				case "ISHModule":
					RealIshType = ISHModulesNS.Folder20ServiceReference.ISHFolderType.ISHModule;
					break;
				case "ISHIllustration":
					RealIshType = ISHModulesNS.Folder20ServiceReference.ISHFolderType.ISHIllustration;
					break;
				//Case "ISHReusedObj"
				//RealIshType = Folder20ServiceReference.ISHFolderType.ISHReusedObj
				case "ISHTemplate":
					RealIshType = ISHModulesNS.Folder20ServiceReference.ISHFolderType.ISHTemplate;
					break;
				case "ISHLibrary":
					RealIshType = ISHModulesNS.Folder20ServiceReference.ISHFolderType.ISHLibrary;
					break;
				case "ISHPublication":
					RealIshType = ISHModulesNS.Folder20ServiceReference.ISHFolderType.ISHPublication;
					break;
			}
			if (FolderExists(ParentFolderID, FolderName)) {
				NewFolderID = GetFolderIDByName(ParentFolderID, FolderName, ref OwnershipGroup, ref ReadAccess);
				CreateFolderResult = "Reused existing ID.";
			} else {
				try {
					CreateFolderResult = oISHAPIObjs.ISHFolderObj.Create(ParentFolderID, RealIshType, FolderName, OwnershipGroup, ref NewFolderID, ReadAccess);
				} catch (Exception ex) {
					return ex.Message;
				}
			}

			return CreateFolderResult;
		}

		/// <summary>
		/// Function that returns the ID of a folder specified by a full CMS Path
		/// Remember to include //Doc at the beginning.
		/// </summary>
		/// <param name="FullCMSPath"></param>
		public long GetFolderIDByPath(string FullCMSPath, ref string Result = "", ref string OwnershipGroup = "", ref string ReadAccess = "")
		{
			FullCMSPath = FullCMSPath.Trim();
			//First, make sure it's a properly formated CMS Path, starting with "//"
			if (FullCMSPath.StartsWith("//") == false) {
				//invalid path given - must start with //
				return -1;
			} else {
				//trim "//" off the front and any / at the end (if there is one)
				FullCMSPath = FullCMSPath.Replace("//", "");
				FullCMSPath = FullCMSPath.TrimEnd("/");
			}
			// Assume the root ID is 0 and drop the first folder that corresponds to it.
			long currentID = 0;
			//if we have more than just the root, we need to trim the root folder name off.
			if (FullCMSPath.Contains("/")) {
				FullCMSPath = FullCMSPath.Remove(0, FullCMSPath.IndexOf("/") + 1);
			} else {
				//if we JUST have the root folder, make sure the user called it the right name.
				string RealRootName = "";
				ISHModulesNS.Folder20ServiceReference.ISHFolderType rootishtype = default(ISHModulesNS.Folder20ServiceReference.ISHFolderType);
				string OutQuery = "";
				string OwnedBy = "";
				long rootfolderid = GetFolderIDByName(0, FullCMSPath, ref OwnershipGroup, ref ReadAccess);
				if (rootfolderid > -1) {
					oISHAPIObjs.ISHFolderObj.GetProperties(rootfolderid, ref RealRootName, ref rootishtype, ref OutQuery, ref OwnershipGroup, ref ReadAccess);
					//if it is the right name, return 0
					if (RealRootName == FullCMSPath) {
						return 0;
					} else {
						//otherwise, return -1
						Result = "Invalid name for root folder specified.";
						return -1;
					}
				} else {
					Result = "Invalid name for root folder specified.";
					return -1;
				}


			}

			long subID = 0;
			//Check to see if we have sub-folders...
			if (FullCMSPath.Length > 0) {
				//Loop through each foldername
				foreach (string foldername in FullCMSPath.Split("/")) {
					subID = GetFolderIDByName(currentID, foldername, ref OwnershipGroup, ref ReadAccess);
					if (!(subID == -1)) {
						currentID = subID;
					} else {
						// We hit a directory in our path that doesn't exist in the CMS.  This will return a value of "0" to the calling function.  Root is invalid for a sub.
						Result = "Path doesn't exist in the currently specified CMS.";
						return -1;
					}
				}
			} else {
				//no subfolders so just do the import at the root.
				currentID = 0;
			}

			return currentID;
		}




		/// <summary>
		/// Function that returns the ID of a subfolder specified by its name
		/// </summary>
		/// <param name="ParentFOlderID">ID of the Parent Folder</param>
		/// <param name="SubFolderName">Name of the subfolder to find</param>
		public long GetFolderIDByName(long ParentFolderID, string SubFolderName, ref string OwnershipGroup, ref string ReadAccess)
		{
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

			//find the id of the one matching the name we're looking for 
			string subid = "";
			string foldername = "";
			ISHModulesNS.Folder20ServiceReference.ISHFolderType curishtype = ISHModulesNS.Folder20ServiceReference.ISHFolderType.ISHNone;
			string curoutquery = "";
			try {
				subid = doc.SelectSingleNode("//ishfolders/ishfolder[@name='" + SubFolderName + "']/@ishfolderref").Value.ToString();

				oISHAPIObjs.ISHFolderObj.GetProperties(Convert.ToInt64(subid), ref foldername, ref curishtype, ref curoutquery, ref OwnershipGroup, ref ReadAccess);
				return Convert.ToInt64(subid);
			} catch (Exception ex) {
				// Unable to find the node we were looking for in the returned XML.  
				return -1;
			}
		}
		/// <summary>
		/// Returns an array of long integers containing all the ids of child sub-folders of the given folder ID.
		/// </summary>
		/// <param name="FolderID">ID of folder to return sub-ids from.</param>
		/// <returns>Array of Long Integers indicating children of specified folder ID.</returns>
		/// <remarks>If the specified FolderID is invalid, the arraylist will be empty.  It is recommended to check to see if the folder ID exists in the CMS before running this method.</remarks>
		public ArrayList GetSubFolderIDs(long FolderID)
		{
			ArrayList subfolderIDs = new ArrayList();
			//get a folderlist of all children of currentfolderid
			string subfolderlistXML = "";
			string CMSReply = null;
			try {
				CMSReply = oISHAPIObjs.ISHFolderObj.GetSubFolders(FolderID.ToString(), 1, ref subfolderlistXML);

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

			if (doc.HasChildNodes) {
				foreach (XmlAttribute subid in doc.SelectNodes("//ishfolders/ishfolder/@ishfolderref")) {
					subfolderIDs.Add(Convert.ToInt64(subid.Value.ToString()));
				}
			}
			return subfolderIDs;
		}
		/// <summary>
		/// Checks to see if a subfolder of a given ID exists within a parentfolder specified by a separate ID.
		/// </summary>
		public bool FolderExists(long ParentFolderID, string SubFolderName)
		{
			string OwnershipGroup = "";
			string ReadAccess = "";
			long subID = GetFolderIDByName(ParentFolderID, SubFolderName, ref OwnershipGroup, ref ReadAccess);

			// Subfolder not found
			if (subID == -1) {
				return false;
			//Subfolder found
			} else {
				return true;
			}
		}
		/// <summary>
		/// Checks to see if a given folder has sub-content (objects).
		/// </summary>
		public bool FolderHasContents(long FolderID)
		{
			string foldercontentXML = "";
			string CMSReply = null;
			try {
				CMSReply = oISHAPIObjs.ISHFolderObj.GetContents(FolderID.ToString(), "", "", ref foldercontentXML);

			} catch (Exception ex) {
			}
			//Load the subfolderlistXML into an xml document
			// Create the reader from the string.
			StringReader strReader = new StringReader(foldercontentXML);

			XmlReader reader = XmlReader.Create(strReader);
			// Create the new XMLdoc and load the content into it.
			XmlDocument doc = new XmlDocument();
			doc.PreserveWhitespace = true;
			doc.Load(reader);

			XmlNodeList significantchildren = doc.SelectNodes("//ishobject");
			//children found
			if (significantchildren.Count > 0) {
				return true;
			// no children found
			} else {
				return false;
			}
		}
		/// <summary>
		/// Checks to see if a given folder has sub-folders.
		/// </summary>
		public object FolderHasSubFolders(long FolderID)
		{
			//get a folderlist of all children of currentfolderid
			string subfolderlistXML = "";
			string CMSReply = null;
			try {
				CMSReply = oISHAPIObjs.ISHFolderObj.GetSubFolders(FolderID.ToString(), 1, ref subfolderlistXML);

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
			XmlNodeList subfolders = doc.SelectNodes("//ishfolder");
			if (subfolders.Count == 0) {
				return false;
			} else {
				return true;
			}
		}

		/// <summary>
		/// Attempts to delete all content (not folders) of a specified folder. Uses IshDocument.DeleteObjectRecursivelyByGUID with "DeleteSubs" set to False for each sub-object.
		/// </summary>
		/// <returns>Returns true if successful.</returns>
		public bool DeleteSubContent(long FolderID)
		{
			StringBuilder RealRequestedMeta = new StringBuilder();
			RealRequestedMeta = oCommonFuncs.BuildRequestedMetadata();
			string requestedmeta = RealRequestedMeta.ToString();
			string foldercontentXML = "";
			string CMSReply = null;
			try {
				CMSReply = oISHAPIObjs.ISHFolderObj.GetContents(FolderID.ToString(), "", requestedmeta, ref foldercontentXML);

			} catch (Exception ex) {
			}
			//Load the subcontentlistXML into an xml document
			// Create the reader from the string.
			StringReader strReader = new StringReader(foldercontentXML);

			XmlReader reader = XmlReader.Create(strReader);
			// Create the new XMLdoc and load the content into it.
			XmlDocument doc = new XmlDocument();
			doc.PreserveWhitespace = true;
			doc.Load(reader);

			XmlNodeList significantchildren = doc.SelectNodes("//ishobject");
			//First, collect the info from each xml node.  Use the structure found in IshDocument
			if (significantchildren.Count > 0) {
				Hashtable ContentHash = new Hashtable();
				ContentHash = GetReportedObjects(doc);
				foreach (DictionaryEntry child in ContentHash) {
					if (DeleteObjectRecursivelyByGUID(child.Value.GUID, child.Value.Version, child.Value.Language, child.Value.Resolution, false)) {
						modErrorHandler.Errors.PrintMessage(1, "Deleting object in targeted folder. GUID: " + child.Key.ToString(), strModuleName);
						//Return True
					} else {
						modErrorHandler.Errors.PrintMessage(2, "Unable to delete object in targeted folder. GUID: " + child.Key.ToString(), strModuleName);
						//Return False
					}
				}
			} else {
				// no significant modules to delete, carry on.
				return true;
			}
			return true;
		}
		/// <summary>
		/// Function that deletes specified folder and all empty sub-folders if they are empty. Returns true if successful, false if not. Optionally allows all sub-content to be deleted as well (using DeleteSubContent).
		/// Remember to include // at the beginning.
		/// </summary>
		/// <param name="delresult">Stores first error message if anything goes wrong during delete.</param>
		public bool DeleteFolderRecursive(long FolderID, ref string DelResult, bool DeleteContentToo = true)
		{
			bool process_result = false;
			bool hassubs = FolderHasSubFolders(FolderID);
			if (hassubs & FolderID > 0) {
				//MsgBox("Has SubFolders")
				foreach (long subfolderID in GetSubFolderIDs(FolderID)) {
					DeleteFolderRecursive(subfolderID, ref DelResult, DeleteContentToo);
					if (DelResult.Length > 0) {
						//Return False
					}
				}
			}
			if (DeleteContentToo == true) {
				if (FolderHasContents(FolderID)) {
					//delete subcontent
					if (DeleteSubContent(FolderID)) {
						process_result = true;
					} else {
						process_result = false;
					}
					//delete folder
					//Return DeleteFolder(FolderID, DelResult)
				}
			}


			//all subfolders and sub-conent have been deleted.  Should be able to delete unless content was left in folders by user... if that's the case, we should return false.
			return DeleteFolder(FolderID, ref DelResult);



		}

		/// <summary>
		/// Deletes a specified empty folder.
		/// </summary>
		/// <param name="DelResult">Result of the attempt.</param>
		/// <returns>Returns True if able to delete the folder, False if not.</returns>
		public bool DeleteFolder(long folderID, ref string DelResult)
		{
			if (FolderHasContents(folderID) == false) {
				try {
					oISHAPIObjs.ISHFolderObj.Delete(folderID);
					DelResult = "";
					return true;
				} catch (Exception ex) {
					DelResult = "Error deleting FolderID: " + folderID.ToString() + " Message: " + ex.Message;
					modErrorHandler.Errors.PrintMessage(3, DelResult, strModuleName);
					return false;
				}


			} else {
				DelResult = "Folder identified by ID#" + folderID.ToString() + " has sub-content.  Unable to delete.";
				modErrorHandler.Errors.PrintMessage(2, DelResult, strModuleName);
				return false;
			}

		}

		#endregion



	}
}
