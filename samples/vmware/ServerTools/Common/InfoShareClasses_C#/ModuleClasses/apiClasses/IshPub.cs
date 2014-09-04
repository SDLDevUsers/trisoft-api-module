using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Xml;
using ErrorHandlerNS;
namespace ISHModulesNS
{
	public class IshPub : mainAPIclass
	{
		#region "Private Members"
			#endregion
		private readonly string strModuleName = "ISHPub";
		#region "Constructors"
		public IshPub(string Username, string Password, string ServerURL)
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
		/// Saves all objects from a specified publication, version, and language
		/// </summary>
		/// <param name="PubGUID"></param>
		/// <param name="PubVer"></param>
		/// <param name="Language"></param>
		/// <param name="SavePath"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		public bool ExportPublicationbyBaseline(string PubGUID, string PubVer, string Language, string SavePath)
		{
			//Get the baseline objects
			Dictionary<string, CMSObject> myBaseline = null;
			myBaseline = GetBaselineObjects(PubGUID, PubVer, Language);
			ArrayList CurRes = new ArrayList();
			CurRes.Add("High");
			CurRes.Add("Low");
			//for each baseline object, save the files to the specified path (getobjbyid with path)
			foreach (KeyValuePair<string, CMSObject> myObject in myBaseline) {
				if (myObject.Value.IshType == "ISHIllustration") {
					foreach (string resolution in CurRes) {
						if (ObjectExists(myObject.Value.GUID, myObject.Value.Version, Language, resolution)) {
							GetObjByID(myObject.Value.GUID, myObject.Value.Version, Language, resolution, SavePath);
						}
					}
				} else {
					GetObjByID(myObject.Value.GUID, myObject.Value.Version, Language, "", SavePath);
				}

			}
		}

		/// <summary>
		/// Resets all map and topic title properties to match their title elements. 
		/// </summary>
		/// <param name="PubGUID"></param>
		/// <param name="PubVer"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		public bool UpdateAllTitleProperties(string PubGUID, string PubVer)
		{
			Dictionary<string, CMSObject> dictBaseLine = GetBaselineObjects(PubGUID, PubVer);
			foreach (KeyValuePair<string, CMSObject> entry in dictBaseLine) {
				if (entry.Value.IshType == "ISHModule" | entry.Value.IshType == "ISHMasterDoc" | entry.Value.IshType == "ISHLibrary") {
					UpdateTitleProperty(entry.Key, entry.Value.Version);
				}
			}
			return true;
		}



		private XmlDocument GetPubObjByID(string GUID, string Version)
		{
			XmlNode MyNode = null;
			XmlDocument MyDoc = new XmlDocument();
			XmlDocument MyMeta = new XmlDocument();
			string XMLString = "";
			string ISHMeta = "";
			string ISHResult = "";

			//Call the CMS to get our content!
			try {
				//[UPGRADE to 2013SP1] Changed the result to return the "XMLString" instead of just some arbitrary response.
				XMLString = oISHAPIObjs.ISHPubOutObj25.Find(ISHModulesNS.PublicationOutput25ServiceReference.StatusFilter.ISHNoStatusFilter, oCommonFuncs.BuildPubMetaDataFilter(GUID, Version).ToString(), oCommonFuncs.BuildFullPubMetadata().ToString());
			} catch (Exception ex) {
				//modErrorHandler.Errors.PrintMessage(3, "Failed to retrieve XML from CMS server: " + ex.Message, strModuleName)
				return null;
			}

			//Load the XML and get the metadata:
			try {
				MyDoc.LoadXml(XMLString);
				return MyDoc;
			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(3, "Failed to retrieve object metadata from returned XML: " + ex.Message, strModuleName);
				return null;
			}
		}


		#endregion
	}
}
