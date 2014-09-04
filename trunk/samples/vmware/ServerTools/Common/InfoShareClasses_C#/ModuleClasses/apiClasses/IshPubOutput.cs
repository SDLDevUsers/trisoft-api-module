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
namespace ISHModulesNS
{

	public class IshPubOutput : mainAPIclass
	{
		#region "Private Members"
			#endregion
		private readonly string strModuleName = "IshPubOutput";
		#region "Constructors"
		public IshPubOutput(string Username, string Password, string ServerURL)
		{
			//Make sure to use the FQDN up to the "WS" portion of your URL: "https://yourserver/InfoShareWS"
			oISHAPIObjs = new ISHObjs(Username, Password, ServerURL);
			//oISHAPIObjs.ISHAppObj.Login("InfoShareAuthor", Username, Password, Context)
		}
		#endregion
		#region "Properties"

		#endregion
		#region "Methods"
		public string DownloadOutput(string PubGUID, string PubVer, string OutLang, string OutType, string myFolder)
		{
			string strEdtID = "";
			string strEdtType = "";
			string strFileSize = "";
			string strMimeType = "";
			string strFileExt = "";
			long lngFileSize = 0;
			long lngIshLngRef = 0;
			string strPubTitle = "";
			string strBuildUser = "";
			string strBuildDate = "";
			string strPubServ = "";
			try {
				GetOutputInfo(PubGUID, PubVer, OutLang, OutType, ref strEdtID, ref strEdtType, ref strFileSize, ref strMimeType, ref strFileExt, ref lngIshLngRef,
				ref strPubTitle, ref strBuildDate, ref strBuildUser, ref strPubServ);
				lngFileSize = Convert.ToInt64(strFileSize);
				long remaining = lngFileSize;
				long plOff = 0;
				byte[] chunks = {
					
				};

				long chunk_size = 256000;
				//gather chunks until file is complete

				while ((remaining > 0)) {

					//This is to make sure we don't ask for a bigger chunk than there is left
					if ((chunk_size > remaining)) {
						chunk_size = remaining;
					}
					byte[] pboutbytes = {
						
					};


					oISHAPIObjs.ISHPubOutObj25.GetNextDataObjectChunkByIshLngRef(lngIshLngRef, strEdtID, ref plOff, ref chunk_size, ref pboutbytes);


					remaining = lngFileSize - plOff;
					//No need to update the offset, it appears GetNextDataObjectChunk... does it automatically.
					//plOff = plOff + chunk_size
					//append new chunk to current chunks
					List<byte> byteList = new List<byte>(chunks);
					byteList.AddRange(pboutbytes);
					byte[] byteArrayAll = byteList.ToArray();
					chunks = byteArrayAll;

				}
				//Create the storage folder:
				if (!Directory.Exists(myFolder)) {
					Directory.CreateDirectory(myFolder);
				}
				//Default CMS naming convention:
				//Lists and Steps=1=PDF - Press size with registration marks=en.pdf
				strPubTitle = oCommonFuncs.RemoveWindowsIllegalChars(strPubTitle);
				string filename = strPubTitle + "=" + PubVer + "=" + OutType + "=" + OutLang + "." + strFileExt;
				string fullfilepath = "";
				try {
					fullfilepath = myFolder + filename;
					ISHModulesNS.My.MyProject.Computer.FileSystem.WriteAllBytes(fullfilepath, chunks, false);
				} catch (Exception ex) {
					fullfilepath = myFolder + PubGUID + "=" + PubVer + "=" + OutType + "=" + OutLang + "." + strFileExt;
					ISHModulesNS.My.MyProject.Computer.FileSystem.WriteAllBytes(fullfilepath, chunks, false);
				}
				return fullfilepath;
			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(3, "Failure downloading content for specified output: " + PubGUID + "-v" + PubVer + "-" + OutLang + "-" + OutType + ". Reason: " + ex.Message, strModuleName + "-DownloadOutput");
				return false;
			}

			//// these are php header declarations for page load over http
			//header('Content-Type: ' . $mime);
			//header('Content-Disposition: attachment; filename="FILENAME' . '.'.$fileextension.'"');

			//// printing total $chunks to web page
			//echo $chunks;

		}
		/// <summary>
		/// ''' Gets an output's metadata along with a lot of file-specific info for downloading the file.
		/// </summary>
		/// <param name="PubGUID">Publication GUID</param>
		/// <param name="PubVer">Publication Version</param>
		/// <param name="OutLang">Language of the output</param>
		/// <param name="OutputType">Output type (PDF - Online, WebWorks, etc.)</param>
		/// <param name="outEdGUID"></param>
		/// <param name="outEDTType"></param>
		/// <param name="outFileSize">Returns the exact size of the file in bytes.</param>
		/// <param name="outMimeType"></param>
		/// <param name="outFileExt"></param>
		/// <param name="outIshLngRef"></param>
		/// <param name="outPubTitle"></param>
		/// <param name="outBuildDate"></param>
		/// <param name="outBuildUser"></param>
		/// <param name="OutPubServ"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		public bool GetOutputInfo(string PubGUID, string PubVer, string OutLang, string OutputType, ref string outEdGUID, ref string outEDTType, ref string outFileSize, ref string outMimeType, ref string outFileExt, ref long outIshLngRef,
		ref string outPubTitle, ref string outBuildDate, ref string outBuildUser, ref string OutPubServ)
		{
			//Get publication title:
			string myfilter = "<ishfields><ishfield name=\"FMAPID\" level=\"lng\">" + PubGUID + "</ishfield><ishfield name=\"DOC-LANGUAGE\" level=\"lng\">" + OutLang + "</ishfield></ishfields>";
			string myrequest = "<ishfields><ishfield name=\"VERSION\" level=\"version\"/><ishfield name=\"FISHPUBSTATUS\" level=\"lng\"/><ishfield name=\"FISHISRELEASED\" level=\"version\"/><ishfield name=\"DOC-LANGUAGE\" level=\"lng\"/><ishfield name=\"FISHOUTPUTFORMATREF\" level=\"lng\"/><ishfield name=\"FTITLE\" level=\"logical\"/></ishfields>";
			string responsexml = "";
			try {
				//[UPGRADE to 2013SP1] Changed the response to return the XML as a variable rather than an out param.
				responsexml = oISHAPIObjs.ISHPubOutObj25.Find(ISHModulesNS.PublicationOutput25ServiceReference.StatusFilter.ISHNoStatusFilter, myfilter, myrequest);
				XmlDocument pubinfo = new XmlDocument();
				pubinfo.LoadXml(responsexml);
				outPubTitle = pubinfo.SelectSingleNode("//ishfield[@name='FTITLE']").InnerText;
			} catch (Exception ex) {
				outPubTitle = PubGUID;
			}



			//Get output's plLngRef num:
			XmlDocument alloutputs = null;

			alloutputs = GetPubOutputsByISHRef(PubGUID, PubVer, OutLang, OutputType);
			if ((alloutputs == null)) {
				return null;
			}
			string strMyISHLangRef = null;
			long lngMyISHLangRef = 0;
			XmlNodeList outputs = alloutputs.SelectNodes("//ishobject");
			//For any returned outputs, get the ishlangref and convert it to a Long integer.
			foreach (XmlNode myoutput in outputs) {
				lngMyISHLangRef = 0;
				strMyISHLangRef = "";
				strMyISHLangRef = myoutput.Attributes.GetNamedItem("ishlngref").InnerText;
				lngMyISHLangRef = Convert.ToInt64(strMyISHLangRef);
			}

			//Get who and when output info:
			string strWhoWhenRequest = "<ishfields><ishfield name=\"VERSION\" level=\"version\"/><ishfield name=\"DOC-LANGUAGE\" level=\"lng\"/><ishfield name=\"FISHPUBLNGCOMBINATION\" level=\"lng\"/><ishfield name=\"FISHOUTPUTFORMATREF\" level=\"lng\"/><ishfield name=\"FISHEVENTID\" level=\"lng\"/><ishfield name=\"FISHPUBLISHER\" level=\"lng\"/></ishfields>";
			string strWhoWhenResult = "";
			//[Upgrade to 2013SP1] Return the result to a variable rather than an out param.
			strWhoWhenResult = oISHAPIObjs.ISHPubOutObj25.GetMetadataByIshLngRef(lngMyISHLangRef, strWhoWhenRequest);
			XmlDocument whowhen = new XmlDocument();
			whowhen.LoadXml(strWhoWhenResult);


			//Get output's metadata
			string requesteddata = "";
			//[Upgrade to 2013SP1] Return the result to a variable rather than an out param.
			requesteddata = oISHAPIObjs.ISHPubOutObj25.GetDataObjectInfoByIshLngRef(lngMyISHLangRef);
			XmlDocument metadata = new XmlDocument();
			metadata.LoadXml(requesteddata);




			try {
				outBuildUser = whowhen.SelectSingleNode("//ishfield[@name='FISHPUBLISHER']").InnerText;
				string EventID = whowhen.SelectSingleNode("//ishfield[@name='FISHEVENTID']").InnerText;
				//Split the event ID and return the various parts: e.g.: "593 cms-dev-app 20110518 11:00:46" where "<eventid> <pubServ> <Date> <Time>
				//TODO: Note that we're dropping the event ID here. Might be useful for sys admin to know that but most users don't need it.
				string[] eventids = null;
				eventids = EventID.Split(" ");
				OutPubServ = eventids[1];
				outBuildDate = eventids[2] + " " + eventids[3];
				outBuildDate = outBuildDate.Insert(6, "-");
				outBuildDate = outBuildDate.Insert(4, "-");
				outEdGUID = metadata.SelectSingleNode("/ishdataobjects/ishdataobject").Attributes.GetNamedItem("ed").InnerText.ToString();
				outEDTType = metadata.SelectSingleNode("/ishdataobjects/ishdataobject").Attributes.GetNamedItem("edt").InnerText.ToString();
				outFileExt = metadata.SelectSingleNode("/ishdataobjects/ishdataobject").Attributes.GetNamedItem("fileextension").InnerText.ToString();
				outFileSize = metadata.SelectSingleNode("/ishdataobjects/ishdataobject").Attributes.GetNamedItem("size").InnerText.ToString();
				outMimeType = metadata.SelectSingleNode("/ishdataobjects/ishdataobject").Attributes.GetNamedItem("mimetype").InnerText.ToString();
				outIshLngRef = lngMyISHLangRef;
				return (true);
			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(3, "Unable to find output for specified GUID: " + PubGUID + "-v" + PubVer + "-" + OutLang + "-" + OutputType + ". Reason: " + ex.Message, strModuleName + "-GetOutputInfo");
				return false;
			}
		}



		public string GetOutputState(string PubGUID, string Version, string Language, string OutputType)
		{
			//Get the output:
			XmlDocument alloutputs = null;

			alloutputs = GetPubOutputsByISHRef(PubGUID, Version, Language, OutputType);
			if ((alloutputs == null)) {
				return null;
			}
			string strMyISHLangRef = null;
			long lngMyISHLangRef = 0;
			XmlNodeList outputs = alloutputs.SelectNodes("//ishobject");
			//Should only return one output, but just in case, let's log an error if there are more than one found.
			if (outputs.Count > 1) {
				modErrorHandler.Errors.PrintMessage(3, "Multiple outputs found. Expected one unique output. PubGUID: " + PubGUID, strModuleName + "-GetOutputState");
				return null;
			}

			string Status = "UNKNOWN";
			foreach (XmlNode myoutput in outputs) {
				lngMyISHLangRef = 0;
				strMyISHLangRef = "";
				strMyISHLangRef = myoutput.Attributes.GetNamedItem("ishlngref").InnerText;
				lngMyISHLangRef = Convert.ToInt64(strMyISHLangRef);
				//Build a requesting filter.
				string requestedMeta = "<ishfields><ishfield name=\"FISHPUBSTATUS\" level=\"lng\"/></ishfields>";
				//Place to store resulting requested info.
				string strRequestedObjects = "";
				//Get the status of the requested output.

				//[UPGRADE to 2013SP1] Changed the result to return the result instead of just some arbitrary response.
				strRequestedObjects = oISHAPIObjs.ISHPubOutObj25.GetMetadataByIshLngRef(lngMyISHLangRef, requestedMeta);
				//.ISHPubOutObj25.GetMetaDataByIshLngRef(lngMyISHLangRef, requestedMeta, strRequestedObjects)
				XmlDocument mydoc = new XmlDocument();
				mydoc.LoadXml(strRequestedObjects);
				//Record the state:
				try {
					Status = mydoc.SelectSingleNode("//ishfield[@name=\"FISHPUBSTATUS\"]").InnerText;
				} catch (Exception ex) {
					Status = "NOTFOUND";
				}

			}
			return Status;



		}
		public bool CanBePublished(long IshLangRef)
		{
			//Build a requesting filter.
			string requestedMeta = "<ishfields><ishfield name=\"FISHPUBSTATUS\" level=\"lng\"/></ishfields>";
			//Place to store resulting requested info.
			string strRequestedObjects = "";
			//Build a list of forbidden states that don't allow publishing.
			List<string> mylist = new List<string>();
			mylist.Add("Publish Pending");
			mylist.Add("Publishing");
			mylist.Add("Released");
			//Get the status of the requested output.
			//[UPGRADE to 2013SP1] Changed the result to return the result instead of just some arbitrary response.
			strRequestedObjects = oISHAPIObjs.ISHPubOutObj25.GetMetadataByIshLngRef(IshLangRef, requestedMeta);
			XmlDocument mydoc = new XmlDocument();
			mydoc.LoadXml(strRequestedObjects);
			string Status = null;

			Status = mydoc.SelectSingleNode("//ishfield[@name=\"FISHPUBSTATUS\"]").InnerText;

			//Check to see if our status won't allow publishing.
			if (mylist.Contains(Status)) {
				return false;
			} else {
				return true;
			}
		}

		/// <summary>
		/// Get all the outputs according to the specified metadata being used as a filter.  
		/// If a filter is left out, it will grab all outputs regardless of what that meta is set to, except for Version.
		/// If Version is not specified, the latest will be used.
		/// Start any returned outputs if they can, report any that can't be started.
		/// </summary>
		/// <param name="PubGUID">Publication GUID</param>
		/// <param name="Version">Version</param>
		/// <param name="Language">Output Language</param>
		/// <param name="OutputType">Output Type</param>
		/// <returns>True or False</returns>
		/// <remarks></remarks>
		public bool StartPubOutput(string PubGUID, string Version = "latest", string Language = "", string OutputType = "all")
		{
			XmlDocument alloutputs = null;

			alloutputs = GetPubOutputsByISHRef(PubGUID, Version, Language, OutputType);
			if ((alloutputs == null)) {
				return null;
			}
			string strMyISHLangRef = null;
			long lngMyISHLangRef = 0;
			XmlNodeList outputs = alloutputs.SelectNodes("//ishobject");
			//For any returned outputs, start the publishing process or report any that can't be started for whatever reason.
			foreach (XmlNode myoutput in outputs) {
				lngMyISHLangRef = 0;
				strMyISHLangRef = "";
				strMyISHLangRef = myoutput.Attributes.GetNamedItem("ishlngref").InnerText;
				lngMyISHLangRef = Convert.ToInt64(strMyISHLangRef);
				string strOutEventID = "";
				try {
					if (CanBePublished(lngMyISHLangRef)) {
						oISHAPIObjs.ISHPubOutObj20.PublishByIshLngRef(lngMyISHLangRef, ref strOutEventID);
					} else {
						modErrorHandler.Errors.PrintMessage(2, "Output is already printing or is released. PubGUID: " + PubGUID + " ISHLangRef of output: " + lngMyISHLangRef.ToString(), strModuleName + "-StartAllPubOutputs");
					}
				} catch (Exception ex) {
					modErrorHandler.Errors.PrintMessage(3, "Failure while attempting to start an output. PubGUID: " + PubGUID + " ISHLangRef of output: " + lngMyISHLangRef.ToString(), strModuleName + "-StartAllPubOutputs");
				}

			}
			return true;
		}

		/// <summary>
		/// Gets an XML list of all publishing outputs of a given publication GUID. If no version is specified, the latest version is used.  Likewise, if no language is specified, all languages are retrieved.
		/// </summary>
		/// <param name="GUID">Publication GUID</param>
		/// <param name="Version">Publication Version</param>
		/// <param name="Language">Publication Language</param>
		/// <returns>XmlDocument</returns>
		/// <remarks></remarks>
		public XmlDocument GetPubOutputsByISHRef(string GUID, string Version = "latest", string Language = "", string OutputType = "all")
		{
			StringBuilder strXMLMetaDataFilter = oCommonFuncs.BuildPubMetaDataFilter(GUID);
			StringBuilder strXMLRequestedMetadata = oCommonFuncs.BuildFullPubMetadata();
			string strOutXMLObjList = null;
			string[] GUIDs = new string[1];
			GUIDs[0] = GUID;
			if (Version == "latest") {
				Version = GetLatestPubVersionNumber(GUID);
			}
			string requestedmeta = oCommonFuncs.BuildFullPubMetadata().ToString();
			string metafilter = oCommonFuncs.BuildPubMetaDataFilter(GUID, Version, Language, OutputType).ToString();
			try {
				strOutXMLObjList = oISHAPIObjs.ISHPubOutObj25.RetrieveMetadata(GUIDs, ISHModulesNS.PublicationOutput25ServiceReference.StatusFilter.ISHNoStatusFilter, metafilter, requestedmeta);
				//ISHPubOutObj25.Find(PublicationOutput25.eISHStatusgroup.ISHNoStatusFilter, strXMLMetaDataFilter.ToString, strXMLRequestedMetadata.ToString, strOutXMLObjList)
				XmlDocument ListofObjects = new XmlDocument();
				ListofObjects.LoadXml(strOutXMLObjList);
				return ListofObjects;
			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(2, "Error retrieving metadata while getting pub outputs. Message: " + ex.Message, strModuleName + "-GetPubOutputsByISHRef");
				return null;
			}

		}
		public bool CancelPublishOperation(string strPubGUID, string strPubVer, string strOutType, string strLanguage)
		{

			string EdGuid = "";
			string EDTType = "";
			long FileSize = 0;
			string mimetype = "";
			string fileext = "";
			long ishlngref = 0;
			string PubTitle = "";
			string state = "";
			string strBuildUser = "";
			string strBuildDate = "";
			string strPubServ = "";
			try {
				state = GetOutputState(strPubGUID, strPubVer, strLanguage, strOutType);
				if (state == "Pending" | state == "Publishing") {
					GetOutputInfo(strPubGUID, strPubVer, strLanguage, strOutType, ref EdGuid, ref EDTType, ref FileSize, ref mimetype, ref fileext, ref ishlngref,
					ref PubTitle, ref strBuildDate, ref strBuildUser, ref strPubServ);
					oISHAPIObjs.ISHPubOutObj20.CancelPublishByIshLngRef(ishlngref);
				}
			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(2, "Unable to cancel publishing on specified output. Skipping. Message: " + ex.Message, strModuleName + "-CancelPublishOperation");
			}
		}
		#endregion
	}
}
