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


	[ComClass(clsISHObj.ClassId, clsISHObj.InterfaceId, clsISHObj.EventsId)]
	public class clsISHObj
	{

		#region "COM GUIDs"
		// These  GUIDs provide the COM identity for this class 
		// and its COM interfaces. If you change them, existing 
		// clients will no longer be able to access the class.
		public const string ClassId = "990ade3a-03fb-49d3-86ee-7b7c21f958d0";
		public const string InterfaceId = "759d75cb-111a-4f9f-9508-887bdc9dafc0";
			#endregion
		public const string EventsId = "36982ff3-d385-4e70-81a3-b7dbeab2acf5";

		// A creatable COM class must have a Public Sub New() 
		// with no parameters, otherwise, the class will not be 
		// registered in the COM registry and cannot be created 
		// via CreateObject.
		//Public Sub New()
		//    CopyDTDFile()
		//    'Load up our XML Templates for each type:
		//    oDocument.LoadXMLTemplates()

		//End Sub
		/// <summary>
		/// Creates the main ISH object to be used while interfacing with the LiveContent Architect system. 
		/// TODO: Need to enable dynamic application switching between servers...
		/// </summary>
		/// <param name="Username"></param>
		/// <param name="Password"></param>
		/// <param name="ServerURL">Server URL.</param>
		/// <remarks>Server URL should include FQDN including "WS". For instance: https://trisoft03.sdlproducts.com/InfoShareWS2 </remarks>

		public clsISHObj(string Username, string Password, string ServerURL)
		{
			SetContext(Username, Password, ServerURL);
			oDocument.oCommonFuncs.LoadXMLTemplates();
			oDocument.oCommonFuncs.CopyDTDFile();
		}






		public IshApplication oApplication;
		public IshDocument oDocument;
		public IshBaseline oBaseline;
		public IshConditions oConditions;
		public IshListOfValues oListOfValues;
		public IshMeta oMeta;
		public IshOutput oOutput;
		public IshPub oPub;
		public IshPubOutput oPubOutput;
		public IshReports oReports;
		public IshSearch oSearch;
		public IshWorkflow oWorkflow;

		public IshFolder oFolder;

		public ISHObjs oIshObjs;
		public string CMSServerURL = new string("");
		//Private Shared m_Context As New String("")


		public ArrayList DeletedGUIDs {
			get { return oDocument.DeletedGUIDs; }
		}
		public ArrayList DeleteFailedGUIDs {
			get { return oDocument.DeleteFailedGUIDs; }
		}


		private string ISHApp {
			get { return "InfoShareAuthor"; }
		}
		private string strModuleName {
			get { return "clsISHObj"; }
		}
		/// <summary>
		/// Context of the current user's credentials on the specified CMS URL.
		/// The easiest way of setting this value is to run SetContext after instantiating the object.
		/// However, it can also be set directly.
		/// </summary>
		public string Context {
			get { return oApplication.Context; }
			set {
				oApplication.Context = value;
				oDocument.Context = value;
				oBaseline.Context = value;
				oConditions.Context = value;
				oMeta.Context = value;
				oOutput.Context = value;
				oPub.Context = value;
				oPubOutput.Context = value;
				oReports.Context = value;
				oSearch.Context = value;
				oWorkflow.Context = value;
				oFolder.Context = value;
			}
		}

		/// <summary>
		/// Used to set the context of the current user's credentials on the specified CMS URL.
		/// </summary>
		public bool SetContext(string uname, string passwd, string RepositoryURL)
		{
			try {
				//First, test the App web reference with the passed URL.
				oApplication = new IshApplication(uname, passwd, RepositoryURL);
				CMSServerURL = RepositoryURL;
			} catch (Exception ex) {
				modErrorHandler.Errors.PrintMessage(3, "Login failed: " + ex.Message.ToString(), strModuleName);
				return false;
			}


			//Create all the new objects as previously defined but with the proper credentials:
			oApplication = new IshApplication(uname, passwd, CMSServerURL);
			oDocument = new IshDocument(uname, passwd, CMSServerURL);
			oBaseline = new IshBaseline(uname, passwd, CMSServerURL);
			oConditions = new IshConditions(uname, passwd, CMSServerURL);
			oListOfValues = new IshListOfValues(uname, passwd, CMSServerURL);
			oMeta = new IshMeta(uname, passwd, CMSServerURL);
			oOutput = new IshOutput(uname, passwd, CMSServerURL);
			oPub = new IshPub(uname, passwd, CMSServerURL);
			oPubOutput = new IshPubOutput(uname, passwd, CMSServerURL);
			oReports = new IshReports(uname, passwd, CMSServerURL);
			oSearch = new IshSearch(uname, passwd, CMSServerURL);
			oWorkflow = new IshWorkflow(uname, passwd, CMSServerURL);
			oFolder = new IshFolder(uname, passwd, CMSServerURL);


			//'Set all the ISHObj web refs to use this passed URL...
			//Try
			//    'Sets the context property which results in all the oISH objects' contexts getting set to the same context.
			//    Context = m_context
			//Catch ex As Exception
			//    modErrorHandler.Errors.PrintMessage(1, "Unable to set Web Reference URL.  Message: " + ex.Message, strModuleName)
			//    Return False
			//End Try




			return true;
		}
	}
}



