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

	public class IshBaseline : mainAPIclass
	{
		#region "Private Members"
			#endregion
		private readonly string strModuleName = "ISHBaseline";
		#region "Constructors"
		public IshBaseline(string Username, string Password, string ServerURL)
		{
			//Make sure to use the FQDN up to the "WS" portion of your URL: "https://yourserver/InfoShareWS"
			oISHAPIObjs = new ISHObjs(Username, Password, ServerURL);
			//oISHAPIObjs.ISHAppObj.Login("InfoShareAuthor", Username, Password, Context)
		}
		#endregion
		#region "Properties"

		#endregion
		#region "Methods"


		#endregion
	}
}
