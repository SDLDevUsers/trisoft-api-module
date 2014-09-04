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

	public class IshMeta : mainAPIclass
	{
		#region "Private Members"
			#endregion
		private readonly string strModuleName = "ISHMeta";
		#region "Constructors"
		public IshMeta(string Username, string Password, string ServerURL)
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
		/// Determines if a user has priviledges of a particular role and belongs to a specified group.
		/// </summary>
		/// <param name="Username">Username in the CMS.</param>
		/// <param name="Role">Role priviledge such as "Administrator", "Author", "Illustrator", etc.</param>
		/// <param name="UserGroup">(OPTIONAL) The group to search within ("Technical Publications", for instance).</param>
		/// <returns></returns>
		/// <remarks></remarks>
		public bool UserHasRole(string Username, string Role, string UserGroup = "Default Department")
		{
			string[] returneduserlist = null;
			ArrayList userlist = new ArrayList();
			//[UPGRADE] Changed the result to return the "retuneduserlist" instead of just true/false
			returneduserlist = oISHAPIObjs.ISHMetaObj.GetUsers(Role, UserGroup);

			foreach (string uname in returneduserlist) {
				userlist.Add(uname);
			}
			if (userlist.Contains(Username)) {
				return true;
			} else {
				return false;
			}


		}
		#endregion
	}
}
