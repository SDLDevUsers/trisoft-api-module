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

	public class IshListOfValues : mainAPIclass
	{
		#region "Private Members"
			#endregion
		private readonly string strModuleName = "IshListOfValues";
		#region "Constructors"
		public IshListOfValues(string Username, string Password, string ServerURL)
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
		/// Returns a list of values given a field. Optionally filter on active or inactive.
		/// </summary>
		/// <param name="strFieldName">Ish Fieldname (FRESOLUTIONS)</param>
		/// <param name="sShowState">Valid values are: All, Active, or Inactive </param>
		/// <returns></returns>
		/// <remarks></remarks>
		public ArrayList GetMetaValues(string strFieldName, string sShowState = "All")
		{
			string fieldID = GetMetaFieldID(strFieldName);

			//pull the list of the values for the specified FieldName as XML from the server:
			string[] slovIds = { fieldID };
			ArrayList ishlovvalueStrings = new ArrayList();
			string ValuesInfo = new string("");
			try {
				//Get the Values list with corresponding container fields.
				ListOfValues25ServiceReference.ActivityFilter myActivityFilter = new ListOfValues25ServiceReference.ActivityFilter();
				switch (sShowState) {
					case "Active":
						myActivityFilter = ISHModulesNS.ListOfValues25ServiceReference.ActivityFilter.Active;
						break;
					case "Inactive":
						myActivityFilter = ISHModulesNS.ListOfValues25ServiceReference.ActivityFilter.Inactive;
						break;
					case "All":
						myActivityFilter = ISHModulesNS.ListOfValues25ServiceReference.ActivityFilter.None;
						break;
					default:
						return null;
				}

				ValuesInfo = oISHAPIObjs.ISHListOfValuesObj.RetrieveValues(slovIds, myActivityFilter);
				XmlDocument doc = new XmlDocument();
				doc.LoadXml(ValuesInfo);

				//Pull the metadata values out of the selected node.
				int p = 0;
				string valueString = new string("");
				foreach (XmlNode myxmlnode in doc.SelectNodes("//ishlovvalue")) {
					valueString = myxmlnode.SelectSingleNode("label").InnerText.ToString();
					ishlovvalueStrings.Add(valueString);
					p = p + 1;
				}
			} catch (Exception ex) {
				return null;
			}

			//Return the string array
			return ishlovvalueStrings;
		}
		/// <summary>
		/// Returns the field ID used by LCA for a given meta field name
		/// </summary>
		/// <param name="strFieldName"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		public string GetMetaFieldID(string strFieldName)
		{
			try {
				//Convert the string to an array of 1
				string[] sarrayFieldName = { strFieldName };
				string FieldInfo = new string("");
				try {
					FieldInfo = oISHAPIObjs.ISHListOfValuesObj.RetrieveLists({
						
					});

				} catch (Exception ex) {
				}

				//pull the ishref from the resulting XML
				XmlDocument doc = new XmlDocument();
				doc.LoadXml(FieldInfo);
				string strFieldID = doc.SelectSingleNode("//ishlov/label[text()=\"" + strFieldName + "\"]/parent::ishlov").Attributes.GetNamedItem("ishref").InnerText;

				return strFieldID;
			} catch (Exception ex) {
				return "";
			}

		}
		/// <summary>
		/// Gets the ID for a given metadata value in a given metadata field set.
		/// </summary>
		/// <param name="sMetaFieldID"></param>
		/// <param name="sMetaValue"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		public string GetMetaValueID(string sMetaFieldID, string sMetaValue)
		{
			try {
				//pull the list of the values for the specified FieldName as XML from the server:
				string[] slovIds = { sMetaFieldID };
				string response = "";
				response = oISHAPIObjs.ISHListOfValuesObj.RetrieveValues(slovIds, ISHModulesNS.ListOfValues25ServiceReference.ActivityFilter.None);
				XmlDocument doc = new XmlDocument();
				doc.LoadXml(response);
				string valuestring = new string("");
				valuestring = doc.SelectSingleNode("/ishlovs/ishlov/ishlovvalues/ishlovvalue/label[text()=\"" + sMetaValue + "\"]/../@ishref").InnerText.ToString();
				return valuestring;

			} catch (Exception ex) {
			}
			return "";
		}





		public string CreateMetaValue(string sISHValueId, string sMetaLabel, string sMetaDescription = "")
		{
			string newLovValueID = "";
			try {
				if (!string.IsNullOrEmpty(sISHValueId) & !string.IsNullOrEmpty(sMetaLabel)) {
					newLovValueID = oISHAPIObjs.ISHListOfValuesObj.CreateValue(sISHValueId, sMetaLabel, sMetaDescription);
				}


			} catch (Exception ex) {
				return newLovValueID;
			}


			return newLovValueID;
		}

		/// <summary>
		/// Modifies a Metadata Value in a given field. May or may not change the state or name. If no name change is desired, leave old and new the same.
		/// Note that leaving the state blank will always activate the metadata.
		/// TODO: Add function for getting current state and leaving as is if no state is specified.
		/// </summary>
		/// <param name="sIshFieldId"></param>
		/// <param name="sMetaLabelOld"></param>
		/// <param name="sMetaLabelNew"></param>
		/// <param name="sMetaDescription"></param>
		/// <param name="sMetaState"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		public bool ModifyMetaValue(string sIshFieldId, string sMetaLabelOld, string sMetaLabelNew, string sMetaDescription = "", string sMetaState = "Active")
		{
			try {
				bool bState = true;
				if (sMetaState == "Inactive") {
					bState = false;
				}

				if (!string.IsNullOrEmpty(sIshFieldId) & !string.IsNullOrEmpty(sMetaLabelOld) & !string.IsNullOrEmpty(sMetaLabelNew)) {
					string sMetaValueID = GetMetaValueID(sIshFieldId, sMetaLabelOld);
					oISHAPIObjs.ISHListOfValuesObj.UpdateValue(sIshFieldId, sMetaValueID, sMetaLabelNew, sMetaDescription, bState);
					return true;

				}
			} catch (Exception ex) {
				return false;
			}
			return false;
		}

		#endregion
	}
}
