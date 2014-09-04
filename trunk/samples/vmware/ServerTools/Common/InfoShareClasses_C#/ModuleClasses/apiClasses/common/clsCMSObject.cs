using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
namespace ISHModulesNS
{
	public class CMSObject
	{
		private string _GUID;
		private string _Version;
		private string _IshType;
		private string _ReportedItems;
		public string GUID {
			get { return _GUID; }
			set { _GUID = value; }
		}
		public string Version {
			get { return _Version; }
			set { _Version = value; }
		}
		public string IshType {
			get { return _IshType; }
			set { _IshType = value; }
		}
		public string ReportedItems {
			get { return _ReportedItems; }
			set { _ReportedItems = value; }
		}

		public CMSObject(string strGUID, string strVersion, string strIshType, string strReportedItems = "<reporteditems/>")
		{
			_GUID = strGUID;
			_Version = strVersion;
			_IshType = strIshType;
			_ReportedItems = strReportedItems;
		}
	}
}
