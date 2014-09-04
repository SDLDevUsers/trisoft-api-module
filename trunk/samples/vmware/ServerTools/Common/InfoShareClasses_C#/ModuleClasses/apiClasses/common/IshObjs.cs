using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.ServiceModel;
using ISHModulesNS;
namespace ISHModulesNS
{
	public class ISHObjs
	{
		//Can't initialize here with a variable.
		//Public ISHAppObj As New Application25ServiceReference.ApplicationClient(New Uri(ServerURL + "/InfoShareWS/Wcf/API25/Application.svc"))
		//Public ISHDocObj As New DocumentObj20ServiceReference.DocumentObjClient(New Uri(ServerURL + "/InfoShareWS/Wcf/API20/DocumentObj.svc"))
		//Public ISHDocObj25 As New DocumentObj25ServiceReference.DocumentObjClient(New Uri(ServerURL + "/InfoShareWS/Wcf/API25/DocumentObj.svc"))
		//Public ISHBaselineObj As New Baseline25ServiceReference.BaselineClient(New Uri(ServerURL + "/InfoShareWS/Wcf/API25/Baseline.svc"))
		//Public ISHCondObj As New ConditionManagement10ServiceReference.ConditionManagementClient(New Uri(ServerURL + "/InfoShareWS/Wcf/API20/ConditionManagement.svc"))
		//Public ISHMetaObj As New MetadataAssist20ServiceReference.MetaDataAssistClient(New Uri(ServerURL + "/InfoShareWS/Wcf/API20/MetaDataAssist.svc"))
		//Public ISHOutputObj As New OutputFormat20ServiceReference.OutputFormatClient(New Uri(ServerURL + "/InfoShareWS/Wcf/API20/OutputFormat.svc"))
		//Public ISHPubObj As New Publication20ServiceReference.PublicationClient(New Uri(ServerURL + "/InfoShareWS/Wcf/API20/Publication.svc"))
		//Public ISHPubOutObj20 As New PublicationOutput20ServiceReference.PublicationOutputClient(New Uri(ServerURL + "/InfoShareWS/Wcf/API20/PublicationOutput.svc"))
		//Public ISHPubOutObj25 As New PublicationOutput25ServiceReference.PublicationOutputClient(New Uri(ServerURL + "/InfoShareWS/Wcf/API25/PublicationOutput.svc"))
		//Public ISHFolderObj As New Folder20ServiceReference.FolderClient(New Uri(ServerURL + "/InfoShareWS/Wcf/API20/Folder.svc"))
		//Public ISHReportsObj As New Reports20ServiceReference.ReportsClient(New Uri(ServerURL + "/InfoShareWS/Wcf/API20/Reports.svc"))
		//Public ISHSearchObj As New Search20ServiceReference.SearchClient(New Uri(ServerURL + "/InfoShareWS/Wcf/API20/Search.svc"))
		//Public ISHWorkflowObj As New Workflow20ServiceReference.WorkflowClient(New Uri(ServerURL + "/InfoShareWS/Wcf/API20/Workflow.svc"))

		//Declare the public properties and then set them when instantiating the object.
		public Application25ServiceReference.ApplicationClient ISHAppObj;
		public DocumentObj20ServiceReference.DocumentObjClient ISHDocObj;
		public DocumentObj25ServiceReference.DocumentObjClient ISHDocObj25;
		public Baseline25ServiceReference.BaselineClient ISHBaselineObj;
		public ConditionManagement10ServiceReference.ConditionManagementClient ISHCondObj;
		public ListOfValues25ServiceReference.ListOfValuesClient ISHListOfValuesObj;
		public MetaDataAssist20ServiceReference.MetaDataAssistClient ISHMetaObj;
		public OutputFormat20ServiceReference.OutputFormatClient ISHOutputObj;
		public Publication20ServiceReference.PublicationClient ISHPubObj;
		public PublicationOutput20ServiceReference.PublicationOutputClient ISHPubOutObj20;
		public PublicationOutput25ServiceReference.PublicationOutputClient ISHPubOutObj25;
		public Folder20ServiceReference.FolderClient ISHFolderObj;
		public Reports20ServiceReference.ReportsClient ISHReportsObj;
		public Search20ServiceReference.SearchClient ISHSearchObj;
		public Workflow20ServiceReference.WorkflowClient ISHWorkflowObj;

		public User25ServiceReference.UserClient ISHUserObj;

		public ISHObjs()
		{

		}


		public ISHObjs(string Username, string Password, string ServerURL)
		{
			try {
				//Initialize all the service refs with the proper ServerURL (https://yourserverURL/InfoShareWS)
				//binding used by the service when it is added (see Reference.vb on the service)
				//Error on objects: "The provided URI scheme 'https' is invalid; expected 'http'."
				//I can't figure out how to set the security if I were to use this one...
				//Dim binding As System.ServiceModel.Channels.Binding = New WSFederationHttpBinding()


				//Recommended Binding on the web (unfortunately, only support SOAP 1.1; doesn't work):
				//Dim binding As new BasicHttpBinding()
				//binding.Security.Mode = BasicHttpSecurityMode.Transport

				///'NOTE: Need to try using 'BasicHttpsBinding()' - it might work.

				//Recommended by google (this one used SOAP 1.2, but didn't use the correct transport type):
				//Dim binding As New WSHttpBinding()

				//Possible binding?
				//Objects have error, "{"The provided URI scheme 'https' is invalid; expected 'http'.
				//Dim binding As New WSFederationHttpBinding()
				//binding.Security.Mode = WSFederationHttpSecurityMode.Message

				//Possible 'https' binding?
				//Fails on GetVersion(): "The incoming policy could not be validated. For more information, please see the event log."
				//Dim binding As New WSFederationHttpBinding()
				//binding.Security.Mode = WSFederationHttpSecurityMode.TransportWithMessageCredential




				// ''Dim appAddress As New EndpointAddress(ServerURL + "/Wcf/API25/Application.svc")
				// ''ISHAppObj = New Application25ServiceReference.ApplicationClient() '(binding, appAddress) '("CustomBinding_Application", ServerURL + "/Wcf/API25/Application.svc")
				//' ''Login because the others will use the credentials for this one when they attempt to bind.
				// ''ISHAppObj.ClientCredentials.UserName.UserName = Username
				// ''ISHAppObj.ClientCredentials.UserName.Password = Password
				// ''Dim version As String = ISHAppObj.GetVersion()

				// ''Dim docAddress As New EndpointAddress(ServerURL + "/Wcf/API20/DocumentObj.svc")
				// ''ISHDocObj = New DocumentObj20ServiceReference.DocumentObjClient() '(binding, docAddress) '(, ServerURL + "/Wcf/API20/DocumentObj.svc")

				// ''Dim doc25Address As New EndpointAddress(ServerURL + "/Wcf/API25/DocumentObj.svc")
				// ''ISHDocObj25 = New DocumentObj25ServiceReference.DocumentObjClient() '(binding, doc25Address) '(, ServerURL + "/Wcf/API25/DocumentObj.svc")

				// ''Dim baselineAddress As New EndpointAddress(ServerURL + "/Wcf/API25/Baseline.svc")
				// ''ISHBaselineObj = New Baseline25ServiceReference.BaselineClient() '(binding, baselineAddress) '(, ServerURL + "/Wcf/API25/Baseline.svc")

				// ''Dim condAddress As New EndpointAddress(ServerURL + "/Wcf/API20/ConditionManagement.svc")
				// ''ISHCondObj = New ConditionManagement10ServiceReference.ConditionManagementClient() '(binding, condAddress) '(, ServerURL + "/Wcf/API20/ConditionManagement.svc")

				// ''ISHListOfValuesObj = New ListOfValues25ServiceReference.ListOfValuesClient() '(binding, baselineAddress) '(, ServerURL + "/Wcf/API25/Baseline.svc")

				// ''Dim metaAddress As New EndpointAddress(ServerURL + "/Wcf/API20/MetaDataAssist.svc")
				// ''ISHMetaObj = New MetaDataAssist20ServiceReference.MetaDataAssistClient() '(binding, metaAddress) '(, ServerURL + "/Wcf/API20/MetaDataAssist.svc")

				// ''Dim outputAddress As New EndpointAddress(ServerURL + "/Wcf/API20/OutputFormat.svc")
				// ''ISHOutputObj = New OutputFormat20ServiceReference.OutputFormatClient() '(binding, outputAddress) '(, ServerURL + "/Wcf/API20/OutputFormat.svc")

				// ''Dim publicationAddress As New EndpointAddress(ServerURL + "/Wcf/API20/Publication.svc")
				// ''ISHPubObj = New Publication20ServiceReference.PublicationClient() '(binding, publicationAddress) '(, ServerURL + "/Wcf/API20/Publication.svc")

				// ''Dim pubOut20Address As New EndpointAddress(ServerURL + "/Wcf/API20/PublicationOutput.svc")
				// ''ISHPubOutObj20 = New PublicationOutput20ServiceReference.PublicationOutputClient() '(binding, pubOut20Address) '(, ServerURL + "/Wcf/API20/PublicationOutput.svc")

				// ''Dim pubOut25Address As New EndpointAddress(ServerURL + "/Wcf/API25/PublicationOutput.svc")
				// ''ISHPubOutObj25 = New PublicationOutput25ServiceReference.PublicationOutputClient() '(binding, pubOut25Address) '(, ServerURL + "/Wcf/API25/PublicationOutput.svc")

				// ''Dim folderAddress As New EndpointAddress(ServerURL + "/Wcf/API20/Folder.svc")
				// ''ISHFolderObj = New Folder20ServiceReference.FolderClient() '(binding, folderAddress) '(, ServerURL + "/Wcf/API20/Folder.svc")

				// ''Dim reportsAddress As New EndpointAddress(ServerURL + "/Wcf/API20/Reports.svc")
				// ''ISHReportsObj = New Reports20ServiceReference.ReportsClient() '(binding, reportsAddress) '(, ServerURL + "/Wcf/API20/Reports.svc")

				// ''Dim searchAddress As New EndpointAddress(ServerURL + "/Wcf/API20/Search.svc")
				// ''ISHSearchObj = New Search20ServiceReference.SearchClient() '(binding, searchAddress) '(, ServerURL + "/Wcf/API20/Search.svc")

				// ''Dim workflowAddress As New EndpointAddress(ServerURL + "/Wcf/API20/Workflow.svc")
				// ''ISHWorkflowObj = New Workflow20ServiceReference.WorkflowClient() '(binding, workflowAddress) '(, ServerURL + "/Wcf/API20/Workflow.svc")

				// ''Dim userAddress As New EndpointAddress(ServerURL + "/Wcf/API25/User.svc")
				// ''ISHUserObj = New User25ServiceReference.UserClient() '(binding, userAddress) '(,ServerURL + "/Wcf/API25/User.svc")




				ISHAppObj = new Application25ServiceReference.ApplicationClient();
				//(binding, appAddress) '("CustomBinding_Application", ServerURL + "/Wcf/API25/Application.svc")
				//Login because the others will use the credentials for this one when they attempt to bind.
				ISHAppObj.ClientCredentials.UserName.UserName = Username;
				ISHAppObj.ClientCredentials.UserName.Password = Password;
				string version = ISHAppObj.GetVersion();


				ISHDocObj = new DocumentObj20ServiceReference.DocumentObjClient();
				//(binding, docAddress) '(, ServerURL + "/Wcf/API20/DocumentObj.svc")
				ISHDocObj25 = new DocumentObj25ServiceReference.DocumentObjClient();
				//(binding, doc25Address) '(, ServerURL + "/Wcf/API25/DocumentObj.svc")
				ISHBaselineObj = new Baseline25ServiceReference.BaselineClient();
				//(binding, baselineAddress) '(, ServerURL + "/Wcf/API25/Baseline.svc")
				ISHCondObj = new ConditionManagement10ServiceReference.ConditionManagementClient();
				//(binding, condAddress) '(, ServerURL + "/Wcf/API20/ConditionManagement.svc")
				ISHListOfValuesObj = new ListOfValues25ServiceReference.ListOfValuesClient();
				//(binding, baselineAddress) '(, ServerURL + "/Wcf/API25/Baseline.svc")
				ISHMetaObj = new MetaDataAssist20ServiceReference.MetaDataAssistClient();
				//(binding, metaAddress) '(, ServerURL + "/Wcf/API20/MetaDataAssist.svc")
				ISHOutputObj = new OutputFormat20ServiceReference.OutputFormatClient();
				//(binding, outputAddress) '(, ServerURL + "/Wcf/API20/OutputFormat.svc")
				ISHPubObj = new Publication20ServiceReference.PublicationClient();
				//(binding, publicationAddress) '(, ServerURL + "/Wcf/API20/Publication.svc")
				ISHPubOutObj20 = new PublicationOutput20ServiceReference.PublicationOutputClient();
				//(binding, pubOut20Address) '(, ServerURL + "/Wcf/API20/PublicationOutput.svc")
				ISHPubOutObj25 = new PublicationOutput25ServiceReference.PublicationOutputClient();
				//(binding, pubOut25Address) '(, ServerURL + "/Wcf/API25/PublicationOutput.svc")
				ISHFolderObj = new Folder20ServiceReference.FolderClient();
				//(binding, folderAddress) '(, ServerURL + "/Wcf/API20/Folder.svc")
				ISHReportsObj = new Reports20ServiceReference.ReportsClient();
				//(binding, reportsAddress) '(, ServerURL + "/Wcf/API20/Reports.svc")
				ISHSearchObj = new Search20ServiceReference.SearchClient();
				//(binding, searchAddress) '(, ServerURL + "/Wcf/API20/Search.svc")
				ISHWorkflowObj = new Workflow20ServiceReference.WorkflowClient();
				//(binding, workflowAddress) '(, ServerURL + "/Wcf/API20/Workflow.svc")
				ISHUserObj = new User25ServiceReference.UserClient();
				//(binding, userAddress) '(,ServerURL + "/Wcf/API25/User.svc")


				ISHPubOutObj20.ClientCredentials.UserName.UserName = Username;
				ISHPubOutObj20.ClientCredentials.UserName.Password = Password;
				ISHPubOutObj25.ClientCredentials.UserName.UserName = Username;
				ISHPubOutObj25.ClientCredentials.UserName.Password = Password;



				ISHFolderObj.ClientCredentials.UserName.UserName = Username;
				ISHFolderObj.ClientCredentials.UserName.Password = Password;

				ISHWorkflowObj.ClientCredentials.UserName.UserName = Username;
				ISHWorkflowObj.ClientCredentials.UserName.Password = Password;

				ISHBaselineObj.ClientCredentials.UserName.UserName = Username;
				ISHBaselineObj.ClientCredentials.UserName.Password = Password;

				ISHCondObj.ClientCredentials.UserName.UserName = Username;
				ISHCondObj.ClientCredentials.UserName.Password = Password;

				ISHDocObj.ClientCredentials.UserName.UserName = Username;
				ISHDocObj.ClientCredentials.UserName.Password = Password;

				ISHDocObj25.ClientCredentials.UserName.UserName = Username;
				ISHDocObj25.ClientCredentials.UserName.Password = Password;

				ISHListOfValuesObj.ClientCredentials.UserName.UserName = Username;
				ISHListOfValuesObj.ClientCredentials.UserName.Password = Password;

				ISHMetaObj.ClientCredentials.UserName.UserName = Username;
				ISHMetaObj.ClientCredentials.UserName.Password = Password;

				ISHOutputObj.ClientCredentials.UserName.UserName = Username;
				ISHOutputObj.ClientCredentials.UserName.Password = Password;

				ISHPubObj.ClientCredentials.UserName.UserName = Username;
				ISHPubObj.ClientCredentials.UserName.Password = Password;

				ISHReportsObj.ClientCredentials.UserName.UserName = Username;
				ISHReportsObj.ClientCredentials.UserName.Password = Password;

				ISHSearchObj.ClientCredentials.UserName.UserName = Username;
				ISHSearchObj.ClientCredentials.UserName.Password = Password;

				ISHUserObj.ClientCredentials.UserName.UserName = Username;
				ISHUserObj.ClientCredentials.UserName.Password = Password;

			} catch (Exception ex) {
			}

		}

		//Public Function Login(ByVal Username As String, ByVal Password As String, ByVal ServerURL As String)

		//    Try
		//        'Initialize all the service refs with the proper ServerURL (https://yourserverURL/InfoShareWS)
		//        'Dim binding As New BasicHttpBinding()
		//        'Dim address as
		//        ISHAppObj = New Application25ServiceReference.ApplicationClient(ISHAppObj.Endpoint.Name, ServerURL + "/Wcf/API25/Application.svc")
		//        ISHDocObj = New DocumentObj20ServiceReference.DocumentObjClient(ISHDocObj.Endpoint.Name, ServerURL + "/Wcf/API20/DocumentObj.svc")
		//        ISHDocObj25 = New DocumentObj25ServiceReference.DocumentObjClient(ISHDocObj25.Endpoint.Name, ServerURL + "/Wcf/API25/DocumentObj.svc")
		//        ISHBaselineObj = New Baseline25ServiceReference.BaselineClient(ISHBaselineObj.Endpoint.Name, ServerURL + "/Wcf/API25/Baseline.svc")
		//        ISHCondObj = New ConditionManagement10ServiceReference.ConditionManagementClient(ISHCondObj.Endpoint.Name, ServerURL + "/Wcf/API20/ConditionManagement.svc")
		//        ISHMetaObj = New MetadataAssist20ServiceReference.MetaDataAssistClient(ISHMetaObj.Endpoint.Name, ServerURL + "/Wcf/API20/MetaDataAssist.svc")
		//        ISHOutputObj = New OutputFormat20ServiceReference.OutputFormatClient(ISHOutputObj.Endpoint.Name, ServerURL + "/Wcf/API20/OutputFormat.svc")
		//        ISHPubObj = New Publication20ServiceReference.PublicationClient(ISHPubObj.Endpoint.Name, ServerURL + "/Wcf/API20/Publication.svc")
		//        ISHPubOutObj20 = New PublicationOutput20ServiceReference.PublicationOutputClient(ISHPubOutObj20.Endpoint.Name, ServerURL + "/Wcf/API20/PublicationOutput.svc")
		//        ISHPubOutObj25 = New PublicationOutput25ServiceReference.PublicationOutputClient(ISHPubOutObj25.Endpoint.Name, ServerURL + "/Wcf/API25/PublicationOutput.svc")
		//        ISHFolderObj = New Folder20ServiceReference.FolderClient(ISHFolderObj.Endpoint.Name, ServerURL + "/Wcf/API20/Folder.svc")
		//        ISHReportsObj = New Reports20ServiceReference.ReportsClient(ISHReportsObj.Endpoint.Name, ServerURL + "/Wcf/API20/Reports.svc")
		//        ISHSearchObj = New Search20ServiceReference.SearchClient(ISHSearchObj.Endpoint.Name, ServerURL + "/Wcf/API20/Search.svc")
		//        ISHWorkflowObj = New Workflow20ServiceReference.WorkflowClient(ISHWorkflowObj.Endpoint.Name, ServerURL + "/Wcf/API20/Workflow.svc")


		//        'Set the user credentials for each service endpoint
		//        ISHPubOutObj20.ClientCredentials.UserName.UserName = Username
		//        ISHPubOutObj20.ClientCredentials.UserName.Password = Password
		//        ISHPubOutObj25.ClientCredentials.UserName.UserName = Username
		//        ISHPubOutObj25.ClientCredentials.UserName.Password = Password

		//        ISHAppObj.ClientCredentials.UserName.UserName = Username
		//        ISHAppObj.ClientCredentials.UserName.Password = Password

		//        ISHFolderObj.ClientCredentials.UserName.UserName = Username
		//        ISHFolderObj.ClientCredentials.UserName.Password = Password

		//        ISHWorkflowObj.ClientCredentials.UserName.UserName = Username
		//        ISHWorkflowObj.ClientCredentials.UserName.Password = Password

		//        ISHBaselineObj.ClientCredentials.UserName.UserName = Username
		//        ISHBaselineObj.ClientCredentials.UserName.Password = Password

		//        ISHCondObj.ClientCredentials.UserName.UserName = Username
		//        ISHCondObj.ClientCredentials.UserName.Password = Password

		//        ISHDocObj.ClientCredentials.UserName.UserName = Username
		//        ISHDocObj.ClientCredentials.UserName.Password = Password

		//        ISHDocObj25.ClientCredentials.UserName.UserName = Username
		//        ISHDocObj25.ClientCredentials.UserName.Password = Password

		//        ISHMetaObj.ClientCredentials.UserName.UserName = Username
		//        ISHMetaObj.ClientCredentials.UserName.Password = Password

		//        ISHOutputObj.ClientCredentials.UserName.UserName = Username
		//        ISHOutputObj.ClientCredentials.UserName.Password = Password

		//        ISHPubObj.ClientCredentials.UserName.UserName = Username
		//        ISHPubObj.ClientCredentials.UserName.Password = Password

		//        ISHReportsObj.ClientCredentials.UserName.UserName = Username
		//        ISHReportsObj.ClientCredentials.UserName.Password = Password

		//        ISHSearchObj.ClientCredentials.UserName.UserName = Username
		//        ISHSearchObj.ClientCredentials.UserName.Password = Password
		//        Return True
		//    Catch ex As Exception
		//        Return False
		//    End Try
		//End Function
	}
}
