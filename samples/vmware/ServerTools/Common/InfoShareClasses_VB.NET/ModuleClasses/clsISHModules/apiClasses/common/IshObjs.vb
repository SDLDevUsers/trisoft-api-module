Public Class ISHObjs
    Public ISHAppObj As New Application20.Application20
    Public ISHDocObj As New DocumentObj20.DocumentObj20
    Public ISHDocObj25 As New DocumentObj25.DocumentObj25
    Public ISHBaselineObj As New Baseline25.BaseLine25
    Public ISHCondObj As New Condition20.Condition20
    Public ISHMetaObj As New MetaDataAssist20.MetaDataAssist20
    Public ISHOutputObj As New OutputFormat20.OutputFormat20
    Public ISHPubObj As New Publication20.Publication20
    Public ISHPubOutObj20 As New PublicationOutput20.PublicationOutput20
    Public ISHPubOutObj25 As New PublicationOutput25.PublicationOutput25
    Public ISHFolderObj As New Folder20.Folder20
    Public ISHReportsObj As New Reports20.Reports20
    Public ISHSearchObj As New Search20.Search20
    Public ISHWorkflowObj As New Workflow20.WorkFlow20
    Sub New(ByVal Username As String, ByVal Password As String, ByVal ServerURL As String)

        ISHPubOutObj20.Url = ServerURL + "/InfoShareWS/PublicationOutput20.asmx"
        ISHPubOutObj25.Url = ServerURL + "/InfoShareWS/PublicationOutput25.asmx"
        ISHAppObj.Url = ServerURL + "/InfoShareWS/Application20.asmx"
        ISHFolderObj.Url = ServerURL + "/InfoShareWS/Folder20.asmx"
        ISHWorkflowObj.Url = ServerURL + "/InfoShareWS/Workflow20.asmx"
        ISHBaselineObj.Url = ServerURL + "/InfoShareWS/Baseline25.asmx"
        ISHCondObj.Url = ServerURL + "/InfoShareWS/Condition20.asmx"
        ISHDocObj.Url = ServerURL + "/InfoShareWS/DocumentObj20.asmx"
        ISHDocObj25.Url = ServerURL + "/InfoShareWS/DocumentObj25.asmx"
        ISHMetaObj.Url = ServerURL + "/InfoShareWS/MetaDataAssist20.asmx"
        ISHOutputObj.Url = ServerURL + "/InfoShareWS/OutputFormat20.asmx"
        ISHPubObj.Url = ServerURL + "/InfoShareWS/Publication20.asmx"
        ISHReportsObj.Url = ServerURL + "/InfoShareWS/Reports20.asmx"
        ISHSearchObj.Url = ServerURL + "/InfoShareWS/Search20.asmx"
    End Sub
End Class
