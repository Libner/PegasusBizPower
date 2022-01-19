Imports bizpower_pegasus2018.passportcard
Imports System.Net
Imports System.Xml
Imports System.IO
Imports System.Web.Services.Protocols


Public Class Insurance_Send
    Inherits System.Web.UI.Page



#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object


    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()



    End Sub

#End Region


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
        'Dim LeadID As String
        ' Dim passportcard_ws As bizpower_pegasus.passportcard.PC_Mobile_WebService = New bizpower_pegasus.passportcard.PC_Mobile_WebService


        ' Response.Write(passportcard_ws.GetType())

        Response.End()

      

        '  passportcard_ws.Mobile_LeadoMat_App_Handler_Vendor()
         '  If Not Page.IsPostBack Then
        ' Dim passportcard_ws As New passportcard.PC_Mobile_WebService
        '    Dim passportcard_ws As New passportcard.PC_Mobile_WebService

        '  passportcard_ws = New passportcard.PC_Mobile_WebService
        ' Response.Write(passportcard_ws.GetType())

        '  Dim LeadID As bizpower_pegasus.passportcard.LeadomatMobile
        'Dim LeadID As String
        ' passportcard_ws.Mobile_LeadoMat_App_Handler_Vendor(4, "46402", "0", "0", "054 0987654;", "TEST@TEST.COM", "test7", "test77", "01/01/2015", "הערות", 243, "4", "נציג", "UT/VRCK5SuOb6bEb10F6WQ==", "", "").LeadNumber.ToString()




        ' Response.Write("LeadID: " & LeadID)
        Response.End()
        '  End If

    End Sub

End Class
