Public Class Rooming_List
    Inherits System.Web.UI.Page
    Protected DepartureId As Integer
    Public func As New bizpower.cfunc
    Protected WithEvents dtlSupplier As DataList
    Protected conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Protected conB As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim cmdSelect As New SqlClient.SqlCommand
    Protected conBIns As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Public SiteRootPath As String = ConfigurationSettings.AppSettings("PegasusUrl")
    Public PegasusSiteDBName As String = ConfigurationSettings.AppSettings("PegasusSiteDBName")
    Protected Userid as String

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
        Userid = Trim(Request.Cookies("bizpegasus")("UserId"))

        If IsNumeric(Request.QueryString("dID")) Then
            DepartureId = CInt(Request.QueryString("dID"))

        End If

        If Not Page.IsPostBack Then
            getSuppliers()
        End If
    End Sub
    Sub getSuppliers()
        ''         Dim cmdSelect As New SqlClient.SqlCommand("select *,Country_Name from bizpower_pegasus_test.dbo.Suppliers S left join pegasus_test.dbo.Countries on pegasus_test.dbo.Countries.Country_Id=S.Country_Id " & _
        ''  " where S.Country_Id in (select Country_Id from pegasus_test.dbo.Tours_Departures left join pegasus_test.dbo.Tours_Countries on pegasus_test.dbo.Tours_Departures.Tour_id=pegasus_test.dbo.Tours_Countries.Tour_id where Departure_Id=" & CInt(DepartureId) & ")" & _
        '' " order by supplier_Name", conB)
        '' Dim cmdSelect As New SqlClient.SqlCommand("select *,Country_Name from bizpower_pegasus.dbo.Suppliers S left join pegasus.dbo.Countries on pegasus.dbo.Countries.Country_Id=S.Country_Id " & _
        ''    " where S.Country_Id in (select Country_Id from pegasus_test.dbo.Tours_Departures left join pegasus_test.dbo.Tours_Countries on pegasus_test.dbo.Tours_Departures.Tour_id=pegasus_test.dbo.Tours_Countries.Tour_id where Departure_Id=" & CInt(DepartureId) & ")" & _
        ''     " order by supplier_Name", conB)
        'Dim cmdSelect As New SqlClient.SqlCommand("select VS.supplier_Id,S.GUID,S.supplier_Name,Departure_Id,Vouchers_Status ,Country_Name " & _
        '          " from VouchersToSuppliers VS   left join Suppliers S on VS.supplier_Id=S.supplier_Id " & _
        '          " left join pegasus_test.dbo.Countries on pegasus_test.dbo.Countries.Country_Id=S.Country_Id  where Departure_Id = " & CInt(DepartureId), conB)

        Dim cmdSelect As New SqlClient.SqlCommand("select distinct VS.supplier_Id,S.GUID,S.supplier_Name,Departure_Id,Country_Name " & _
                     " from VouchersToSuppliers VS   left join Suppliers S on VS.supplier_Id=S.supplier_Id " & _
                     " left join " & PegasusSiteDBName & ".dbo.Countries on " & PegasusSiteDBName & ".dbo.Countries.Country_Id=S.Country_Id  where Departure_Id = " & CInt(DepartureId), conB)



        cmdSelect.CommandType = CommandType.Text
        conB.Open()


        dtlSupplier.DataSource = cmdSelect.ExecuteReader()
        dtlSupplier.DataBind()
        conB.Close()

    End Sub

End Class
