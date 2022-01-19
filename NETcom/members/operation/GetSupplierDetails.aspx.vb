Public Class GetSupplierDetails
    Inherits System.Web.UI.Page
    Public SId As Integer
    Protected conB As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Public func As New bizpower.cfunc

    Dim rssub As SqlClient.SqlDataReader
    Protected Country_Name, supplier_Name, Country_Id, supplier_Name1, supplier_Name2, supplier_Name3, supplier_Name4, supplier_Job1, supplier_Job2, supplier_Job3, supplier_Job4 As String
    Protected supplier_Tel1, supplier_Tel2, supplier_Tel3, supplier_Tel4, supplier_Ext1, supplier_Ext2, supplier_Ext3, supplier_Ext4, supplier_Phone1, supplier_Phone2, supplier_Phone3, supplier_Phone4 As String
    Protected supplier_Email1, supplier_Email2, supplier_Email3, supplier_Email4, supplier_Descr1, supplier_Descr2, supplier_Descr3, supplier_Descr4, TermsPayment, supplier_Descr As String


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
        If IsNumeric(Request.QueryString("Id")) Then
            SId = Request.QueryString("Id")
        Else
            SId = 0
        End If
        If SId > 0 Then
            Dim cmdTmp = New SqlClient.SqlCommand("Select Country_Name,supplier_Name,S.Country_Id,supplier_Name1,supplier_Name2,supplier_Name3,supplier_Name4,supplier_Job1," & _
   " supplier_Job2,supplier_Job3,supplier_Job4,supplier_Tel1,supplier_Tel2,supplier_Tel3,supplier_Tel4," & _
   " supplier_Ext1,supplier_Ext2,supplier_Ext3,supplier_Ext4," & _
   " supplier_Phone1,supplier_Phone2,supplier_Phone3,supplier_Phone4," & _
   " supplier_Email1,supplier_Email2,supplier_Email3,supplier_Email4," & _
   " supplier_Descr1,supplier_Descr2,supplier_Descr3,supplier_Descr4," & _
   " TermsPayment,supplier_Descr From Suppliers S left join pegasus.dbo.Countries C on S.Country_id=C.Country_Id Where supplier_Id = @SId", conB)
            cmdTmp.CommandType = CommandType.Text
            cmdTmp.Parameters.Add("@SId", SqlDbType.Int).Value = CInt(SId)
            conB.Open()
            rssub = cmdTmp.ExecuteReader(CommandBehavior.SingleRow)
            If rssub.Read() Then
                If Not rssub("supplier_Name") Is DBNull.Value Then
                    supplier_Name = Trim(rssub("supplier_Name"))
                Else
                    supplier_Name = ""
                End If
                If Not rssub("Country_Name") Is DBNull.Value Then
                    Country_Name = Trim(rssub("Country_Name"))
                Else
                    Country_Name = ""
                End If


                Country_Id = Trim(rssub("Country_Id"))
                supplier_Name1 = Trim(rssub("supplier_Name1"))
                supplier_Name2 = Trim(rssub("supplier_Name2"))
                supplier_Name3 = Trim(rssub("supplier_Name3"))
                supplier_Name4 = Trim(rssub("supplier_Name4"))
                supplier_Job1 = Trim(rssub("supplier_Job1"))
                supplier_Job2 = Trim(rssub("supplier_Job2"))
                supplier_Job3 = Trim(rssub("supplier_Job3"))
                supplier_Job4 = Trim(rssub("supplier_Job4"))

                supplier_Tel1 = Trim(rssub("supplier_Tel1"))
                supplier_Tel2 = Trim(rssub("supplier_Tel2"))
                supplier_Tel3 = Trim(rssub("supplier_Tel3"))
                supplier_Tel4 = Trim(rssub("supplier_Tel4"))


                supplier_Ext1 = Trim(rssub("supplier_Ext1"))
                supplier_Ext2 = Trim(rssub("supplier_Ext2"))
                supplier_Ext3 = Trim(rssub("supplier_Ext3"))
                supplier_Ext4 = Trim(rssub("supplier_Ext4"))

                supplier_Phone1 = Trim(rssub("supplier_Phone1"))
                supplier_Phone2 = Trim(rssub("supplier_Phone2"))
                supplier_Phone3 = Trim(rssub("supplier_Phone3"))
                supplier_Phone4 = Trim(rssub("supplier_Phone4"))

                supplier_Email1 = Trim(rssub("supplier_Email1"))
                supplier_Email2 = Trim(rssub("supplier_Email2"))
                supplier_Email3 = Trim(rssub("supplier_Email3"))
                supplier_Email4 = Trim(rssub("supplier_Email4"))



                supplier_Descr1 = Trim(rssub("supplier_Descr1"))
                supplier_Descr2 = Trim(rssub("supplier_Descr2"))
                supplier_Descr3 = Trim(rssub("supplier_Descr3"))
                supplier_Descr4 = Trim(rssub("supplier_Descr4"))

                TermsPayment = Trim(rssub("TermsPayment"))
                supplier_Descr = Trim(rssub("supplier_Descr"))
            End If
            rssub = Nothing
        End If
    End Sub

End Class
