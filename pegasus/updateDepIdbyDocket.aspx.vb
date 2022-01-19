Public Class updateDepIdbyDocket
    Inherits System.Web.UI.Page
    Public dr_m As SqlClient.SqlDataReader
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conU As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Public BizpowerDBName As String = ConfigurationSettings.AppSettings("BizpowerDBName")
    Public PegasusSiteDBName As String = ConfigurationSettings.AppSettings("PegasusSiteDBName")
    Protected App_Dep, DocketpfileNum As Integer


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

        Dim cmdSelect As New SqlClient.SqlCommand("SELECT Tour_Travelers.Docket_pfileNum,  Departure_Date,Departure_Date_End," & _
                      " APPEALS.APPEAL_ID, APPEALS.Departure_Id, FORM_VALUE.FIELD_ID, FORM_VALUE.FIELD_VALUE FROM APPEALS INNER JOIN" & _
                      " FORM_VALUE ON APPEALS.APPEAL_ID = FORM_VALUE.APPEAL_ID  INNER JOIN Tour_Travelers ON FORM_VALUE.Field_Value = cast(Tour_Travelers.Docket_pfileNum as varchar) " & _
                      " inner join " & PegasusSiteDBName & ".dbo.Tours_Departures on " & PegasusSiteDBName & ".dbo.Tours_Departures.Departure_Id=APPEALS.Departure_Id  where FORM_VALUE.Field_id=40622 " & _
                      " and IsNumeric(FORM_VALUE.Field_Value)=1 and Tour_Travelers.Departure_Id<>APPEALS.Departure_Id", con)
              cmdSelect.CommandType = CommandType.Text

        con.Open()
        dr_m = cmdSelect.ExecuteReader()
        While dr_m.Read
            If Not dr_m("Departure_Id") Is DBNull.Value Then
                App_Dep = dr_m("Departure_Id")
            End If
            If Not dr_m("Docket_pfileNum") Is DBNull.Value Then
                DocketpfileNum = dr_m("Docket_pfileNum")
            End If
            conU.Open()
            Dim cmdUpd As New SqlClient.SqlCommand("update Tour_Travelers set Departure_Id=@App_Dep where Docket_pfileNum=@DocketpfileNum", conU)
            cmdUpd.Parameters.Add("@App_Dep", SqlDbType.Int).Value = CInt(App_Dep)
            cmdUpd.Parameters.Add("@DocketpfileNum", SqlDbType.BigInt).Value = CInt(DocketpfileNum)
            cmdUpd.ExecuteNonQuery()
            conU.Close()
            Response.Write("update Tour_Travelers set Departure_Id=" & App_Dep & " where Docket_pfileNum=" & DocketpfileNum)

        End While

        dr_m.Close()
        cmdSelect.Dispose()
        con.Close()

    End Sub

End Class
