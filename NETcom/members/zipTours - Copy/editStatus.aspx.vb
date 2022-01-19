Public Class editStatus
    Inherits System.Web.UI.Page
    Protected con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Dim cmdSelect As New SqlClient.SqlCommand
    Protected func As New include.funcs
    Protected WithEvents btnSubmit As UI.WebControls.Button
    Protected DepartureId As Integer
    Protected StatusId As String
    Protected sStatus As HtmlSelect
    Public dtStatus As New DataTable
    Dim primKeydtStatus(0) As Data.DataColumn





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
          If IsNumeric(Request.QueryString("DepartureId")) Then
            DepartureId = Request.QueryString("DepartureId")

        Else
            DepartureId = 0
        End If
        If Not Page.IsPostBack Then
            If DepartureId > 0 Then
                cmdSelect = New SqlClient.SqlCommand("SELECT Status_ForBizpower FROM Tours_Departures " & _
                " WHERE (Departure_Id = @Departure_Id)", con)
                cmdSelect.Parameters.Add("@Departure_Id", SqlDbType.Int).Value = CInt(DepartureId)
                con.Open()
                Dim dr As SqlClient.SqlDataReader
                dr = cmdSelect.ExecuteReader(CommandBehavior.SingleRow)
                If dr.Read Then
                    If Not IsDBNull(dr("Status_ForBizpower")) Then
                        StatusId = dr("Status_ForBizpower")
                    End If

                End If
                con.Close()
            End If
            GetStatus()
            If StatusId <> "" Then
                sStatus.Items.FindByValue(StatusId).Selected = True
            End If

        Else

            '  If Request.Form("sStatus") <> "" Then
          
            cmdSelect = New SqlClient.SqlCommand("update Tours_Departures set Status_ForBizpower=@StatusId WHERE (Departure_Id = @DepartureId)", con)
            cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
            cmdSelect.Parameters.Add("@StatusId", SqlDbType.Int).Value = sStatus.Value
            con.Open()
            cmdSelect.ExecuteNonQuery()
            con.Close()
            Dim cScript As String
            cScript = "<script language='javascript'>"
            cScript += "self.close();window.opener.location.href = window.opener.location </script>"
            RegisterStartupScript("ReloadScrpt", cScript)
        End If
        '   End If


    End Sub
    Sub GetStatus()
        Dim cmdSelect As New SqlClient.SqlCommand("select  status_id,status_Name,status_Color,status_FntColor from Status_ForBizpower ORDER BY status_id", con)
        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtStatus)
        con.Close()
        primKeydtStatus(0) = dtStatus.Columns("status_id")
        dtStatus.PrimaryKey = primKeydtStatus

        sStatus.Items.Clear()
        For i As Integer = 0 To dtStatus.Rows.Count - 1
            Dim list As New ListItem(dtStatus.Rows(i)("status_Name"), dtStatus.Rows(i)("status_id"))
            list.Attributes.Add("style", "background-color:" & dtStatus.Rows(i)("status_Color") & ";color:" & dtStatus.Rows(i)("status_FntColor"))
            sStatus.Items.Add(list)
        Next
        '''cmdSelect = New SqlClient.SqlCommand("SELECT status_id,status_Name,status_Color from Status_ForBizpower ORDER BY status_Name", con)
        ''''" ORDER BY Category_Order, Tour_Order", con) 'old version
        '''cmdSelect.CommandType = CommandType.Text
        ''con.Open()
        ''Dim ds As New DataSet
        ''Dim myData As SqlClient.SqlDataAdapter
        ''myData = New SqlClient.SqlDataAdapter("SELECT status_id,status_Name,status_Color from Status_ForBizpower ORDER BY status_Name ", con)
        ''        myData.Fill(ds, "AllTables")
        ''drpStatus.DataSource = ds
        ''drpStatus.DataSource = ds.Tables(0)
        ''drpStatus.DataTextField = ds.Tables(0).Columns("status_Name").ColumnName.ToString()
        ''drpStatus.DataValueField = ds.Tables(0).Columns("status_id").ColumnName.ToString()
        ''drpStatus.DataBind()
        ''Dim i As Integer
        ''For i = 0 To drpStatus.Items.Count - 1
        ''    drpStatus.Items(i).Attributes.Add("style", "background-color:" + ds.Tables(0).Rows(i)("status_Color").ToString())
        ''Next

        ''con.Close()
        ''drpStatus.Items.Insert(0, New System.Web.UI.WebControls.ListItem("--בחר סטטוס--", 0))
        ''drpStatus.Items(0).Attributes.Add("style", "background-color:#ff0000")

        ''If Not IsNothing(drpStatus.Items.FindByValue(StatusId)) Then
        ''    drpStatus.Items.FindByValue(StatusId).Selected = True
        ''End If

    End Sub
End Class
