Public Class logs_report
    Inherits System.Web.UI.Page
    Protected userId, orgID, strLocal, user_name, org_name, perSize As String
    Protected cmbChangeType, cmbChangeTable, cmbChangeTableSales As HtmlSelect
    Dim con As New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Public dir_var, align_var, dir_obj_var, lang_id As String
    Public arrTitles(1), arrButtons(1)
    Protected func As New bizpower.cfunc
    Public date_start, date_startSales, date_startNewsL, date_endNewsL, date_end, date_endSales, ChangeType, ChangeTable As String
      

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here     
        strLocal = "http://" & Trim(Request.ServerVariables("SERVER_NAME"))
        If Len(Trim(Request.ServerVariables("SERVER_PORT"))) > 0 Then
            strLocal = strLocal & ":" & Trim(Request.ServerVariables("SERVER_PORT"))
        End If
        strLocal = strLocal & Application("VirDir") & "/"
        If IsNothing(Request.Cookies("bizpegasus")) Then
            '   Response.Write("1=" & Len(Request.Cookies("bizpegasus")))
            Response.Redirect(strLocal)
        End If
        If Not IsNumeric(Request.Cookies("bizpegasus")("UserId")) Then
            Response.Redirect(strLocal)
        End If
        userId = Trim(Request.Cookies("bizpegasus")("UserId"))
        orgID = Trim(Request.Cookies("bizpegasus")("orgId"))
        perSize = Trim(Request.Cookies("bizpegasus")("perSize"))
        lang_id = Trim(Request.Cookies("bizpegasus")("LANGID"))
        org_name = Request.Cookies("bizpegasus")("ORGNAME")
        org_name = HttpContext.Current.Server.UrlDecode(org_name)
        user_name = Trim(Request.Cookies("bizpegasus")("UserName"))
        user_name = HttpContext.Current.Server.UrlDecode(user_name)

        If IsNumeric(lang_id) = False Or lang_id = "" Then
            lang_id = "1"
        End If
        If lang_id = "2" Then
            dir_var = "rtl" : align_var = "left" : dir_obj_var = "ltr"
        Else
            dir_var = "ltr" : align_var = "right" : dir_obj_var = "rtl"
        End If
        If date_start = "" Then
            date_start = "01" & "/" & Month(Now()) & "/" & Year(Now())
        End If
        If date_startSales = "" Then
            date_startSales = "01" & "/" & Month(Now()) & "/" & Year(Now())
        End If
        If date_startNewsL = "" Then
            date_startNewsL = "01" & "/" & Month(Now()) & "/" & Year(Now())
        End If
        If date_end = "" Then
            date_end = DateTime.DaysInMonth(Year(Now()), Month(Now())) & "/" & Month(Now()) & "/" & Year(Now())
        End If
        If date_endSales = "" Then
            date_endSales = DateTime.DaysInMonth(Year(Now()), Month(Now())) & "/" & Month(Now()) & "/" & Year(Now())
        End If
        If date_endNewsL = "" Then
            date_endNewsL = DateTime.DaysInMonth(Year(Now()), Month(Now())) & "/" & Month(Now()) & "/" & Year(Now())
        End If


        If Page.IsPostBack = False Then

            Dim cmdSelect As New SqlClient.SqlCommand("SELECT DISTINCT Change_Type FROM Changes", con)
            cmdSelect.CommandType = CommandType.Text
            con.Open()
            cmbChangeType.DataSource = cmdSelect.ExecuteReader()
            cmbChangeType.DataValueField = "Change_Type" : cmbChangeTable.DataTextField = "Change_Type"
            cmbChangeType.DataBind()
            con.Close()
            cmbChangeType.Items.Insert(0, New ListItem("", ""))

            cmdSelect = New SqlClient.SqlCommand("SELECT DISTINCT Change_Table FROM Changes", con)
            cmdSelect.CommandType = CommandType.Text
            con.Open()
            cmbChangeTable.DataSource = cmdSelect.ExecuteReader()
            cmbChangeTable.DataValueField = "Change_Table" : cmbChangeTable.DataTextField = "Change_Table"
            cmbChangeTable.DataBind()
            con.Close()
            cmbChangeTable.Items.Insert(0, New ListItem("", ""))

        End If

    End Sub

   

End Class